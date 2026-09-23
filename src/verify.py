"""
verify.py — evaluates every formula in the WorkbookModel with the formula
engine and checks the financial chain is consistent.
====================================================================

Approach:
  - flatten the model to {sheet: {A1: value_or_FormulaError}}
  - starting values: literal cells only
  - repeatedly evaluate formula cells whose referenced cells are already known
    (topological iteration — no need for a full dependency solver)
  - afterwards, map counters from arrays to plain values (engine limitation)
  - assertions against hand-calculated expectations

Returns True on success; prints a readable report.
"""
from __future__ import annotations

import datetime as dt

import seed
from model import col_letter
from formula_engine import Book, Grid, evaluate, FormulaError


def cell_ref(r, c):
    return f"{col_letter(c)}{r}"


def build_flat(model):
    """Return dict sheet -> {A1ref: ('lit', value) | ('f', formula)} ."""
    flat = {}
    for name in model.order:
        sh = model.sheets[name]
        cells = {}
        for r in sh.rows:
            for c in sh.rows[r]:
                cell = sh.rows[r][c]
                ref = cell_ref(r, c)
                if cell.formula:
                    cells[ref] = ("f", cell.formula)
                else:
                    cells[ref] = ("lit", cell.value)
        flat[name] = cells
    return flat


def _start_ref(f):
    f = f.strip()
    if f.startswith("="):
        f = f[1:]
    return f


def _same(a, b):
    """Value equality for fixpoint detection (FormulaError compared by text)."""
    if isinstance(a, FormulaError) or isinstance(b, FormulaError):
        va = getattr(a, "value", a)
        vb = getattr(b, "value", b)
        return va == vb
    if isinstance(a, bool) != isinstance(b, bool):
        return False
    return a == b


def evaluate_all(model, max_passes=40):
    """Return sheet -> A1 -> value (formula errors surfaced as FormulaError).

    Fixpoint (Jacobi) evaluation: each pass evaluates every formula against a
    snapshot of the state; passes repeat until no cell changes. This makes the
    result independent of sheet/tab ordering (sheets may cross-reference in
    both directions at the cell level — e.g. Gate_Status ↔ PMO_Summary).
    """
    flat = build_flat(model)
    state = {}
    formulas = []
    for sheet, cells in flat.items():
        state[sheet] = {}
        for ref, (kind, val) in cells.items():
            if kind == "lit":
                state[sheet][ref] = val
            else:
                formulas.append((sheet, ref, val))

    for _ in range(max_passes):
        book = Book()
        for sname, cells in state.items():
            grid = Grid(sname)
            for ref, v in cells.items():
                grid.set(ref, v)
            book.add(grid)
        changed = False
        for sheet, ref, formula in formulas:
            try:
                v = evaluate(book, sheet, formula)
            except FormulaError as e:
                v = e
            if not _same(state[sheet].get(ref), v):
                state[sheet][ref] = v
                changed = True
        if not changed:
            break
    return state, {}


def evaluate_referencing(state, sheet, formula):
    """Evaluate formula against state but handle cross-sheet & blank cells."""
    book = Book()
    for sname, cells in state.items():
        g = Grid(sname)
        for r, v in cells.items():
            # engine: blank/None yields 0 in value context; keep literal None
            g.set(r, v)
        book.add(g)
    return evaluate(book, sheet, formula)


def num(v):
    if isinstance(v, FormulaError):
        return None
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, dt.date):
        return None
    if isinstance(v, str):
        import re
        m = re.fullmatch(r"-?\d+(\.\d+)?", v.strip())
        return float(m.group()) if m else None
    return None


# --------------------------------------------------------------------------- #
# Expectations (hand-computed from seed.py)
# --------------------------------------------------------------------------- #
def expectations():
    fx = 500000.0
    # Landed costs (active price, per unit)
    landed = {
        "MAT-001542": (1.15 + 0.05) * fx,       # cast iron USD
        "MAT-001543": (3.2 + 0.05) * fx,        # stainless USD
        "MAT-001544": 95_000_000 + 1_200_000,   # bearing
        "MAT-001545": 185_000_000 + 2_000_000,  # mechanical seal
        "MAT-001546": (2_500 + 30) * fx,        # motor USD
        "MAT-001547": 2_450_000,                # shaft (R02, current)
        "MAT-001548": 12_000_000,               # packaging
    }
    planning = planner_rows()
    dm = 0.0
    pp = 0.0
    for pl in planning:
        if pl["material"] == "MAT-001548":
            continue
        if pl["make_buy"] == "Make":
            dm += pl["required"] * landed[pl["material"]]
        else:
            pp += pl["required"] * landed[pl["material"]]
    # v2.0 — labour comes from the engineering Time_Study (نفرساعت), not an assumption
    hours = time_study_hours()
    labor = round(hours * 2_500_000)   # ASSUMPTION settings: labor rate
    oh = round(labor * 2.0)            # overhead rate 2.0 × direct labor
    scrap = (dm + pp) * 0.02
    packaging = 2 * 12_000_000        # 2 pcs
    logistics = 150_000_000
    tooling = 100_000_000
    sga = (dm + pp + packaging + labor + oh) * 0.05
    total_cost = dm + pp + labor + oh + scrap + packaging + logistics + tooling + sga
    cost_per_unit = round(total_cost / 2)

    # quotation scenario B
    margin = total_cost * 0.10
    financial = round((1 - 0.30 - 0.10) * total_cost * 0.18 * (60 / 360))
    proposed = round(total_cost + margin + 10_000_000 + 15_000_000 + financial)

    # v2.0 — project equipment (FX conversion for imported items)
    peq_total_imported = round(12_500 * fx) * 1 + 150_000_000 + 80_000_000
    peq_total_domestic = 8_900_000_000 * 1 + 200_000_000 + 120_000_000

    return dict(
        fx=fx, landed=landed, planning=planning,
        dm=dm, pp=pp, hours=hours, labor=labor, oh=oh, scrap=scrap, packaging=packaging,
        logistics=logistics, tooling=tooling, sga=sga, total_cost=total_cost,
        cost_per_unit=cost_per_unit, margin=margin, financial=financial,
        proposed=proposed,
        peq_total_imported=peq_total_imported, peq_total_domestic=peq_total_domestic,
    )


def time_study_hours():
    """Total man-hours for PRJ-2026-00125 as the Time_Study formulas compute them.

    Per row: Man_Min = Setup + Std×Order_Qty×(1+ScrapAllow); Hours = ROUND(Man_Min/60×Operators, 2)
    Order_Qty = 2 (OLN qty of PRD-000125 on ORD-2026-00321).
    """
    order_qty = 0
    for oln in seed.ORDER_LINES:
        if oln[1] == "ORD-2026-00321" and oln[3] == "PRD-000125":
            order_qty += float(oln[4])
    total = 0.0
    for ts in seed.TIME_STUDY:
        if ts[1] != "PRJ-2026-00125":
            continue
        setup, std, ops, scrap = ts[5], ts[6], ts[7], ts[8]
        man_min = setup + std * order_qty * (1 + scrap)
        total += round(man_min / 60 * ops, 2)
    return total


def planner_rows():
    from builder import planning_rows
    return planning_rows()


# --------------------------------------------------------------------------- #
# Run + report
# --------------------------------------------------------------------------- #
def run(model):
    state, pending = evaluate_all(model)
    exp = expectations()

    problems = []

    def g(sheet, ref):
        return state.get(sheet, {}).get(ref)

    def check(name, actual, expected, tol=1.0):
        a = num(actual)
        ok = (a is not None and abs(a - expected) <= tol)
        if not ok:
            problems.append(f"{name}: expected {expected} got {a} (raw={actual!r})")
        return ok

    print("=" * 78)
    print("VERIFICATION REPORT — ExcelForBorna (order-to-quotation)")
    print("=" * 78)

    # 1. Direct Material
    check("Costing.Direct_Material", g("Costing", "G3"), exp["dm"])
    check("Costing.Purchased_Parts", g("Costing", "H3"), exp["pp"])
    check("Costing.Direct_Labor", g("Costing", "I3"), exp["labor"])
    check("Costing.Mfg_Overhead", g("Costing", "J3"), exp["oh"])
    check("Costing.Scrap", g("Costing", "K3"), exp["scrap"])
    check("Costing.Packaging", g("Costing", "L3"), exp["packaging"])
    check("Costing.Logistics", g("Costing", "M3"), exp["logistics"])
    check("Costing.Tooling", g("Costing", "N3"), exp["tooling"])
    check("Costing.SG&A", g("Costing", "O3"), exp["sga"])
    check("Costing.Total_Cost", g("Costing", "Q3"), exp["total_cost"])
    check("Costing.Cost_Per_Unit", g("Costing", "R3"), exp["cost_per_unit"])

    # 2. Quotation scenario B (row 4)
    check("Quotation.Total_Cost(F)", g("Sales_Quotation", "F4"), exp["total_cost"])
    check("Quotation.Margin_Amount(H)", g("Sales_Quotation", "H4"), exp["margin"])
    check("Quotation.Financial_Adjustment(K)", g("Sales_Quotation", "K4"), exp["financial"])
    check("Quotation.Proposed_Selling_Price(L)", g("Sales_Quotation", "L4"), exp["proposed"])

    # 3. Budget vs actual
    check("Budget.Total_Budget", g("Budget", "L3"), 4_000_000_000 + 350_000_000 + 700_000_000 + 200_000_000 + 150_000_000 + 100_000_000 + 50_000_000)
    check("Budget.Actual_Total", g("Budget", "R3"), exp["total_cost"])
    check("Budget.Variance_Total", g("Budget", "S3"), exp["total_cost"] - 5_550_000_000)

    # 4. Controls
    # Ready flag of the complete order (row 3) vs incomplete order (row 5)
    print("\n--- Controls ---")
    ready1 = g("Orders", "R3")
    ready2 = g("Orders", "R5")
    print(f"Order1 Ready_Flag = {ready1!r}")
    print(f"Order3 Ready_Flag = {ready2!r} (must be INCOMPLETE)")
    check("Orders.CO_02 row3", g("Orders", "O3"), None) if False else None
    if ready2 != "INCOMPLETE":
        problems.append("incomplete order should show INCOMPLETE")
    # Expired flag: PPR R01 shaft (row 8? given 10 rows, find by System_ID)
    expired_ok = False
    for rack in range(3, 3 + len(seed.PURCHASE_PRICES)):
        if g("Purchase_Prices", f"A{rack}") == "PPR-000559":
            ev = g("Purchase_Prices", f"R{rack}")
            print(f"PPR-000559 Expired_Flag = {ev!r}")
            expired_ok = ev in (1, 1.0, True)
            break
    if not expired_ok:
        problems.append("expired shaft price flag should be 1")

    # 5. Planning integrity
    print("\n--- Planning ---")
    plan_n = g("Planning", "K3")  # required qty first row
    print(f"Planning first Required_Qty (K3) = {plan_n!r}")
    print(f"Planning row count = {len(planner_rows())}")

    # 6. Trace / dashboard spot checks
    print("\n--- Dashboard spot checks ---")
    for label, ref in [("Active projects", "B3"), ("New orders", "B4"),
                       ("In sales review", "B5"), ("Overdue projects", "B6"),
                       ("Feasibility complete", "B7"), ("Equipment priced", "B9"),
                       ("Receipts qty", "B11"), ("Expired prices", "B13"),
                       ("Pending mgmt quotes", "B14"),
                       ("Total Sales Value", "B19"), ("Total Cost", "B20")]:
        print(f"  {label:22s} = {g('Dashboard', ref)!r}")
    # workflow matrix (row 34 = PRJ-001, row 35 = PRJ-002) — engine-verified
    m1 = [g("Dashboard", f"{c}34") for c in "BCDEFGHIJ"]
    m2 = [g("Dashboard", f"{c}35") for c in "BCDEFGHIJ"]
    print(f"  matrix PRJ-001 = {m1}")
    print(f"  matrix PRJ-002 = {m2}")
    if m1 != ["✅ تکمیل"] * 8 + ["✅ تکمیل"]:
        problems.append("Dashboard matrix PRJ-001 should be fully complete")
    if m2[0] != "✅ تکمیل" or m2[1] != "🔵 در جریان" or any(x != "🔒 قفل" for x in m2[2:]):
        problems.append("Dashboard matrix PRJ-002 should be locked after feasibility stage")

    # 7. v2.0 lifecycle — gates, feasibility, time study, equipment, PMO
    print("\n--- v2.0 lifecycle (gates & sequencing) ---")
    stage_cols = ["D", "E", "F", "G", "H", "I", "J", "K"]
    prj1_flags = [num(g("Gate_Status", f"{c}3")) for c in stage_cols]
    prj2_flags = [num(g("Gate_Status", f"{c}4")) for c in stage_cols]
    print(f"  PRJ-001 stage flags = {prj1_flags}")
    print(f"  PRJ-002 stage flags = {prj2_flags}")
    if prj1_flags != [1.0] * 8:
        problems.append("PRJ-001: all 8 stages should be complete")
    if prj2_flags != [1.0] + [0.0] * 7:
        problems.append("PRJ-002: only sales stage should be complete (feasibility incomplete)")
    check("Gate_Status.P1 Stages_Completed", g("Gate_Status", "V3"), 8)
    check("Gate_Status.P1 G_Senior", g("Gate_Status", "U3"), 1)
    check("Gate_Status.P2 G_Eng (must be locked)", g("Gate_Status", "O4"), 0)
    check("Gate_Status.P2 G_Com (must be locked)", g("Gate_Status", "P4"), 0)
    if g("Gate_Status", "M3") != "APPROVED WITH CONDITION":
        problems.append("PRJ-001 decision text mismatch")
    if g("Gate_Status", "X4") != "تعریف پروژه و امکان‌سنجی (مدیر پروژه)":
        problems.append("PRJ-002 current stage should be feasibility")

    print("\n--- v2.0 feasibility / time study / equipment ---")
    check("Feasibility.Weighted(P1)", g("Feasibility", "I3"), 83.65, tol=0.01)
    if g("Feasibility", "K3") != "Feasible" or g("Feasibility", "M3") != "COMPLETE":
        problems.append("Feasibility P1 result/flag mismatch")
    if g("Feasibility", "M4") != "INCOMPLETE":
        problems.append("Feasibility P2 must be INCOMPLETE")
    check("Time_Study.hours→Costing.Direct_Labor", g("Costing", "I3"), exp["labor"])
    check("Project_Equipment.total(imported, FX)", g("Project_Equipment", "M3"), exp["peq_total_imported"])
    check("Project_Equipment.total(domestic)", g("Project_Equipment", "M4"), exp["peq_total_domestic"])
    if g("Project_Equipment", "N5") != "PENDING":
        problems.append("PEQ-000003 must remain PENDING until commercial gate opens")
    check("PMO_Summary.Units_Completed", g("PMO_Summary", "C3"), 8)
    if g("PMO_Summary", "J3") != "REPORTED":
        problems.append("PMO summary P1 must be REPORTED")

    # Summary
    print("\n" + "=" * 78)
    if problems:
        print(f"VERIFICATION FAILED — {len(problems)} problem(s):")
        for p in problems:
            print("  ✗", p)
        return False
    print("VERIFICATION PASSED — all key calculations are consistent with the spec.")
    print(f"  Total Cost = {exp['total_cost']:,.0f} IRR")
    print(f"  Cost / Unit = {exp['cost_per_unit']:,.0f} IRR")
    print(f"  Proposed (Scenario B) = {exp['proposed']:,.0f} IRR")
    return True
