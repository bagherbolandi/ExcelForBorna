"""
renderer.py — writes a WorkbookModel to a real .xlsx via openpyxl
==================================================================
Includes: Excel Tables + Structured References, Data Validation dropdowns,
number formats, freeze panes, auto-filter, conditional formatting (flag-style),
column widths, and a light professional style.

Every formula is copied verbatim from the model (which the formula engine has
already verified), so file and verification stay in lockstep.
"""
from __future__ import annotations

import datetime as dt

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule, FormulaRule

import seed
from model import WorkbookModel

# --------------------------------------------------------------------------- #
# Palette
# --------------------------------------------------------------------------- #
INK = "1F2937"
ACCENT = "0F4C81"        # deep steel blue
ACCENT_LIGHT = "DCE9F7"
HEADER_FILL = "0F4C81"
HEADER_FONT = "FFFFFF"
NOTE_FILL = "F7F3E8"
NOTE_BORDER = Side(style="thin", color="D8CFB0")
WARN_AMBER = "B45900"
OK_GREEN = "1E7B3C"
MISS_RED = "9C0006"
STRIPE = "F3F7FB"

thin = Side(style="thin", color="CBD5E1")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

FONT_DEFAULT = Font(name="Calibri", size=11, color=INK)
FONT_TITLE = Font(name="Calibri", size=14, bold=True, color=ACCENT)
FONT_HEADER = Font(name="Calibri", size=11, bold=True, color=HEADER_FONT)
FONT_NOTE = Font(name="Calibri", size=9, italic=True, color="6B7280")

FILL_HEADER = PatternFill("solid", fgColor=HEADER_FILL)
FILL_NOTE = PatternFill("solid", fgColor=NOTE_FILL)
FILL_STRIPE = PatternFill("solid", fgColor=STRIPE)


# --------------------------------------------------------------------------- #
# Number/date formats per header keyword
# --------------------------------------------------------------------------- #
def fmt_for(header: str):
    h = header.lower()
    if any(k in h for k in ("setup_min", "std_min", "man_min")):
        return "#,##0"
    if any(k in h for k in ("date", "_d", "valid_", "create", "modify", "end_", "start_", "raised")):
        if "days" in h or "lead" in h:
            return "0"
        return "yyyy-mm-dd"
    if any(k in h for k in ("%", "percent", "scrap", "margin", "progress", "advance", "on_delivery", "sg&a")):
        return "0.0%"
    if any(k in h for k in ("price", "cost", "amount", "budget", "margin_irr", "value", "landed", "unit_", "freight",
                             "insurance", "customs", "handling", "other", "adjustment", "salary", "limit", "fx",
                             "total", "variance")):
        return "#,##0"
    if any(k in h for k in ("qty", "quantity", "moq", "days", "hours", "level", "sequence", "score",
                              "approval_", "progress", "lead_", "credit_", "line_no", "validity")):
        return "#,##0.##"
    if any(k in h for k in ("rate", "factor", "stock", "reserved", "qty_per", "scrap", "on_hand")):
        return "#,##0.00##"
    return "General"


# --------------------------------------------------------------------------- #
# Data validation dropdowns
# --------------------------------------------------------------------------- #
COLUMN_LISTS = {
    "Status_ID": [s[1] for s in seed.STATUSES],
    "Stage_Status": ["Completed", "In Progress", "Pending", "Not Started", "Blocked"],
    "Status": [s[1] for s in seed.STATUSES],
    "Decision": ["APPROVED", "APPROVED WITH CONDITION", "RETURN FOR REVISION", "REJECTED"],
    "Stage": ["Sales", "Engineering", "Planning", "Procurement", "Finance", "Management"],
    "Activity": ["Engineering", "BOM Release", "Procurement", "Supplier Confirmation", "Material Arrival",
                 "Production Planning", "Production", "QC", "Packing", "Delivery"],
    "Material_Type": ["Raw Material", "Purchased Part", "Assembly", "Sub Assembly", "Packaging", "Consumable"],
    "Make_Buy": ["Make", "Buy"],
    "Make_Buy_Default": ["Make", "Buy"],
    "Request_Type": ["New", "Repeat", "Amendment"],
    "Confidentiality_Level": ["Public", "Internal", "Confidential", "Strictly Confidential"],
    "Priority": ["Low", "Medium", "High", "Critical"],
    "UoM": ["pcs", "kg", "m", "m2", "liter", "set", "ton"],
    "Base_UoM": ["pcs", "kg", "m", "m2", "liter", "set", "ton"],
    "Currency": ["IRR", "USD", "EUR"],
    "Incoterm": ["EXW", "FCA", "CPT", "CIP", "DAP", "DDP"],
    "Incoterm_Default": ["EXW", "FCA", "CPT", "CIP", "DAP", "DDP"],
    "Payment_Risk_Level": ["Low", "Medium", "High"],
    "Price_Confirm_Status": ["Quoted", "Confirmed", "Revoked"],
    "Price_Status": ["Approved", "Repriced", "Expired"],
    "Issue_Status": ["Open", "In Progress", "Closed"],
    "Severity": ["Low", "Medium", "High", "Critical"],
    "Customer_Type": ["Corporate", "Government", "Distributor", "Retail"],
    "Tax_Status": ["VAT", "Exempt", "N/A"],
    # v2.0
    "Origin": ["Domestic", "Imported"],
    "QC_Status": ["Accepted", "Rejected", "Quarantine"],
    "Gantt_Updated_Flag": ["Yes", "No"],
    "Feasibility_Updated_Flag": ["Yes", "No"],
    "Risk_Level": ["Low", "Medium", "High"],
}


def _attach_validations(ws, sh, name):
    """Attach list data validation to known columns (header row = sh row 2)."""
    hdr_row = None
    for tbl in sh.tables:
        hdr_row = tbl.header_row
        break
    if hdr_row is None:
        return
    hdr_to_col = {}
    for c, cell in sh.rows.get(hdr_row, {}).items():
        if cell.value:
            hdr_to_col[str(cell.value)] = c
    if not hdr_to_col:
        return
    last = max(sh.max_row(), hdr_row + 1)
    for hdr, col in hdr_to_col.items():
        vals = COLUMN_LISTS.get(hdr)
        if not vals:
            continue
        formula1 = '"' + ",".join(str(v) for v in vals) + '"'
        if len(formula1) > 255:
            continue
        dv = DataValidation(type="list", formula1=formula1, allow_blank=True,
                            showDropDown=False, showErrorMessage=True)
        dv.error = "مقدار خارج از فهرست مجاز است"
        dv.errorTitle = "Data Validation"
        dv.prompt = f"انتخاب از فهرست مجاز برای {hdr}"
        dv.promptTitle = hdr
        ws.add_data_validation(dv)
        letter = get_column_letter(col)
        dv.add(f"{letter}{hdr_row+1}:{letter}{last}")



# --------------------------------------------------------------------------- #
# Main render
# --------------------------------------------------------------------------- #
def render(model: WorkbookModel, out_path: str):
    import renderer_v2
    wb = Workbook()
    wb.remove(wb.active)
    dashboard_chart_args = None

    for name in model.order:
        sh = model.sheets[name]
        ws = wb.create_sheet(title=name[:31])
        ws.sheet_view.rightToLeft = True   # RTL — Persian UI

        # write values/formulas
        for r in sh.rows:
            for c, cell in sh.rows[r].items():
                if cell.formula:
                    ws.cell(row=r, column=c, value=cell.formula)
                else:
                    ws.cell(row=r, column=c, value=cell.value)

        # ---- styling (only for <= 400 rows to stay fast) ----
        if sh.max_row() <= 500:
            _style(ws, sh, name)

        # ---- v2.0 gate data validations FIRST (sequencing locks take
        # precedence where ranges overlap list dropdowns) ----
        try:
            renderer_v2.attach_gate_validations(ws, name, sh)
        except Exception as e:
            print(f"[warn] gate DV on {name}: {e}")

        # ---- data validation dropdowns ----
        try:
            _attach_validations(ws, sh, name)
        except Exception as e:
            print(f"[warn] DV on {name}: {e}")

        # ---- v2.0 Gantt bars ----
        if name == "Gantt":
            n_rows = len(seed.SCHEDULE) + 1
            renderer_v2.finish_gantt(ws, n_rows, first_period_col=10, n_periods=26)

        # ---- v2.0 Dashboard charts (deferred until all sheets exist) ----
        if name == "Dashboard":
            marker = None
            cell = sh.get(1, 26)
            if cell and isinstance(cell.value, str) and cell.value.startswith("CHART1_ROWS="):
                marker = cell.value.split("=", 1)[1]
            if marker:
                cf, cl = (int(x) for x in marker.split(":"))
                dashboard_chart_args = (ws, cf, cl)

        # ---- tables ----
        for t in sh.tables:
            last_col = get_column_letter(t.n_cols)
            ref = f"A{t.header_row}:{last_col}{t.last_data_row}"
            openpyxl_table = Table(displayName=t.name, ref=ref)
            style = TableStyleInfo(
                name="TableStyleMedium2", showFirstColumn=False,
                showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            openpyxl_table.tableStyleInfo = style
            ws.add_table(openpyxl_table)

        # freeze
        ws.freeze_panes = sh.freeze

    if dashboard_chart_args is not None:
        ws, cf, cl = dashboard_chart_args
        renderer_v2.add_dashboard_charts(ws, wb, cf, cl, len(seed.PROJECTS))

    # ---- v2.0 access control: protect all sheets, then save & inject
    # password-protected edit ranges (Protected Ranges) per role ----
    renderer_v2.protect_all(wb)
    wb.save(out_path)
    renderer_v2.inject_protected_ranges(out_path, model.order)
    return out_path


def _style(ws, sh, name):
    # title row
    title_cell = ws.cell(row=1, column=1)
    if title_cell.value:
        title_cell.font = FONT_TITLE
    # note cell (row 1 col 2)
    note_cell = ws.cell(row=1, column=2)
    if note_cell.value:
        note_cell.font = FONT_NOTE

    has_table = bool(sh.tables)
    header_row = 2 if has_table else None

    max_r = sh.max_row()
    max_c = sh.max_col()

    # column widths
    if max_c <= 40:
        for c in range(1, max_c + 1):
            letter = get_column_letter(c)
            hdr = ws.cell(row=header_row, column=c).value if header_row else None
            width = _width_for(hdr)
            ws.column_dimensions[letter].width = width

    # style header row
    if header_row:
        for c in range(1, max_c + 1):
            cell = ws.cell(row=header_row, column=c)
            if cell.value is not None:
                cell.font = FONT_HEADER
                cell.fill = FILL_HEADER
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # style data rows + number formats
    for r in range(3, max_r + 1):
        for c in range(1, max_c + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            if header_row:
                hdr = ws.cell(row=header_row, column=c).value
                if hdr:
                    f = fmt_for(str(hdr))
                    if f != "General":
                        cell.number_format = f
            if header_row and r % 2 == 0:
                cell.fill = FILL_STRIPE

    # conditional formatting on key result columns
    if header_row:
        _conditional(ws, sh, header_row, max_r)


def _width_for(hdr):
    if not hdr:
        return 12
    h = str(hdr)
    # generous estimate for headers
    return min(34, max(11, int(len(h) * 1.15) + 2))


def _conditional(ws, sh, header_row, max_r):
    """Flag critical computed columns with cell-is rules (green/amber/red)."""
    def col_of(substr):
        for c in range(1, sh.max_col() + 1):
            v = ws.cell(row=header_row, column=c).value
            if v and substr.lower() in str(v).lower():
                return get_column_letter(c)
        return None

    rules = [
        ("Ready_Flag", '"READY"', OK_GREEN, '"INCOMPLETE"', MISS_RED),
        ("Formula_Result", '"OK"', OK_GREEN, '"ERROR"', MISS_RED),
        ("Expired_Flag", "1", MISS_RED, "0", OK_GREEN),
        ("Overdue_Flag", "1", MISS_RED, "0", OK_GREEN),
        ("Delay_Days", None, None, ">0", WARN_AMBER),
    ]
    for col_name, good, good_c, bad, bad_c in rules:
        col = col_of(col_name)
        if not col:
            continue
        rng = f"{col}3:{col}{max_r}"
        if good_c and good is not None:
            ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[good], fill=PatternFill("solid", fgColor="E6F4EA"), font=Font(color=OK_GREEN, bold=True)))
        if bad_c:
            op = "greaterThan" if bad.startswith(">") else "equal"
            formula = [bad] if op == "greaterThan" else [bad]
            ws.conditional_formatting.add(rng, CellIsRule(operator=op, formula=formula, fill=PatternFill("solid", fgColor="FDE7E9"), font=Font(color=MISS_RED, bold=True)))
