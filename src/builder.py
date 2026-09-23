"""
builder.py — builds the WorkbookModel from seed data
=====================================================
Every formula here is written as a real A1 Excel formula. The same model is
later handed to the formula engine for verification and to the renderer for
the final .xlsx. Deterministic by construction.

Design decisions (documented):
  - "Financial cost of credit" is computed at the *quotation* level
    (Total_Cost × credit portion × annual rate × days/360) to avoid a
    circular reference with Costing. Costing = goods cost of the product.
  - Multi-currency: the electric motor is quoted in USD; its Landed Cost is
    converted to IRR via the Fx rate in Settings (auditable, not hard-coded).
"""
from __future__ import annotations

import datetime as dt

import seed
from model import WorkbookModel, col_letter

F = lambda s: ("F", s)


def D(s):
    if s is None:
        return None
    if isinstance(s, dt.date):
        return s
    if isinstance(s, str):
        s = s.strip()
        if s == "":
            return None
    return dt.date.fromisoformat(str(s))


def setting_num(key):
    for k, v, _d, _a, _desc in seed.SETTINGS:
        if k == key:
            return float(v)
    raise KeyError(key)


class TB:
    """Table writer: title row 1, headers row 2, data from row 3."""
    def __init__(self, wb, sheet, name, headers, note=None):
        self.wb = wb
        self.ws = wb.sheet(sheet)
        self.name = name
        self.headers = headers
        self.col = {h: i + 1 for i, h in enumerate(headers)}
        self.row = 2
        self.ws.set(1, 1, value=sheet)
        if note:
            self.ws.set(1, 2, value=note)
        for c, h in enumerate(headers, start=1):
            self.ws.set(2, c, value=h)
        self.data_start = 3

    def add(self, cells):
        self.row += 1
        r = self.row
        if isinstance(cells, dict):
            for name, v in cells.items():
                self._put(r, self.col[name], v)
        else:
            for c, v in enumerate(cells, start=1):
                self._put(r, c, v)
        return r

    def _put(self, r, c, v):
        if isinstance(v, tuple) and v and v[0] == "F":
            self.ws.set(r, c, formula=v[1])
        else:
            self.ws.set(r, c, value=v)

    def finish(self):
        t = self.ws.add_table(self.name, header_row=2, first_data_row=3)
        t.n_cols = len(self.headers)
        t.last_data_row = max(3, self.row)
        self.ws.freeze = "A3"
        return t


# --------------------------------------------------------------------------- #
# Settings / Lists
# --------------------------------------------------------------------------- #
def build_settings(wb):
    headers = ["Key", "Value", "Data_Type", "ASSUMPTION", "Description"]
    t = TB(wb, "Settings", "tblSettings", headers,
           note="همه نرخ‌ها/ضرایب سازمانی اینجا ویرایش می‌شوند؛ پیش‌فرض‌ها ASSUMPTION هستند (هیچ عددی داخل فرمول‌ها Hard-code نشده).")
    addr = {}
    for key, value, dtype, asmp, desc in seed.SETTINGS:
        r = t.add([key, value, dtype, "YES" if asmp else "NO", desc])
        addr[key] = f"Settings!$B${r}"
    t.finish()
    return addr


def build_lists(wb):
    headers = ["List_Name", "Value"]
    t = TB(wb, "Lists", "tblLists", headers, note="فهرست‌های Data Validation (داده‌های مجاز).")
    ranges = {}
    prev = None
    start = None
    prev_r = None
    for name, values in seed.LISTS:
        for v in [x.strip() for x in values.split(",") if x.strip()]:
            r = t.add([name, v])
            if name != prev:
                if prev is not None and start is not None:
                    ranges[prev] = (start, prev_r)
                start = r
                prev = name
            prev_r = r
    if prev is not None and start is not None:
        ranges[prev] = (start, prev_r)
    t.finish()
    return ranges


def build_reference(wb, sheet, tname, headers, rows, note=None):
    t = TB(wb, sheet, tname, headers, note=note)
    for row in rows:
        t.add(row)
    return t.finish()


# --------------------------------------------------------------------------- #
# Planning compute (single source for Planning + Costing material lines)
# --------------------------------------------------------------------------- #
def planning_rows():
    """Return a list of dicts describing each active BOM leaf for the demo project."""
    order_qty = {}   # (project, product) -> qty
    project_of_order = {}
    for o in seed.ORDERS:
        project_of_order[o[0]] = o[2]
    for oln in seed.ORDER_LINES:
        prj = project_of_order.get(oln[1])
        if prj:
            order_qty[(prj, oln[3])] = order_qty.get((prj, oln[3]), 0) + float(oln[4])

    # map product_revision -> product, product -> product_revision (current)
    prd_of_prv = {pr[0]: pr[1] for pr in seed.PRODUCT_REVISIONS}
    prv_of_prd = {}
    for pr in seed.PRODUCT_REVISIONS:
        prv_of_prd.setdefault(pr[1], pr[0])

    active_bom = {}   # product -> active BOM revision id (via product_revision)
    for bh in seed.BOM_HEADERS:
        if bh[9] == 1:
            prd = prd_of_prv.get(bh[2])
            if prd:
                active_bom.setdefault(prd, bh[0])

    rows = []
    for (prj, prod), qty in order_qty.items():
        if prj != "PRJ-2026-00125":
            continue
        brel = active_bom.get(prod)
        if not brel:
            continue
        for bd in seed.BOM_DETAILS:
            if bd[1] == brel and bd[3] == 1:  # level-1 leaves of the active revision
                required = bd[6] * (1 + bd[8]) * qty
                rows.append(dict(
                    project=prj, product=prod, bom_rev=brel, bom_line=bd[0],
                    material=bd[5], make_buy=bd[9], bom_qty=bd[6], uom=bd[7],
                    scrap=bd[8], supplier=bd[10] if len(bd) > 10 else "",
                    order_qty=qty, required=required,
                ))
    return rows


# --------------------------------------------------------------------------- #
# Main build
# --------------------------------------------------------------------------- #
def build_model():
    wb = WorkbookModel()
    SET = build_settings(wb)
    LISTS = build_lists(wb)

    build_reference(wb, "Statuses", "tblStatuses",
                    ["Status_ID", "Status_Name", "Status_Group", "Sequence"], seed.STATUSES,
                    note="وضعیت سراسری (Universal Status) — همه جداول از این فهرست استفاده می‌کنند.")
    build_reference(wb, "Departments", "tblDepartments",
                    ["Department_ID", "Department_Name"], seed.DEPARTMENTS)
    build_reference(wb, "Roles", "tblRoles",
                    ["Role_ID", "Role_Name"], seed.ROLES)
    build_reference(wb, "Users", "tblUsers",
                    ["User_ID", "Name", "Department_ID", "Role_ID", "Email", "Approval_Level", "Is_Active"],
                    seed.USERS, note="کاربران سازمانی؛ Approval_Level پایه ماتریس تأیید است.")
    build_reference(wb, "Approval_Matrix", "tblApprovalMatrix",
                    ["Entity", "Gate", "Approver_Role", "Min_Approval_Level", "Responsible", "Reviewer", "Approver", "Informed"],
                    seed.APPROVAL_MATRIX)
    build_reference(wb, "RACI", "tblRACI",
                    ["Activity", "Sales", "Engineering", "Planning", "Procurement", "Finance", "PMO", "Management"],
                    seed.RACI)
    build_reference(wb, "Customers", "tblCustomers",
                    ["System_ID", "Customer_Code", "Customer_Name", "Legal_Name", "Customer_Type",
                     "Country", "City", "Delivery_Location", "Payment_Condition", "Credit_Limit",
                     "Currency", "Tax_Status", "Is_Active", "Sales_Owner", "Notes"],
                    seed.CUSTOMERS)
    build_reference(wb, "Products", "tblProducts",
                    ["System_ID", "Product_Code", "Product_Name", "Product_Type", "Base_UoM", "Category", "Is_Active"],
                    seed.PRODUCTS)
    build_reference(wb, "Product_Revisions", "tblProductRevisions",
                    ["System_ID", "Product_ID", "Revision_No", "Drawing_No", "Drawing_Revision",
                     "BOM_Policy", "Approved_By", "Approve_Date", "Change_Reason"],
                    [[r[0], r[1], r[2], r[3], r[4], r[5], r[6], D(r[7]), r[8]] for r in seed.PRODUCT_REVISIONS])
    build_reference(wb, "Materials", "tblMaterials",
                    ["System_ID", "Material_Code", "Material_Description", "Material_Type", "Specification",
                     "UoM", "Make_Buy_Default", "Preferred_Supplier", "Is_Active"],
                    seed.MATERIALS)
    build_reference(wb, "Suppliers", "tblSuppliers",
                    ["System_ID", "Supplier_Code", "Supplier_Name", "Legal_Name", "Country", "Currency",
                     "Payment_Terms_Default", "Incoterm_Default", "Rating", "Approval_Status", "Contact", "Is_Active"],
                    seed.SUPPLIERS)
    build_reference(wb, "Inventory", "tblInventory",
                    ["System_ID", "Material_ID", "Stock_On_Hand", "Reserved"],
                    seed.INVENTORY, note="موجودی/رزرو برای محاسبه کسری در Planning.")

    _build_orders(wb)
    _build_order_lines(wb)
    _build_projects(wb)
    _build_project_statuses(wb)
    _build_engineering(wb)
    _build_bom(wb)
    _build_quotes(wb)
    _build_procurement(wb)
    _build_purchase_prices(wb)
    _build_planning(wb, SET)
    _build_schedule(wb)
    _build_costlining(wb, SET)
    _build_budget(wb)
    _build_payment_terms(wb, SET)
    _build_sales_quotation(wb, SET)
    _build_quotation_versions(wb)
    _build_management_approval(wb)
    _build_workflow_history(wb)
    _build_change_log(wb)
    _build_issues(wb)
    _build_error_checks(wb)
    _build_test_results(wb)
    _build_trace(wb)
    _build_report(wb)
    _build_named_ranges_doc(wb, SET)

    # v2.0 — full lifecycle extensions (feasibility, equipment, time study,
    # receipts, PMO summary, gate status, Gantt, access control) + tab reorder
    import builder_v2
    builder_v2.extend(wb, SET)
    builder_v2.build_workflow_v2(wb)
    builder_v2.build_dashboard_v2(wb)
    builder_v2.reorder(wb)

    return wb, SET


# --------------------------------------------------------------------------- #
SECTION = {}


def _build_orders(wb):
    headers = ["System_ID", "Customer_ID", "Order_Date", "Requested_Delivery_Date", "Request_Type",
               "Total_Qty", "Delivery_Terms", "Payment_Terms", "Validity_Days",
               "Confidentiality_Level", "Sales_Responsible", "Doc_Ref", "Status_ID",
               "CO_01_Customer_Valid", "CO_02_Qty_Defined", "CO_03_Delivery_Defined", "CO_04_Payment_Defined",
               "Ready_Flag", "Duplicate_Flag", "Create_Date", "Created_By"]
    t = TB(wb, "Orders", "tblOrders", headers,
           note="کنترل اولیه سفارش (CO_01..CO_04)؛ تا Ready_Flag=READY نباشد ورود به مهندسی مجاز نیست.")
    for o in seed.ORDERS:
        r = t.row + 1
        t.add([
            o[0], o[1], D(o[3]), D(o[4]), o[5], None, o[6], o[7], o[8], o[9], o[10], o[11], o[12],
            F(f'=IF(TRIM($B{r})<>"","OK","MISSING")'),
            F(f'=IF(SUMIF(Order_Lines!$B$3:$B$20,$A{r},Order_Lines!$E$3:$E$20)>0,"OK","MISSING")'),
            F(f'=IF(TRIM($D{r})<>"","OK","MISSING")'),
            F(f'=IF(TRIM($H{r})<>"","OK","MISSING")'),
            F(f'=IF(AND($N{r}="OK",$O{r}="OK",$P{r}="OK",$Q{r}="OK"),"READY","INCOMPLETE")'),
            F(f'=IF(COUNTIF($A$3:$A$1000,$A{r})>1,"DUP","OK")'),
            D("2026-09-20"), "USR-02",
        ])
        t.ws.set(r, 6, formula=f'=SUMIF(Order_Lines!$B$3:$B$20,$A{r},Order_Lines!$E$3:$E$20)')
    t.finish()


def _build_order_lines(wb):
    headers = ["System_ID", "Order_ID", "Line_No", "Product_ID", "Quantity", "UoM", "Latest_FLAG",
               "Create_Date", "Created_By"]
    t = TB(wb, "Order_Lines", "tblOrderLines", headers,
           note="هر سفارش می‌تواند چند ردیف محصول داشته باشد (Order → Order_Line → Product).")
    for ln in seed.ORDER_LINES:
        t.add([ln[0], ln[1], ln[2], ln[3], ln[4], ln[5], ln[6], D("2026-09-20"), "USR-02"])
    t.finish()


def _build_projects(wb):
    headers = ["System_ID", "Project_Code", "Order_ID", "Customer_ID", "Product_ID", "Project_Name",
               "Project_Manager", "Start_Date", "Due_Date", "Priority", "Status_ID", "Progress_%",
               "Sponsor", "Team_List", "Overdue_Flag", "Create_Date", "Created_By"]
    t = TB(wb, "Projects", "tblProjects", headers)
    for p in seed.PROJECTS:
        r = t.row + 1
        t.add([p[0], p[1], p[2], p[3], p[4], p[5], p[6], D(p[7]), D(p[8]), p[9], p[10],
               p[11] / 100.0, p[12], p[13],
               F(f'=IF(ISBLANK($I{r}),0,IF(AND($I{r}<TODAY(),$K{r}<>"Closed"),1,0))'),
               D("2026-09-21"), "USR-03"])
    t.finish()


def _build_project_statuses(wb):
    headers = ["System_ID", "Project_ID", "Stage", "Stage_Status", "Start_Date", "End_Date", "Remark"]
    t = TB(wb, "Project_Statuses", "tblProjectStatuses", headers)
    for s in seed.PROJECT_STATUSES:
        t.add([s[0], s[1], s[2], s[3], D(s[4]), D(s[5]), s[6]])
    t.finish()


def _build_engineering(wb):
    headers = ["System_ID", "Project_ID", "Product_ID", "Product_Revision_ID", "Drawing_No",
               "Drawing_Revision", "BOM_Revision_ID", "Status_ID", "Process", "Machine", "Tooling",
               "Packaging", "Engineering_Notes", "Submitted_By", "Approved_By", "Create_Date", "Created_By"]
    t = TB(wb, "Engineering", "tblEngineering", headers)
    for e in seed.ENGINEERING:
        t.add([e[0], e[1], e[2], e[3], e[4], e[5], e[6], e[7],
               "Machining + Assembly", "CNC-5Axis", "Fixture-250", "ISPM-15",
               "مطابق نقشه DRW", "USR-04", "USR-04", D("2026-09-25"), "USR-04"])
    t.finish()


def _build_bom(wb):
    hh = ["System_ID", "BOM_ID", "Product_Revision_ID", "Revision_No", "Rev_Date", "Change_Reason",
          "Changed_By", "Approved_By", "Status_ID", "Is_Active", "Create_Date", "Created_By"]
    t = TB(wb, "BOM_Header", "tblBOMHeader", hh,
           note="نسخه‌بندی BOM؛ نسخه قبلی هرگز Overwrite نمی‌شود (Is_Active=0 / Status=Superseded).")
    for h in seed.BOM_HEADERS:
        t.add([h[0], h[1], h[2], h[3], D(h[4]), h[5], h[6], h[7], h[8], h[9], D(h[4]), h[6]])
    t.finish()

    dh = ["System_ID", "BOM_Revision_ID", "Parent_BOM_Line_ID", "Level_No", "Sequence_No",
          "Material_ID", "Qty_Per", "UoM", "Scrap_%", "Net_Qty", "Gross_Qty", "Make_Buy",
          "Approved_Supplier", "Create_Date", "Created_By"]
    dt = TB(wb, "BOM_Detail", "tblBOMDetail", dh,
            note="BOM چندسطحی (Parent_BOM_Line_ID / Level_No / Sequence_No). Gross_Qty = Qty_Per × (1 + Scrap%).")
    for b in seed.BOM_DETAILS:
        r = dt.row + 1
        dt.add([b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7], b[8],
                    F(f"=$G{r}"),
                    F(f"=ROUND($G{r}*(1+$I{r}),3)"),
                    b[9], b[10] if len(b) > 10 else "", D("2026-09-28"), "USR-04"])
    dt.finish()


def _build_quotes(wb):
    qh = ["System_ID", "Project_ID", "Material_ID", "Qty_Required", "Quoted_Date", "Valid_Until",
          "Status_ID", "Selected_Line_ID", "Create_Date", "Created_By"]
    t = TB(wb, "Supplier_Quotes", "tblSupplierQuotes", qh,
           note="استعلام به تفکیک ماده؛ انتخاب تأمین‌کننده در ردیف‌ها (نه فقط کمترین قیمت) توسط مسئول مجاز.")
    for q in seed.SUPPLIER_QUOTES:
        t.add([q[0], q[1], q[2], q[3], D(q[4]), D(q[5]), q[6], q[7], D(q[4]), "USR-06"])
    t.finish()

    lh = ["System_ID", "Supplier_Quote_ID", "Supplier_ID", "Unit_Price", "Currency", "MOQ",
          "Lead_Time_Days", "Payment_Terms", "Incoterm", "Freight", "Valid_To", "Price_Source",
          "Price_Confirm_Status", "Quality_Score", "OnTime_Score", "Capacity_Score", "Risk_Score",
          "Longevity_Score", "Score_Total", "Selected", "Landed_Estimate", "Create_Date", "Created_By"]
    lt = TB(wb, "Supplier_Quote_Lines", "tblSupplierQuoteLines", lh,
            note="معیارهای قیمت/کیفیت/زمان/ریسک؛ Score_Total صرفاً جنبه مقایسه دارد و سیستم به‌طور خودکار «بهترین» را اعلام نمی‌کند.")
    for q in seed.SUPPLIER_QUOTE_LINES:
        r = lt.row + 1
        lt.add([q[0], q[1], q[2], q[3], q[4], q[5], q[6], q[7], q[8], q[9],
                    D(q[10]), q[11], q[12], q[13], q[14], q[15], q[16], q[17],
                    F(f"=SUM($N{r}:$R{r})"), q[18],
                    F(f"=$D{r}+$J{r}"), D("2026-09-15"), "USR-06"])
    lt.finish()


def _build_procurement(wb):
    h = ["System_ID", "Project_ID", "Material_ID", "Supplier_ID", "Buyer", "Currency", "Fx_Rate",
         "Buy_Qty", "Unit_Purchase_Price", "MOQ_Applied", "Lead_Time_Days", "Payment_Terms", "Incoterm",
         "Freight_Per_Unit", "Insurance_Per_Unit", "Customs_Per_Unit", "Handling_Per_Unit", "Other_Per_Unit",
         "Landed_Cost_Unit", "Price_Date", "Price_Source", "Price_Valid_To", "Price_Expired_Flag",
         "Status_ID", "Create_Date", "Created_By"]
    t = TB(wb, "Procurement", "tblProcurement", h,
           note="تصمیم خرید تأییدشده + Price Build-up: Unit + Freight + Insurance + Customs + Handling + Other = Landed (با تبدیل ارز).")
    for i, p in enumerate(_proc_rows(), start=1):
        r = t.row + 1
        t.add([
            f"PRC-{i:04d}", p["project"], p["material"], p["supplier"], "USR-06", p["currency"], p["fx"],
            p["buy_qty"], p["unit"], p["moq"], p["lead"], p["pay"], p["incoterm"],
            p["freight"], p["insurance"], p["customs"], p["handling"], p["other"],
            F(f'=ROUND(($I{r}+$N{r}+$O{r}+$P{r}+$Q{r}+$R{r})*$G{r},0)'),
            D(p["price_date"]), p["source"], D(p["valid_to"]),
            F(f'=IF(ISBLANK($V{r}),"",IF($V{r}<TODAY(),1,0))'),
            p["status"], D("2026-09-28"), "USR-06",
        ])
    t.finish()


def _proc_rows():
    """Derive procurement rows from the selected supplier quote line."""
    out = []
    sel = {q[0]: q for q in seed.SUPPLIER_QUOTE_LINES if q[18] == 1}
    for sq in seed.SUPPLIER_QUOTES:
        sl = sel.get(sq[7])
        if not sl:
            continue
        out.append(dict(
            project=sq[1], material=sq[2], supplier=sl[2], currency=sl[4],
            fx=1 if sl[4] == "IRR" else setting_num("Fx_USD_TO_IRR"),
            buy_qty=sq[3], unit=sl[3], moq=sl[5], lead=sl[6], pay=sl[7], incoterm=sl[8],
            freight=sl[9], insurance=0, customs=0, handling=0, other=0,
            price_date="2026-09-15", source=sl[11], valid_to=sl[10], status="Approved",
        ))
    return out


def _build_purchase_prices(wb):
    h = ["System_ID", "Material_ID", "Supplier_ID", "Revision", "Price_Date", "Price_Source",
         "Currency", "Fx_Rate", "Unit_Price", "Freight_Per_Unit", "Insurance_Per_Unit",
         "Customs_Per_Unit", "Handling_Per_Unit", "Other_Per_Unit", "Landed_Cost_Unit",
         "Valid_From", "Valid_To", "Expired_Flag", "Status_ID", "Approved_By", "Create_Date", "Created_By"]
    t = TB(wb, "Purchase_Prices", "tblPurchasePrices", h,
           note="قیمت خرید با Revision؛ Expired_Flag=1 یعنی قیمت منقضی و از Costing کنار گذاشته می‌شود. Landed=(اجزا)×Fx.")
    for p in seed.PURCHASE_PRICES:
        r = t.row + 1
        t.add([p[0], p[1], p[2], p[3], D(p[4]), p[5], p[6], p[7], p[8], p[9], p[10],
                   p[11], p[12], p[13],
                   F(f'=ROUND(($I{r}+$J{r}+$K{r}+$L{r}+$M{r}+$N{r})*$H{r},0)'),
                   D(p[14]), D(p[15]),
                   F(f'=IF(ISBLANK($Q{r}),1,IF($Q{r}<TODAY(),1,0))'),
                   p[16], p[17], D(p[4]), "USR-06"])
    t.finish()


def _build_planning(wb, SET):
    h = ["System_ID", "Project_ID", "BOM_Revision_ID", "BOM_Line_ID", "Product_ID", "Material_ID",
         "Make_Buy", "Order_Qty", "Bom_Qty_Per", "Scrap_%", "Required_Qty", "Stock_On_Hand", "Reserved",
         "Deficit", "Buy_Qty", "Lead_Time_Days", "Supply_Date", "Status_ID",
         "Need_Date", "Suggested_Order_Date", "Create_Date", "Created_By"]
    t = TB(wb, "Planning", "tblPlanning", h,
           note="Required = Bom_Qty × (1+Scrap%) × Order_Qty. Order_Qty زنده از Order_Lines خوانده می‌شود. "
                "Need_Date = زمان‌بندی نیاز به اقلام برای تولید (ورودی برنامه‌ریزی)؛ تاریخ سفارش‌گذاری پیشنهادی = تاریخ نیاز − لیدتایم.")
    seq = 0
    for pl in planning_rows():
        seq += 1
        r = t.row + 1
        t.add([
            f"PLN-{seq:06d}", pl["project"], pl["bom_rev"], pl["bom_line"], pl["product"], pl["material"],
            pl["make_buy"],
            F(f'=SUMIFS(Order_Lines!$E$3:$E$20,Order_Lines!$D$3:$D$20,$E{r},'
              f'Order_Lines!$B$3:$B$20,XLOOKUP($B{r},Projects!$A$3:$A$10,Projects!$C$3:$C$10,0))'),
            pl["bom_qty"], pl["scrap"],
            F(f"=ROUND($H{r}*$I{r}*(1+$J{r}),2)"),
            F(f'=IFERROR(XLOOKUP($F{r},Inventory!$B$3:$B$20,Inventory!$C$3:$C$20,0),0)'),
            F(f'=IFERROR(XLOOKUP($F{r},Inventory!$B$3:$B$20,Inventory!$D$3:$D$20,0),0)'),
            F(f'=MAX($K{r}-($L{r}-$M{r}),0)'),
            F(f"=$N{r}"),
            15, F(f"=TODAY()+$P{r}"),
            "Approved",
            D("2026-11-18"),  # Need_Date — زمان‌بندی نیاز به اقلام جهت تولید (ورودی برنامه‌ریزی)
            F(f'=IF($S{r}="","",$S{r}-$P{r})'),
            D("2026-10-01"), "USR-05",
        ])
    t.finish()


def _build_schedule(wb):
    h = ["System_ID", "Project_ID", "Activity", "Start_Date", "End_Date", "Duration",
         "Responsible", "Dependency", "Status_ID", "Delay_Days", "Delay_Reason", "Create_Date", "Created_By"]
    t = TB(wb, "Schedule", "tblSchedule", h,
           note="زمان‌بندی فعالیت‌ها؛ Duration = End - Start + 1 و Delay برای فعالیت‌های معوق محاسبه می‌شود.")
    for s in seed.SCHEDULE:
        r = t.row + 1
        t.add([s[0], s[1], s[2], D(s[3]), D(s[4]),
                   F(f'=IF(ISBLANK($E{r}),0,$E{r}-$D{r}+1)'),
                   s[5], s[6], s[7],
                   F(f'=IF(ISBLANK($E{r}),0,IF($I{r}="Done",0,IF(AND($I{r}<>"Done",TODAY()>$D{r}),TODAY()-$D{r},0)))'),
                   "", D("2026-09-28"), "USR-05"])
    # overdue activity on the second project (delay demo)
    r = t.row + 1
    t.add(["SCH-000266", "PRJ-2026-00126", "Engineering", D("2026-09-01"), D("2026-09-15"),
               F(f'=IF(ISBLANK($E{r}),0,$E{r}-$D{r}+1)'),
               "USR-04", "", "In Progress",
               F(f'=IF(ISBLANK($E{r}),0,IF($I{r}="Done",0,IF(AND($I{r}<>"Done",TODAY()>$D{r}),TODAY()-$D{r},0)))'),
               "در انتظار تأیید نقشه", D("2026-09-01"), "USR-04"])
    t.finish()


# --------------------------------------------------------------------------- #
# Costing + Cost_Lines
# --------------------------------------------------------------------------- #
def _build_costlining(wb, SET):
    lh = ["System_ID", "Costing_ID", "Project_ID", "Category", "Material_ID", "Supplier_ID",
          "Qty", "UoM", "Unit_Cost", "Amount", "Source_Ref", "Memo", "Create_Date", "Created_By"]
    lt = TB(wb, "Cost_Lines", "tblCostLines", lh,
            note="ریز اجزای قیمت تمام‌شده (Drill-down). Direct Material = اقلام Make، Purchased Parts = اقلام Buy (غیر بسته‌بندی)، Packaging = بسته‌بندی.")
    CST = "CST-000321"
    PRJ = "PRJ-2026-00125"

    # active purchase price (last row per material = highest revision) → landed addr
    mat_landed = {}
    for i, p in enumerate(seed.PURCHASE_PRICES):
        mat_landed[p[1]] = f"Purchase_Prices!$O${3 + i}"

    cats = {"Direct Material": [], "Purchased Parts": [], "Packaging": []}
    cli = 1
    for pl in planning_rows():
        mat = pl["material"]
        if mat == "MAT-001548":
            continue  # packaging modelled separately (dedicated Cost_Line)
        if mat not in mat_landed:
            continue
        cat = "Direct Material" if pl["make_buy"] == "Make" else "Purchased Parts"
        r = lt.row + 1
        lt.add([
            f"CLI-{cli:06d}", CST, PRJ, cat, mat, pl["supplier"],
            pl["required"], pl["uom"],
            F(f"={mat_landed[mat]}"),
            F(f"=ROUND($G{r}*$I{r},0)"),
            f"PPR-{mat}", "از قیمت خرید فعال (Landed)", D("2026-10-05"), "USR-07",
        ])
        cats[cat].append(f"$J${r}")
        cli += 1

    def cat_sum(cat):
        a = cats[cat]
        return "+".join(a) if a else "0"

    labor_rate = SET["Labor_Rate_Per_Hour"]
    oh_rate = SET["Manufacturing_Overhead_Rate"]
    scrap_factor = SET["Scrap_Factor"]
    sga_rate = SET["SG&A_Rate"]
    tooling = SET["Tooling_Cost_Per_Project"]
    logistics = SET["Logistics_Base_Cost"]
    packaging_unit = SET["Packaging_Unit_Cost"]

    dm = cat_sum("Direct Material")
    pp = cat_sum("Purchased Parts")
    pk = "0"   # resolved after the Packaging line is added below

    # labour / overhead — نفرساعت به‌صورت زنده از زمان‌سنجی مهندسی (Time_Study) خوانده می‌شود
    r = lt.row + 1
    lt.add([f"CLI-{cli:06d}", CST, PRJ, "Direct Labor", "", "",
                F(f'=SUMIFS(Time_Study!$L$3:$L$20,Time_Study!$B$3:$B$20,$C{r})'), "hr",
                F(f"={labor_rate}"), F(f"=ROUND($G{r}*$I{r},0)"),
                "Time_Study", "نفرساعت از زمان‌سنجی مهندسی × نرخ دستمزد (Settings)", D("2026-10-05"), "USR-07"])
    labor_cells = f"$J${r}"; cli += 1
    r = lt.row + 1
    lt.add([f"CLI-{cli:06d}", CST, PRJ, "Manufacturing Overhead", "", "", 1, "lot",
                F(f"={oh_rate}"), F(f"=ROUND({labor_cells}*$I{r},0)"),
                "Settings", "ASSUMPTION: نرخ سربار × دستمزد مستقیم (جذب بر پایه کار)", D("2026-10-05"), "USR-07"])
    oh_cells = f"$J${r}"; cli += 1
    # scrap
    r = lt.row + 1
    lt.add([f"CLI-{cli:06d}", CST, PRJ, "Scrap", "", "", 1, "lot",
                F(f"={scrap_factor}"), F(f"=ROUND(({dm}+{pp})*$I{r},0)"),
                "Settings", "ASSUMPTION: ضریب ضایعات فرآیند", D("2026-10-05"), "USR-07"]); cli += 1
    # packaging (mirrors BOM packaging line cost)
    r = lt.row + 1
    lt.add([f"CLI-{cli:06d}", CST, PRJ, "Packaging", "MAT-001548", "", 2, "pcs",
                F(f"={packaging_unit}"), F(f"=ROUND($G{r}*$I{r},0)"),
                "Settings", "بسته‌بندی از Settings/تأمین‌کننده", D("2026-10-05"), "USR-07"])
    pk = f"$J${r}"; cli += 1
    # logistics
    r = lt.row + 1
    lt.add([f"CLI-{cli:06d}", CST, PRJ, "Logistics", "", "", 1, "lot",
                F(f"={logistics}"), F(f"=ROUND($G{r}*$I{r},0)"),
                "Settings", "ASSUMPTION: لجستیک پایه پروژه", D("2026-10-05"), "USR-07"]); cli += 1
    # tooling
    r = lt.row + 1
    lt.add([f"CLI-{cli:06d}", CST, PRJ, "Tooling", "", "", 1, "lot",
                F(f"={tooling}"), F(f"=ROUND($G{r}*$I{r},0)"),
                "Settings", "ASSUMPTION: استهلاک قالب (تسهیم خطی)", D("2026-10-05"), "USR-07"]); cli += 1
    # SG&A only if policy enabled
    r = lt.row + 1
    lt.add([f"CLI-{cli:06d}", CST, PRJ, "SG&A", "", "", 1, "lot",
                F(f"={sga_rate}"),
                F(f'=ROUND(IF(Settings!$B$8=1,({dm}+{pp}+{pk}+{labor_cells}+{oh_cells})*$I{r},0),0)'),
                "Settings", "ASSUMPTION: نرخ اداری/فروش (در صورت سیاست سازمان)", D("2026-10-05"), "USR-07"]); cli += 1
    lt.finish()

    # Costing header
    ch = ["System_ID", "Project_ID", "Revision", "Status_ID", "Cost_Basis", "Qty_Basis",
          "Direct_Material", "Purchased_Parts", "Direct_Labor", "Mfg_Overhead", "Scrap",
          "Packaging", "Logistics", "Tooling", "SG&A", "Other_Allocated", "Total_Cost",
          "Cost_Per_Unit", "Approved_By", "Create_Date", "Created_By"]
    ct = TB(wb, "Costing", "tblCosting", ch,
            note="قیمت تمام‌شده = SUMIFS روی Cost_Lines به تفکیک Category. Cost_Per_Unit = Total_Cost / Qty_Basis.")
    first_l, last_l = 3, 2 + cli
    amt_col = col_letter(10)
    costcol = col_letter(2)
    catcol = col_letter(4)

    def cat_formula(cat, r):
        return (f'=SUMIFS(Cost_Lines!${amt_col}${first_l}:${amt_col}${last_l},'
                f'Cost_Lines!${costcol}${first_l}:${costcol}${last_l},$A{r},'
                f'Cost_Lines!${catcol}${first_l}:${catcol}${last_l},"{cat}")')

    for c in seed.COSTING:
        r = ct.row + 1
        ct.add([
            c[0], c[1], c[2], c[3], D(c[4]), c[5],
            F(cat_formula("Direct Material", r)), F(cat_formula("Purchased Parts", r)),
            F(cat_formula("Direct Labor", r)), F(cat_formula("Manufacturing Overhead", r)),
            F(cat_formula("Scrap", r)), F(cat_formula("Packaging", r)), F(cat_formula("Logistics", r)),
            F(cat_formula("Tooling", r)), F(cat_formula("SG&A", r)), 0,
            F(f'=SUM($G{r}:$P{r})'),
            F(f'=IF($F{r}=0,0,ROUND($Q{r}/$F{r},0))'),
            c[6], D(c[4]), "USR-07",
        ])
    ct.finish()


def _build_budget(wb):
    h = ["System_ID", "Project_ID", "Revision", "Period", "Budget_Material", "Budget_Labor",
         "Budget_Overhead", "Budget_Procurement", "Budget_Logistics", "Budget_Tooling", "Budget_Other",
         "Total_Budget", "Actual_Material", "Actual_Labor", "Actual_Overhead",
         "Actual_Logistics", "Actual_Other", "Actual_Total", "Variance_Total", "Budget_vs_Actual",
         "Create_Date", "Created_By"]
    t = TB(wb, "Budget", "tblBudget", h,
           note="بودجه اولیه در برابر عملکرد (Actual از Costing). Variance و Budget_vs_Actual زنده محاسبه می‌شوند.")
    for b in seed.BUDGET:
        r = t.row + 1
        t.add([
            b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7], b[8], b[9], b[10],
            F(f"=SUM($E{r}:$K{r})"),
            F(f"=Costing!$G$3+Costing!$H$3"),
            F(f"=Costing!$I$3"),
            F(f"=Costing!$J$3"),
            F(f"=Costing!$L$3+Costing!$M$3"),
            F(f"=Costing!$N$3+Costing!$O$3"),
            F(f"=Costing!$Q$3"),
            F(f"=$R{r}-$L{r}"),
            F(f'=IF($L{r}=0,"",ROUND(($R{r}/$L{r})-1,4))'),
            D("2026-09-25"), "USR-07",
        ])
    t.finish()


def _build_payment_terms(wb, SET):
    h = ["System_ID", "Quotation_ID", "Advance_Payment_%", "Payment_On_Delivery_%", "Credit_Days",
         "Installment_Terms", "Bank_Guarantee", "Currency", "Payment_Risk_Level", "Fin_Cost_Rate_Annual",
         "Fin_Cost_Amount", "Is_Active_FLAG", "Create_Date", "Created_By"]
    t = TB(wb, "Payment_Terms", "tblPaymentTerms", h,
           note="شرایط پرداخت؛ هزینه مالی اعتبار = بخش اعتباری × نرخ سالانه × مدت/360 (در Sales_Quotation لحاظ می‌شود).")
    for pt in seed.PAYMENT_TERMS:
        r = t.row + 1
        t.add([
            pt[0], pt[1], pt[2], pt[3], pt[4], pt[5], pt[6], pt[7], pt[8],
            F(f"={SET['Financial_Cost_Rate_Annual']}"),
            F(f"=IF(Sales_Quotation!$K$4=0,0,Sales_Quotation!$K$4)"),
            pt[9], D("2026-10-06"), "USR-01",
        ])
    t.finish()


def _build_sales_quotation(wb, SET):
    h = ["System_ID", "Project_ID", "Quotation_Code", "Scenario", "Costing_ID", "Total_Cost",
         "Desired_Margin_%", "Margin_Amount", "Commercial_Adjustment", "Risk_Adjustment",
         "Financial_Adjustment", "Proposed_Selling_Price", "Currency", "Basis_Memo", "Status_ID",
         "Create_Date", "Created_By"]
    t = TB(wb, "Sales_Quotation", "tblSalesQuotation", h,
           note="سناریوهای A/B/C. Proposed = Cost + Margin + Commercial + Risk + Financial (هزینه مالی اعتبار). هیچ سناریو بدون Basis_Memo معتبر نیست.")
    fin_rate = SET["Financial_Cost_Rate_Annual"]
    # payment terms single row = row 3
    for i, q in enumerate(seed.SALES_QUOTATION):
        r = t.row + 1
        t.add([
            q[0], q[1], q[2], q[3], q[4],
            F(f'=IF($D{r}<>"",XLOOKUP($E{r},Costing!$A$3:$A$10,Costing!$Q$3:$Q$10,0),0)'),
            q[5],
            F(f"=ROUND($F{r}*$G{r},0)"),
            q[6], q[7],
            F(f'=ROUND((1-Payment_Terms!$C$3-Payment_Terms!$D$3)*$F{r}*{fin_rate}*(Payment_Terms!$E$3/360),0)'),
            F(f"=ROUND($F{r}+$H{r}+$I{r}+$J{r}+$K{r},0)"),
            q[9], q[10], q[11], D("2026-10-06"), "USR-01",
        ])
    t.finish()


def _build_quotation_versions(wb):
    h = ["System_ID", "Quotation_ID", "Version_No", "Costing_Revision_ID", "Total_Cost",
         "Margin_Amount", "Proposed_Selling_Price", "Change_Summary", "Approved_Step", "Status_ID",
         "Create_Date", "Created_By"]
    t = TB(wb, "Quotation_Versions", "tblQuotationVersions", h,
           note="نسخه‌های پیشنهاد؛ نسخه قبلی هرگز Overwrite نمی‌شود. (نمونه: نسخه V01 از سناریو B)")
    sq_row = 4  # Scenario B
    for v in seed.QUOTATION_VERSIONS:
        r = t.row + 1
        t.add([
            v[0], v[1], v[2], v[3],
            F(f"=Sales_Quotation!$F${sq_row}"),
            F(f"=Sales_Quotation!$H${sq_row}"),
            F(f"=Sales_Quotation!$L${sq_row}"),
            v[4], v[5], "Approved", D("2026-10-06"), "USR-01",
        ])
    t.finish()


def _build_management_approval(wb):
    h = ["System_ID", "Quotation_Version_ID", "Project_ID", "Decision", "Decision_Maker",
         "Decision_Date", "Comment", "Version_Ref", "Status_ID", "Create_Date", "Created_By"]
    t = TB(wb, "Management_Approval", "tblManagementApproval", h,
           note="تصمیم مدیریت ارشد: APPROVED / APPROVED WITH CONDITION / RETURN FOR REVISION / REJECTED.")
    for a in seed.MANAGEMENT_APPROVAL:
        t.add([a[0], a[1], a[2], a[3], a[4], D(a[5]), a[6], a[7], a[8], D(a[5]), "USR-08"])
    t.finish()


def _build_workflow_history(wb):
    h = ["Workflow_History_ID", "Project_ID", "Entity_Type", "Entity_ID", "From_Status",
         "To_Status", "Action_Date", "User_ID", "Comment"]
    t = TB(wb, "Workflow_History", "tblWorkflowHistory", h,
           note="لاگ انتقال وضعیت — پایه Audit (چه کسی / چه زمانی / از چه وضعیتی / به چه وضعیتی).")
    for w in seed.WORKFLOW_HISTORY:
        t.add([w[0], w[1], w[2], w[3], w[4], w[5], D(w[6]), w[7], w[8]])
    t.finish()


def _build_change_log(wb):
    h = ["Change_Log_ID", "Project_ID", "Entity_Type", "Entity_ID", "Field_Changed", "Old_Value",
         "New_Value", "Change_Reason", "Changed_By", "Change_Date", "Revision_After", "Create_Date", "Created_By"]
    t = TB(wb, "Change_Log", "tblChangeLog", h,
           note="کنترل تغییرات — هر تغییر پس از تأیید Revision جدید می‌سازد؛ مقدار قبلی ثبت می‌ماند.")
    for c in seed.CHANGE_LOG:
        t.add([c[0], c[1], c[2], c[3], c[4], c[5], c[6], c[7], c[8], D(c[9]), c[10], D(c[9]), "USR-04"])
    t.finish()


def _build_issues(wb):
    h = ["Issue_ID", "Project_ID", "Error_Code", "Error_Description", "Owner", "Required_Action",
         "Status", "Severity", "Raised_Date", "Create_Date", "Created_By"]
    t = TB(wb, "Issues", "tblIssues", h, note="ثبت ریسک/خطا با مالک و اقدام لازم.")
    for i in seed.ISSUES:
        t.add([i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], D(i[8]), D(i[8]), "USR-03"])
    t.finish()


def _build_error_checks(wb):
    h = ["Rule_Key", "Error_Code", "Check_Description", "Severity", "Owner", "Formula_Result",
         "Message", "Required_Action", "Status", "Create_Date", "Created_By"]
    t = TB(wb, "Error_Checks", "tblErrorChecks", h,
           note="کنترل‌های خودکار (Error/Exception Handling). Formula_Result زنده محاسبه می‌شود.")
    of, ol = 3, 3 + len(seed.ORDERS) - 1
    bhf, bhl = 3, 3 + len(seed.BOM_HEADERS) - 1
    pf, pl = 3, 3 + len(seed.PURCHASE_PRICES) - 1

    def add(rule, code, desc, sev, owner, formula, msg, action):
        t.add([rule, code, desc, sev, owner, F(formula), msg, action, "Open", D("2026-09-22"), "USR-09"])

    add("DK-01", "E-018", "Duplicate Order_ID", "Error", "Sales",
        f'=IF(COUNTIF(Orders!$S$3:$S$1000,"DUP")>0,"ERROR","OK")',
        "شناسه سفارش تکراری است", "بررسی و اصلاح")
    add("G1", "E-001", "سفارش ناقص (Ready_Flag)", "Block", "Sales",
        f'=IF(COUNTIF(Orders!$R${of}:$R${ol},"INCOMPLETE")>0,"BLOCKED","OK")',
        "هر سفارش باید Ready_Flag=READY باشد تا وارد مهندسی شود", "تکمیل اطلاعات")
    add("DK-02", "E-019", "Duplicate BOM_ID+Revision", "Error", "Engineering",
        f'=IF(MAX(COUNTIFS(BOM_Header!$B${bhf}:$B${bhl},BOM_Header!$B${bhf}:$B${bhl},BOM_Header!$D${bhf}:$D${bhl},BOM_Header!$D${bhf}:$D${bhl}))>1,"ERROR","OK")',
        "ترکیب BOM_ID+Revision تکراری است", "صدور Revision جدید")
    add("G2", "E-002", "BOM بدون Revision فعال", "Block", "Engineering",
        f'=IF(COUNTIF(BOM_Header!$I${bhf}:$I${bhl},"Approved")>0,"OK","BLOCKED")',
        "بدون BOM Revision تأییدشده نمی‌توان وارد Planning شد", "تأیید BOM")
    add("PX-01", "E-003", "قیمت خرید بدون تاریخ", "Error", "Procurement",
        f'=IF(COUNTIF(Purchase_Prices!$E${pf}:$E${pl},"")>0,"ERROR","OK")',
        "قیمتی بدون تاریخ وجود دارد", "ثبت Price_Date")
    add("PX-02", "E-004", "قیمت بدون منبع", "Error", "Procurement",
        f'=IF(COUNTIF(Purchase_Prices!$F${pf}:$F${pl},"")>0,"ERROR","OK")',
        "قیمتی بدون منبع وجود دارد", "ثبت Price_Source")
    add("PX-05", "E-010", "قیمت منقضی", "Warning", "Procurement",
        f'=IF(COUNTIF(Purchase_Prices!$R${pf}:$R${pl},1)>0,"WARN","OK")',
        "یک یا چند قیمت منقضی شده و از Costing کنار گذاشته می‌شود", "استعلام/تمدید قیمت")
    add("PX-03", "E-007", "قیمت منفی", "Error", "Finance",
        f'=IF(COUNTIF(Purchase_Prices!$I${pf}:$I${pl},"<0")>0,"ERROR","OK")',
        "قیمت منفی ثبت شده است", "اصلاح قیمت")
    add("RV-01", "E-015", "تغییر پس از Approval", "Warning", "All",
        f'="تغییر فقط با Revision جدید (Create New Revision)"',
        "رکورد تأییدشده نباید Overwrite شود", "ایجاد Revision")
    add("BM-01", "E-016", "BOM نسخه قدیمی موجود", "Info", "Engineering",
        f'=IF(COUNTIF(BOM_Header!$J${bhf}:$J${bhl},0)>0,"INFO","OK")',
        "BOM قدیمی (Superseded) موجود است؛ فقط نسخه فعال مصرف می‌شود", "اطمینان از استفاده نسخه فعال")
    add("SC-01", "E-011", "تأخیر فعالیت", "Warning", "PMO",
        f'=IF(SUM(Schedule!$J$3:$J$20)>0,"WARN","OK")',
        "فعالیت معوق در زمان‌بندی وجود دارد", "بازبینی Schedule")
    add("BG-01", "E-013", "اختلاف بودجه و قیمت تمام‌شده", "Info", "Finance",
        f'=IF(Budget!$S$3<>0,"DIFF","OK")',
        "اختلاف بین بودجه و Actual ثبت شده است", "بازبینی Costing/Budget")
    add("CS-01", "E-021", "Costing ناقص", "Error", "Finance",
        f'=IF(Costing!$Q$3=0,"ERROR","OK")',
        "Total_Cost صفر است", "تکمیل ردیف‌های هزینه")
    add("SQ-02", "E-008", "Margin غیرمنطقی", "Warning", "Sales",
        f'=IF(OR(COUNTIF(Sales_Quotation!$G$3:$G$10,"<0")>0,COUNTIF(Sales_Quotation!$G$3:$G$10,">0.35")>0),"WARN","OK")',
        "حاشیه سود منفی یا بالاتر از سقف منطقی است", "بازبینی سناریو")
    # --- v2.0 gate controls (تقدم و تأخر ورود اطلاعات) ---
    add("GT-01", "E-022", "گیت: امکان‌سنجی ناقص", "Block", "PMO",
        f'=IF(COUNTIF(Feasibility!$M$3:$M$20,"INCOMPLETE")>0,"BLOCKED","OK")',
        "تا امکان‌سنجی کامل نشود، گیت مهندسی برای آن پروژه بسته است", "تکمیل امکان‌سنجی توسط مدیر پروژه")
    add("GT-02", "E-023", "گیت: تجهیزات قیمت‌دهی‌نشده", "Warning", "Procurement",
        f'=IF(COUNTIF(Project_Equipment!$N$3:$N$20,"PENDING")>0,"WARN","OK")',
        "تجهیزاتی وجود دارد که بازرگانی هنوز قیمت آن‌ها را ثبت نکرده است", "قیمت‌دهی در Project_Equipment")
    add("GT-03", "E-024", "گیت: رسید انبار بدون ردیف برنامه‌ریزی", "Error", "Planning",
        f'=IF(COUNTIF(Receipts!$C$3:$C$20,"")>0,"ERROR","OK")',
        "رسید انبار باید به ردیف معتبر برنامه‌ریزی (PLN) متصل باشد", "اصلاح Planning_ID")
    add("GT-04", "E-025", "گیت: جمع‌بندی گزارش‌نشده به ارشد", "Warning", "PMO",
        f'=IF(COUNTIF(PMO_Summary!$J$3:$J$10,"NOT REPORTED")>0,"WARN","OK")',
        "پروژه‌هایی هنوز گزارش جمع‌بندی آن‌ها به مدیریت ارشد اعلام نشده است", "تکمیل PMO_Summary و ثبت تاریخ اعلام")
    add("GT-05", "E-026", "گیت: زمان‌سنجی ثبت‌نشده برای پروژه در حال مهندسی", "Block", "Engineering",
        f'=IF(COUNTA(Time_Study!$A$3:$A$20)=0,"BLOCKED","OK")',
        "بدون زمان‌سنجی، نفرساعت و دستمزد قابل محاسبه نیست", "ثبت عملیات در Time_Study")
    t.finish()


def _build_test_results(wb):
    h = ["Test_ID", "Test_Scenario", "Input", "Expected_Result", "Actual_Result", "Pass_Fail", "Issue",
         "Executed_Date", "By"]
    t = TB(wb, "Test_Results", "tblTestResults", h,
           note="نتایج ۱۵ تست اجباری؛ پس از اجرای راستی‌آزمایی، Actual_Result/Pass_Fail به‌روز می‌شوند.")
    tests = [
        ("T1", "سفارش کامل", "Order ORD-2026-00321 کامل", "Ready_Flag=READY و عبور تا Approval", "PENDING", "PENDING", ""),
        ("T2", "اطلاعات ناقص", "Order بدون مقدار", "CO_02=MISSING + Ready=INCOMPLETE + Issue", "PENDING", "PENDING", ""),
        ("T3", "BOM Revision جدید", "صدور R05 بالای R04", "R05 فعال و R04 حفظ شود", "PENDING", "PENDING", ""),
        ("T4", "تغییر قیمت تأمین‌کننده", "شفت R01→R02", "Landed جدید + Revision حفظ", "PENDING", "PENDING", ""),
        ("T5", "تغییر مقدار سفارش", "Qty 2→3", "Planning/نیاز مقیاس می‌شود", "PENDING", "PENDING", ""),
        ("T6", "تغییر شرایط پرداخت", "Credit 60→90", "Financial_Adjustment افزایش", "PENDING", "PENDING", ""),
        ("T7", "عودت مدیریت", "Decision=RETURN FOR REVISION", "Quotation=Returned + Issue", "PENDING", "PENDING", ""),
        ("T8", "رد پیشنهاد", "Decision=REJECTED", "Quotation=Rejected/Closed", "PENDING", "PENDING", ""),
        ("T9", "تصویب پیشنهاد", "Decision=APPROVED", "APPROVED + Workflow_History", "PENDING", "PENDING", ""),
        ("T10", "تغییر بعد از Approval", "تغییر Qty پس از APPROVED", "Revision جدید + Approval قبلی حفظ", "PENDING", "PENDING", ""),
        ("T11", "قیمت منقضی", "Valid_To < امروز", "Expired_Flag=1 و حذف از Costing", "PENDING", "PENDING", ""),
        ("T12", "تأخیر تأمین‌کننده", "Arrival بعد از برنامه", "Schedule.Delay>0 / Issue", "PENDING", "PENDING", ""),
        ("T13", "اختلاف بودجه/هزینه", "Actual≠Budget", "Variance مشخص شود", "PENDING", "PENDING", ""),
        ("T14", "Duplicate Order", "Order_ID تکراری", "Error E-018", "PENDING", "PENDING", ""),
        ("T15", "تغییر BOM بعد از Costing", "BOM Revision جدید", "Costing مجدد لازم + نسخه قبل حفظ", "PENDING", "PENDING", ""),
        ("T16", "گیت امکان‌سنجی", "حذف امتیاز از امکان‌سنجی پروژه", "قفل گیت مهندسی (ورود اطلاعات مهندسی مجاز نیست)", "PENDING", "PENDING", ""),
        ("T17", "گیت بازرگانی", "پروژه فاقد زنجیره کامل (امکان‌سنجی ناقص)", "تجهیزات پروژه قابل قیمت‌دهی نیست (گیت بازرگانی بسته)", "PENDING", "PENDING", ""),
        ("T18", "زنجیره نفرساعت", "افزایش زمان استاندارد عملیات در زمان‌سنجی", "دستمزد مستقیم در قیمت تمام‌شده افزایش می‌یابد", "PENDING", "PENDING", ""),
        ("T19", "نرخ ارز اقلام وارداتی", "تغییر نرخ ارز در Settings", "قیمت ریالی تجهیزات/اقلام وارداتی به‌روز می‌شود", "PENDING", "PENDING", ""),
        ("T20", "گیت جمع‌بندی/تصمیم", "حذف تاریخ اعلام گزارش به ارشد", "گیت تصمیم مدیریت ارشد بسته می‌شود", "PENDING", "PENDING", ""),
    ]
    for tid, name, inp, exp, act, pf, iss in tests:
        t.add([tid, name, inp, exp, act, pf, iss, D("2026-09-23"), "USR-09"])
    t.finish()


def _build_workflow_sheet(wb):
    """DEPRECATED (v2.0): replaced by builder_v2.build_workflow_v2 — kept for reference only."""
    ws = wb.sheet("Workflow")
    ws.set(1, 1, value="Workflow — گردش کار سراسری و گیت‌ها")
    stages = [
        ("NEW — ثبت سفارش", "Sales"),
        ("SALES REVIEW — کنترل اولیه", "Sales"),
        ("PROJECT CREATED", "PMO"),
        ("ENGINEERING + BOM", "Engineering"),
        ("BOM APPROVAL", "Engineering Lead"),
        ("PLANNING — مقداردهی BOM", "Planning"),
        ("PROCUREMENT — استعلام/انتخاب", "Procurement"),
        ("COSTING — بهای تمام‌شده", "Finance"),
        ("FINANCE REVIEW", "Finance"),
        ("SALES QUOTATION", "Sales"),
        ("MANAGEMENT APPROVAL", "Management"),
        ("APPROVED / REVISION / REJECTED", "Management"),
    ]
    for i, (name, owner) in enumerate(stages):
        r = 3 + i * 2
        ws.set(r, 1, value=f"{i+1:02d}")
        ws.set(r, 2, value=name)
        ws.set(r, 3, value=f"Owner: {owner}")
        ws.set(r, 4, value="↓" if i < len(stages) - 1 else "")
    r = 3 + len(stages) * 2 + 1
    ws.set(r, 1, value="بازگشت‌ها: RETURN FOR REVISION → گام مشخص؛ هر تغییر پس از تأیید → Revision جدید + Change_Log + Workflow_History.")
    ws.set(r + 1, 1, value="هیچ مرحله‌ای بدون تکمیل پیش‌نیاز قابل عبور نیست (گیت‌ها در Error_Checks و ستون‌های قفل).")


def _build_dashboard(wb):
    """DEPRECATED (v2.0): replaced by builder_v2.build_dashboard_v2 — kept for reference only."""
    ws = wb.sheet("Dashboard")
    ws.set(1, 1, value="Dashboard — مدیریت فرآیند سفارش تا تصویب قیمت (KPI + Funnel + Timeline)")
    n_ord = len(seed.ORDERS)
    n_proj = len(seed.PROJECTS)
    of, ol = 3, 3 + n_ord - 1
    pf, pl = 3, 3 + n_proj - 1

    def kpi(label, formula, r):
        ws.set(r, 1, value=label)
        ws.set(r, 2, formula=formula)
        return r + 1

    r = 3
    r = kpi("پروژه‌های فعال", f'=COUNTIF(Projects!$K${pf}:$K${pl},"Approved")', r)
    r = kpi("سفارش‌های جدید", f'=COUNTA(Orders!$A${of}:$A${ol})', r)
    r = kpi("در انتظار بررسی فروش", f'=COUNTIF(Orders!$M${of}:$M${ol},"Draft")', r)
    r = kpi("پروژه‌های تأخیرخورده", f'=SUM(Projects!$N${pf}:$N${pl})', r)
    r = kpi("BOMهای در انتظار تأیید", '=COUNTIF(BOM_Header!$I$3:$I$5,"Waiting Approval")', r)
    r = kpi("قیمت‌های منقضی", '=COUNTIF(Purchase_Prices!$R$3:$R$10,1)', r)
    r = kpi("Costingهای تکمیل‌نشده", '=COUNTIF(Costing!$D$3:$D$3,"<>Approved")', r)
    r = kpi("پیشنهاد در انتظار مدیریت", '=COUNTIF(Sales_Quotation!$O$3:$O$5,"Submitted")', r)
    r = kpi("پیشنهاد تصویب‌شده", '=COUNTIF(Management_Approval!$D$3:$D$3,"APPROVED")', r)
    r = kpi("پیشنهاد مشروط", '=COUNTIF(Management_Approval!$D$3:$D$3,"APPROVED WITH CONDITION")', r)
    r = kpi("پیشنهاد برگشتی", '=COUNTIF(Management_Approval!$D$3:$D$3,"RETURN FOR REVISION")', r)
    r = kpi("پیشنهاد ردشده", '=COUNTIF(Management_Approval!$D$3:$D$3,"REJECTED")', r)
    r = kpi("Total Sales Value (پیشنهاد جاری)", '=Quotation_Versions!$G$3', r)
    r = kpi("Total Cost", '=SUM(Costing!$Q$3:$Q$3)', r)
    r = kpi("Gross Margin (IRR)", '=Quotation_Versions!$G$3-SUM(Costing!$Q$3:$Q$3)', r)
    r = kpi("زمان تصویب (روز، از ثبت پیشنهاد)", '=IFERROR(Management_Approval!$F$3-Sales_Quotation!$P$3,0)', r)
    r = kpi("زمان پیشنهاد (روز، از Costing)", '=IFERROR(Sales_Quotation!$P$3-Costing!$T$3,0)', r)

    r += 1
    ws.set(r, 1, value="Funnel (تعداد)")
    funnel = [
        ("Leads / سفارش‌ها", f'=COUNTA(Orders!$A${of}:$A${ol})'),
        ("Orders → Engineering", f'=COUNTIF(Orders!$R${of}:$R${ol},"READY")'),
        ("Engineering", '=COUNTIF(Engineering!$H$3:$H$4,"Approved")'),
        ("Costing", '=COUNTIF(Costing!$D$3:$D$3,"Approved")'),
        ("Quotation", '=COUNTIF(Sales_Quotation!$O$3:$O$5,"Submitted")'),
        ("Approval (کل تصمیمات)", '=COUNTA(Management_Approval!$D$3:$D$3)'),
        ("Approved", '=COUNTIF(Management_Approval!$D$3:$D$3,"APPROVED")'),
    ]
    for fname, fform in funnel:
        r += 1
        ws.set(r, 1, value=fname)
        ws.set(r, 2, formula=fform)

    r += 2
    ws.set(r, 1, value="Project Timeline (وضعیت واحدها)")
    ws.set(r + 1, 1, value="پروژه")
    for si, stg in enumerate(["Sales", "Engineering", "Planning", "Procurement", "Finance", "Management"]):
        ws.set(r + 1, 2 + si, value=stg)
    ps_first, ps_last = 3, 3 + len(seed.PROJECT_STATUSES) - 1
    stage_header_row = r + 1
    for pi, proj in enumerate(seed.PROJECTS):
        rr = r + 2 + pi
        ws.set(rr, 1, value=proj[0])
        for si, stg in enumerate(["Sales", "Engineering", "Planning", "Procurement", "Finance", "Management"]):
            stage_col = col_letter(2 + si)
            ws.set(rr, 2 + si, formula=(
                f'=IFERROR(XLOOKUP(1,(Project_Statuses!$B${ps_first}:$B${ps_last}=$A{rr})*'
                f'(Project_Statuses!$C${ps_first}:$C${ps_last}={stage_col}${stage_header_row}),'
                f'Project_Statuses!$D${ps_first}:$D${ps_last}),"Pending")'))


def _build_trace(wb):
    ws = wb.sheet("_Trace")
    ws.set(1, 1, value="_Trace — ردیابی قیمت نهایی تا منبع (Selling Price ← Cost ← BOM ← Material ← Supplier Quote ← Supplier)")
    ws.set(2, 1, value="سطح")
    ws.set(2, 2, value="عنصر")
    ws.set(2, 3, value="مرجع")
    ws.set(2, 4, value="مبلغ/مقدار (IRR)")
    rows = [
        ("Quotation", "Proposed Selling Price (Scenario B)", "Service", "=Sales_Quotation!L4"),
        ("Quotation", "Margin Amount", "Service", "=Sales_Quotation!H4"),
        ("Quotation", "Total Cost", "Service", "=Sales_Quotation!F4"),
        ("Costing", "Direct Material", "Service", "=Costing!G3"),
        ("Costing", "Purchased Parts", "Service", "=Costing!H3"),
        ("Costing", "Direct Labor", "Service", "=Costing!I3"),
        ("Costing", "Manufacturing Overhead", "Service", "=Costing!J3"),
        ("Costing", "Scrap", "Service", "=Costing!K3"),
        ("Costing", "Packaging", "Service", "=Costing!L3"),
        ("Costing", "Logistics + Tooling + SG&A", "Service", "=Costing!M3+Costing!N3+Costing!O3"),
        ("BOM", "BOM فعال پروژه", "Service", '=XLOOKUP("PRV-000125-R03",BOM_Header!$C$3:$C$10,BOM_Header!$A$3:$A$10,"؟")'),
        ("Material", "مکانیکال سیل MAT-001545 (نمونه)", "Service", '=XLOOKUP("MAT-001545",Materials!$A$3:$A$10,Materials!$C$3:$C$10,"؟")'),
        ("Supplier Quote", "۳ تأمین‌کننده سیل (مقایسه)", "Service", '=XLOOKUP("SQT-2026-00545",Supplier_Quote_Lines!$B$3:$B$20,Supplier_Quote_Lines!$D$3:$D$20,"؟")'),
    ]
    for i, (lvl, name, src, f) in enumerate(rows):
        r = 3 + i
        ws.set(r, 1, value=lvl)
        ws.set(r, 2, value=name)
        ws.set(r, 3, value=src)
        ws.set(r, 4, formula=f)
    ws.set(18, 1, value="مسیر کامل: Final Price → Costing → Cost_Line → Material → BOM → Supplier Quote → Supplier")


def _build_report(wb):
    ws = wb.sheet("Report_Quotation")
    ws.set(1, 1, value="گزارش اجرایی پیشنهاد فروش (قابل چاپ / PDF)")
    fields = [
        ("مشتری", '=XLOOKUP(Orders!$B$3,Customers!$A$3:$A$20,Customers!$C$3:$C$20,"")'),
        ("پروژه", '=Projects!$F$3'),
        ("محصول", '=XLOOKUP(Order_Lines!$D$3,Products!$A$3:$A$20,Products!$C$3:$C$20,"")'),
        ("مقدار", '=SUM(Order_Lines!$E$3:$E$20)'),
        ("قیمت تمام‌شده (Total Cost)", '=Costing!$Q$3'),
        ("قیمت پیشنهادی (Scenario B)", '=Sales_Quotation!$L$4'),
        ("Margin (IRR)", '=Sales_Quotation!$L$4-Costing!$Q$3'),
        ("شرایط پرداخت", '=TEXT(Payment_Terms!$C$3,"0%")&" پیش‌پرداخت / اعتبار "&Payment_Terms!$E$3&" روز"'),
        ("زمان تحویل", '=TEXT(Projects!$I$3,"yyyy-mm-dd")'),
        ("ریسک‌های مهم", '="انقضای قیمت شفت؛ تأمین مکانیکال سیل؛ تأخیر احتمالی تأمین"'),
        ("تغییرات نسخه", '=BOM_Header!$F$3'),
        ("تأییدهای انجام‌شده", '="BOM R05 ✓ / Costing ✓ / Quotation ✓ / Management ✓"'),
        ("تصمیم مدیریت", '=Management_Approval!$D$3'),
    ]
    for i, (label, formula) in enumerate(fields):
        r = 3 + i * 2
        ws.set(r, 1, value=label)
        ws.set(r, 2, formula=formula)


def _build_named_ranges_doc(wb, SET):
    h = ["Named_Range", "Refers_To", "Purpose"]
    t = TB(wb, "_NamedRanges", "tblNamedRanges", h,
           note="Named Ranges پیشنهادی برای خوانایی فرمول‌ها (در نسخه نهایی register می‌شوند).")
    t.add(["Rng_Fx_USD", SET["Fx_USD_TO_IRR"].replace("$B$", "$B$"), "نرخ تبدیل USD→IRR (مرجع)"])
    t.add(["Rng_ReportingCurrency", SET["Reporting_Currency"].replace("$B$", "$B$"), "ارز گزارشگری"])
    t.add(["Rng_LaborRate", SET["Labor_Rate_Per_Hour"].replace("$B$", "$B$"), "نرخ دستمزد ساعتی"])
    t.add(["Rng_OverheadRate", SET["Manufacturing_Overhead_Rate"].replace("$B$", "$B$"), "نرخ سربار تولید"])
    t.add(["Rng_FinCostRate", SET["Financial_Cost_Rate_Annual"].replace("$B$", "$B$"), "نرخ هزینه مالی سالانه"])
    t.finish()
