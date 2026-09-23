"""
renderer_v2.py — v2.0 rendering extensions
===========================================
1. RTL sheet views + tab colors
2. Gate Data Validation: hard blocking of data entry while the upstream
   stage's Gate_Status gate is closed (تقدم و تأخر ورود اطلاعات)
3. Sheet protection (admin password) + password-protected edit ranges per
   role → each unit edits only its own section; PM sees all but edits own
4. Gantt conditional-formatting bars
5. Dashboard charts
6. Post-save XML injection of <protectedRanges> (native Excel feature)
"""
from __future__ import annotations

import re
import zipfile
import shutil
import os

import seed
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.protection import SheetProtection, hash_password

# --------------------------------------------------------------------------- #
# Role passwords (single source of truth: seed.ACCESS_ROLES)
# --------------------------------------------------------------------------- #
ROLE_PASSWORDS = {a[0]: a[3] for a in seed.ACCESS_ROLES}

# --------------------------------------------------------------------------- #
# Tab colors (visual swimlanes)
# --------------------------------------------------------------------------- #
TAB_COLORS = {
    "Dashboard": "FFC000", "Workflow": "FFC000", "Gate_Status": "C00000",
    "Access_Control": "808080",
    # sales
    "Orders": "70AD47", "Order_Lines": "70AD47", "Sales_Quotation": "70AD47",
    "Payment_Terms": "70AD47", "Quotation_Versions": "70AD47",
    # PM
    "Projects": "0F4C81", "Feasibility": "0F4C81", "Schedule": "0F4C81",
    "Gantt": "0F4C81", "PMO_Summary": "0F4C81",
    # engineering
    "Engineering": "ED7D31", "BOM_Header": "ED7D31", "BOM_Detail": "ED7D31",
    "Time_Study": "ED7D31", "Materials": "ED7D31", "Equipment": "ED7D31",
    "Project_Equipment": "ED7D31",
    # commercial
    "Supplier_Quotes": "7030A0", "Supplier_Quote_Lines": "7030A0",
    "Procurement": "7030A0", "Purchase_Prices": "7030A0",
    # planning / warehouse
    "Planning": "2E9CCA", "Receipts": "2E9CCA", "Inventory": "2E9CCA",
    # finance
    "Cost_Lines": "806000", "Costing": "806000", "Budget": "806000",
    # senior management
    "Management_Approval": "C00000",
}
TAB_DEFAULT = "A6A6A6"

# --------------------------------------------------------------------------- #
# Gate Data Validation configuration
# Each entry: sheet -> list of (sqref_cols, key_column, gate_formula, message)
# The formula template uses {r} for the row and is anchored on the key column.
# --------------------------------------------------------------------------- #
G_OK = "OK"


def _proj_gate(gate_col, key="B"):
    """Entry allowed when project-key row's cumulative gate == 1 (or row blank)."""
    return (f'=OR(${key}{{r}}="",IFERROR(INDEX(Gate_Status!${gate_col}$3:${gate_col}$20,'
            f'MATCH(${key}{{r}},Gate_Status!$B$3:$B$20,0)),0)=1)')


def _any_gate(cell):
    return f"=Gate_Status!${cell}$1=1"


GATE_DV = {
    # ---- stage 2: مدیر پروژه (پس از آماده‌شدن سفارش)
    "Projects": [
        ("B:N", "C",
         '=OR($C{r}="",COUNTIFS(Orders!$A$3:$A$20,$C{r},Orders!$R$3:$R$20,"READY")>0)',
         "گیت ۱: ابتدا سفارش باید توسط فروش بررسی و تأیید شده باشد (Ready_Flag=READY)."),
    ],
    "Feasibility": [("D:L", "B", _proj_gate("N"),
                     "گیت ۱: ورود اطلاعات امکان‌سنجی فقط برای پروژه‌ای مجاز است که سفارش آن آماده باشد.")],
    "Schedule": [("D:E", "B", _proj_gate("N"),
                  "گیت ۱: زمان‌بندی فقط برای پروژه با سفارش آماده مجاز است."),
                 ("G:H", "B", _proj_gate("N"),
                  "گیت ۱: زمان‌بندی فقط برای پروژه با سفارش آماده مجاز است."),
                 ("K:K", "B", _proj_gate("N"),
                  "گیت ۱: زمان‌بندی فقط برای پروژه با سفارش آماده مجاز است.")],
    # ---- stage 3: مهندسی (پس از امکان‌سنجی کامل)
    "Engineering": [("B:M", "B", _proj_gate("O"),
                     "گیت ۲: ابتدا امکان‌سنجی پروژه باید توسط مدیر پروژه تکمیل شده باشد.")],
    "BOM_Header": [("C:J", None, _any_gate("AA"),
                    "گیت ۲: صدور BOM فقط وقتی حداقل یک پروژه به مرحله مهندسی مجاز رسیده باشد.")],
    "BOM_Detail": [("B:I", None, _any_gate("AA"),
                    "گیت ۲: ثبت ردیف BOM فقط پس از تکمیل امکان‌سنجی مجاز است."),
                   ("M:M", None, _any_gate("AA"),
                    "گیت ۲: ثبت ردیف BOM فقط پس از تکمیل امکان‌سنجی مجاز است.")],
    "Time_Study": [("B:H", "B", _proj_gate("O"),
                    "گیت ۲: زمان‌سنجی فقط پس از تکمیل امکان‌سنجی پروژه مجاز است.")],
    "Project_Equipment": [
        ("B:F", "B", _proj_gate("O"),
         "گیت ۲: انتخاب تجهیزات پروژه فقط پس از تکمیل امکان‌سنجی (مرحله مهندسی) مجاز است."),
        ("G:L", "B", _proj_gate("P"),
         "گیت ۳: قیمت‌دهی تجهیزات فقط توسط بازرگانی و پس از تکمیل مهندسی مجاز است."),
    ],
    # ---- stage 4: بازرگانی (پس از مهندسی)
    "Supplier_Quotes": [("B:H", None, _any_gate("AA"),
                         "گیت ۳: استعلام فقط وقتی حداقل یک پروژه از مهندسی عبور کرده باشد مجاز است.")],
    "Supplier_Quote_Lines": [("C:T", None, _any_gate("AA"),
                              "گیت ۳: ثبت ردیف استعلام فقط پس از عبور از مهندسی مجاز است.")],
    "Procurement": [("C:V", None, _any_gate("AC"),
                     "گیت ۳: ثبت خرید فقط پس از تکمیل مرحله مهندسی مجاز است.")],
    "Purchase_Prices": [("B:Q", None, _any_gate("AC"),
                         "گیت ۳: ثبت قیمت خرید فقط پس از تکمیل مرحله مهندسی مجاز است."),
                        ("T:T", None, _any_gate("AC"),
                         "گیت ۳: تأیید قیمت فقط پس از تکمیل مرحله مهندسی مجاز است.")],
    # ---- stage 5: برنامه‌ریزی تولید و انبارها (پس از بازرگانی)
    "Planning": [("R:S", "B", _proj_gate("Q"),
                  "گیت ۴: ثبت وضعیت/تاریخ نیاز فقط پس از تکمیل قیمت‌دهی بازرگانی مجاز است.")],
    "Receipts": [("B:B", None, _any_gate("AE"),
                  "گیت: ثبت رسید انبار پس از عبور حداقل یک پروژه از مرحله برنامه‌ریزی مجاز است."),
                 ("E:H", None, _any_gate("AE"),
                  "گیت: ثبت رسید انبار پس از عبور حداقل یک پروژه از مرحله برنامه‌ریزی مجاز است.")],
    # ---- stage 6: مالی (پس از برنامه‌ریزی)
    "Cost_Lines": [("B:L", "C", _proj_gate("R"),
                    "گیت ۵: ثبت هزینه فقط پس از تکمیل برنامه‌ریزی تولید و انبارها مجاز است.")],
    "Costing": [("B:F", "B", _proj_gate("R"),
                 "گیت ۵: ثبت قیمت تمام‌شده فقط پس از تکمیل برنامه‌ریزی مجاز است."),
                ("S:S", "B", _proj_gate("R"),
                 "گیت ۵: تأیید قیمت تمام‌شده فقط پس از تکمیل برنامه‌ریزی مجاز است.")],
    "Budget": [("B:K", "B", _proj_gate("R"),
                "گیت ۵: تدوین بودجه فقط پس از تکمیل برنامه‌ریزی مجاز است.")],
    # ---- stage 7: فروش (پس از مالی)
    "Sales_Quotation": [("G:N", "B", _proj_gate("S"),
                         "گیت ۶: تعیین قیمت نهایی/حاشیه فقط پس از تأیید قیمت تمام‌شده توسط مالی مجاز است.")],
    "Payment_Terms": [
        ("C:I", "B",
         '=OR($B{r}="",IFERROR(INDEX(Gate_Status!$S$3:$S$20,MATCH(XLOOKUP($B{r},'
         'Sales_Quotation!$C$3:$C$10,Sales_Quotation!$B$3:$B$10,""),Gate_Status!$B$3:$B$20,0)),0)=1)',
         "گیت ۶: شرایط پرداخت فقط پس از تکمیل مرحله مالی مجاز است."),
    ],
    "Quotation_Versions": [
        ("B:H", "B",
         '=OR($B{r}="",IFERROR(INDEX(Gate_Status!$S$3:$S$20,MATCH(XLOOKUP($B{r},'
         'Sales_Quotation!$C$3:$C$10,Sales_Quotation!$B$3:$B$10,""),Gate_Status!$B$3:$B$20,0)),0)=1)',
         "گیت ۶: صدور نسخه پیشنهاد فقط پس از تکمیل مرحله مالی مجاز است."),
    ],
    # ---- stage 8: جمع‌بندی مدیریت پروژه (پس از قیمت‌گذاری فروش)
    "PMO_Summary": [("E:H", "B", _proj_gate("T"),
                     "گیت ۷: جمع‌بندی و اعلام به ارشد فقط پس از تکمیل قیمت‌گذاری فروش مجاز است.")],
    # ---- stage 9: مدیریت ارشد (پس از جمع‌بندی)
    "Management_Approval": [
        ("D:H", "C", _proj_gate("U"),
         "گیت ۸: ثبت تصمیم مدیریت ارشد فقط پس از ارسال گزارش جمع‌بندی مجاز است."),
    ],
}

GATE_DV_TITLE = "قفل گیت گردش کار"


def attach_gate_validations(ws, name, sh):
    cfg = GATE_DV.get(name)
    if not cfg:
        return
    last = max(sh.max_row(), 3) + 40   # extra room for future rows
    for cols, key, formula_tpl, msg in cfg:
        if ":" in cols:
            c1, c2 = cols.split(":")
            rng = f"{c1}3:{c2}{last}"
        else:
            rng = f"{cols}3:{cols}{last}"
        formula = formula_tpl.replace("{r}", "3")
        dv = DataValidation(type="custom", formula1=formula, allow_blank=True,
                            showErrorMessage=True)
        dv.error = msg + " (وضعیت گیت‌ها: شیت Gate_Status)"
        dv.errorTitle = GATE_DV_TITLE
        dv.prompt = msg
        dv.promptTitle = GATE_DV_TITLE
        dv.showInputMessage = False
        ws.add_data_validation(dv)
        dv.add(rng)


# --------------------------------------------------------------------------- #
# Gantt finishing
# --------------------------------------------------------------------------- #
def finish_gantt(ws, n_rows, first_period_col, n_periods):
    last_col = get_column_letter(first_period_col + n_periods - 1)
    last_row = 2 + n_rows
    bar_range = f"{get_column_letter(first_period_col)}3:{last_col}{last_row}"
    green = PatternFill("solid", fgColor="C6EFCE")
    blue = PatternFill("solid", fgColor="DDEBF7")
    red = PatternFill("solid", fgColor="FFC7CE")
    gray = PatternFill("solid", fgColor="EDEDED")
    ws.conditional_formatting.add(bar_range, FormulaRule(
        formula=[f'AND({get_column_letter(first_period_col)}3=1,$I3="Done")'], fill=green))
    ws.conditional_formatting.add(bar_range, FormulaRule(
        formula=[f'AND({get_column_letter(first_period_col)}3=1,$I3="In Progress")'], fill=blue))
    ws.conditional_formatting.add(bar_range, FormulaRule(
        formula=[f'AND({get_column_letter(first_period_col)}3=1,$I3="Late")'], fill=red))
    ws.conditional_formatting.add(bar_range, FormulaRule(
        formula=[f'AND({get_column_letter(first_period_col)}3=1,$I3="Planned")'], fill=gray))
    for k in range(n_periods):
        letter = get_column_letter(first_period_col + k)
        ws.column_dimensions[letter].width = 4.2
        for r in range(3, last_row + 1):
            pass
        ws.cell(row=2, column=first_period_col + k).number_format = "mm/dd"
        ws.cell(row=2, column=first_period_col + k).alignment = Alignment(
            horizontal="center", vertical="center", text_rotation=90)
    for r in range(3, last_row + 1):
        for k in range(n_periods):
            ws.cell(row=r, column=first_period_col + k).alignment = Alignment(
                horizontal="center", vertical="center")


# --------------------------------------------------------------------------- #
# Dashboard charts
# --------------------------------------------------------------------------- #
def add_dashboard_charts(ws, wb, chart_first, chart_last, n_projects):
    # chart 1 — number of projects completing each stage
    c1 = BarChart()
    c1.type = "col"
    c1.style = 10
    c1.title = "تعداد پروژه‌های تکمیل‌کننده هر مرحله"
    c1.y_axis.title = "تعداد پروژه"
    data = Reference(ws, min_col=2, min_row=chart_first, max_row=chart_last)
    cats = Reference(ws, min_col=1, min_row=chart_first, max_row=chart_last)
    c1.add_data(data, titles_from_data=False)
    c1.set_categories(cats)
    c1.height = 8.5
    c1.width = 24
    c1.legend = None
    ws.add_chart(c1, "O3")

    # chart 2 — completion percent per project (from Gate_Status)
    gs = wb["Gate_Status"]
    c2 = BarChart()
    c2.type = "bar"
    c2.style = 11
    c2.title = "درصد پیشروی پروژه‌ها در گردش کار"
    data2 = Reference(gs, min_col=23, min_row=3, max_row=2 + n_projects)  # W col
    cats2 = Reference(gs, min_col=2, min_row=3, max_row=2 + n_projects)   # B col
    c2.add_data(data2, titles_from_data=False)
    c2.set_categories(cats2)
    c2.height = 6.5
    c2.width = 14
    c2.legend = None
    ws.add_chart(c2, "O21")


# --------------------------------------------------------------------------- #
# Protection
# --------------------------------------------------------------------------- #
EDIT_RANGES = {
    # NOTE: ranges must cover INPUT columns only — formula columns stay locked,
    # otherwise a role could overwrite live calculations.
    "Orders": [("ROL-02", "B3:E20"), ("ROL-02", "G3:M20")],           # F=Total_Qty formula
    "Order_Lines": [("ROL-02", "B3:G20")],
    "Projects": [("ROL-03", "B3:N20")],                                # O=Overdue formula
    "Feasibility": [("ROL-03", "D3:L20")],                             # M.. flags are formulas
    "Schedule": [("ROL-03", "D3:E30"), ("ROL-03", "G3:I30"), ("ROL-03", "K3:K30")],  # F,J formulas
    "Engineering": [("ROL-04", "B3:M20")],
    "BOM_Header": [("ROL-04", "C3:J20")],
    "BOM_Detail": [("ROL-04", "B3:I60"), ("ROL-04", "L3:M60")],        # J,K formulas
    "Time_Study": [("ROL-04", "B3:I20")],                              # J..M formulas
    "Materials": [("ROL-04", "B3:I20")],
    "Equipment": [("ROL-04", "B3:J20")],
    "Project_Equipment": [("ROL-04", "B3:F20"), ("ROL-06", "G3:L20")],  # M,N formulas
    "Supplier_Quotes": [("ROL-06", "B3:H20")],
    "Supplier_Quote_Lines": [("ROL-06", "C3:R30"), ("ROL-06", "T3:T30")],  # S=Score_Total formula
    "Procurement": [("ROL-06", "C3:R20"), ("ROL-06", "T3:V20")],       # S,W formulas
    "Purchase_Prices": [("ROL-06", "B3:S30"), ("ROL-06", "T3:T30")],   # R=Expired formula
    "Planning": [("ROL-05", "R3:S30")],
    "Receipts": [("ROL-10", "B3:B30"), ("ROL-10", "E3:H30")],          # C,D formulas
    "Inventory": [("ROL-05", "C3:D30"), ("ROL-10", "C3:D30")],
    "Cost_Lines": [("ROL-07", "B3:I60"), ("ROL-07", "K3:L60")],        # J=Amount formula
    "Costing": [("ROL-07", "B3:F10"), ("ROL-07", "S3:S10")],           # G..R formulas
    "Budget": [("ROL-07", "B3:K10")],                                  # L..S formulas
    "Payment_Terms": [("ROL-02", "C3:I10")],                           # J,K formulas
    "Sales_Quotation": [("ROL-02", "G3:G10"), ("ROL-02", "I3:J10"), ("ROL-02", "M3:N10")],  # F,H,K,L formulas
    "Quotation_Versions": [("ROL-02", "B3:D10"), ("ROL-02", "H3:H10")],  # E,F,G formulas
    "PMO_Summary": [("ROL-03", "E3:H10")],                             # C,D,I,J,K formulas
    "Management_Approval": [("ROL-08", "D3:H10")],
    "Settings": [("ROL-09", "B3:B30")],
    "Customers": [("ROL-02", "B3:O20")],
    "Suppliers": [("ROL-06", "B3:L20")],
    "Products": [("ROL-04", "B3:G20")],
    "Product_Revisions": [("ROL-04", "B3:I20")],
}


def protect_all(wb):
    """Protect every sheet with the admin password; editing is only possible
    inside role-specific Protected Ranges (injected after save)."""
    for name in wb.sheetnames:
        ws = wb[name]
        sp = SheetProtection(sheet=True)
        sp.set_password(seed.ADMIN_PASSWORD)
        # allow selecting cells (locked or not); block structure-ish actions
        sp.formatCells = True
        sp.formatColumns = True
        sp.formatRows = True
        sp.insertColumns = True
        sp.insertRows = True
        sp.insertHyperlinks = True
        sp.deleteColumns = True
        sp.deleteRows = True
        sp.sort = True
        sp.autoFilter = False   # استفاده از فیلتر جداول مجاز باشد (داده را تغییر نمی‌دهد)
        sp.pivotTables = True
        sp.selectLockedCells = False
        sp.selectUnlockedCells = False
        ws.protection = sp
        color = TAB_COLORS.get(name, TAB_DEFAULT)
        ws.sheet_properties.tabColor = color


def _protected_ranges_xml(sheet_name):
    entries = EDIT_RANGES.get(sheet_name, [])
    if not entries:
        return ""
    parts = []
    for i, (role, sqref) in enumerate(entries, start=1):
        pwd = hash_password(ROLE_PASSWORDS[role])
        pname = f"{role}_{i}"
        parts.append(f'<protectedRange name="{pname}" sqref="{sqref}" password="{pwd}"/>')
    return "<protectedRanges>" + "".join(parts) + "</protectedRanges>"


def inject_protected_ranges(path, sheet_order):
    """Inject <protectedRanges> right after <sheetProtection> in each sheet XML."""
    tmp = path + ".tmp"
    # map sheet name -> sheetN.xml via workbook.xml + rels
    with zipfile.ZipFile(path, "r") as zin:
        wbxml = zin.read("xl/workbook.xml").decode("utf-8")
        relsxml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")

    rid_of_name = dict(re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml))
    target_of_rid = {}
    for m in re.finditer(r"<Relationship\b[^>]*/>", relsxml):
        tag = m.group(0)
        rid = re.search(r'Id="(rId\d+)"', tag)
        tgt = re.search(r'Target="([^"]+)"', tag)
        if rid and tgt:
            target_of_rid[rid.group(1)] = tgt.group(1)
    file_of_sheet = {}
    for name in sheet_order:
        rid = rid_of_name.get(name)
        if not rid:
            continue
        target = target_of_rid.get(rid, "")
        base = target.lstrip("/")
        if not base.startswith("xl/"):
            base = "xl/" + base.split("xl/")[-1]
        file_of_sheet[name] = base

    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in file_of_sheet.values():
                sheet_name = next(n for n, f in file_of_sheet.items() if f == item.filename)
                xml = data.decode("utf-8")
                pr = _protected_ranges_xml(sheet_name)
                if pr:
                    new, n = re.subn(r"(<sheetProtection[^>]*/>)",
                                     lambda m: m.group(1) + pr, xml, count=1)
                    if n == 0:
                        raise RuntimeError(f"sheetProtection not found in {item.filename}")
                    xml = new
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)
