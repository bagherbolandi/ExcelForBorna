#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
سامانه مدیریت پروژه برنا کامپوزیت ایرانیان — Workbook builder
ساخت فایل اکسل مدیریت پروژه با تقویم شمسی، گانت، داشبورد، BOM/بازرگانی و سطوح دسترسی.

Build:
    python3 build/build_workbook.py --out out/Project_Management_Borna.xlsx

References used in design: PMI PMBOK 7 + Practice Standard for Scheduling,
GAO-16-32G (Schedule Assessment Guide), HM Treasury Five-Case Model, ISO 21502.
"""
import argparse
import datetime as dt

import jdatetime
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter as gcl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.worksheet.datavalidation import DataValidation
import re as _re

# ----------------------------------------------------------------------------
# constants
# ----------------------------------------------------------------------------
FONT_NAME = "Tahoma"
EPOCH = dt.date(1899, 12, 30)

NAVY = "1F3864"; BLUE = "2F5597"; MIDBLUE = "8EAADB"; LBLUE = "D9E2F2"; PALEBLUE = "EEF3FB"
YELLOW = "FFF2CC"; GRAY = "F2F2F2"; SOFTGRAY = "EDEDED"; DARKTXT = "404040"; WHITE = "FFFFFF"
GREEN = "C6EFCE"; DGREEN = "006100"; LGREEN = "70AD47"
RED = "FFC7CE"; DRED = "9C0006"; BRED = "FF6B6B"
AMBER = "FFEB9C"; DAMBER = "9C6500"; LORANGE = "FFD966"
GRIDC = "BFBFBF"; PHASE = "DEEAF6"; BARPLAN = "BDD7EE"
INPUT_F = YELLOW; NOTE_F = "FBE5D6"

MONTHS_P = ["فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور", "مهر",
            "آبان", "آذر", "دی", "بهمن", "اسفند"]
WEEKDAYS_P = ["شنبه", "یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "پنجشنبه", "جمعه"]
WEEKEND_CODE = "0000001"          # تعطیلی: فقط جمعه
PERSIAN_DIGITS = "۰۱۲۳۴۵۶۷۸۹"

def ser(g):
    return (g - EPOCH).days

def j_of(g):
    j = jdatetime.date.fromgregorian(date=g)
    return (j.year, j.month, j.day)

def jt(g):
    y, m, d = j_of(g)
    return f"{y:04d}/{m:02d}/{d:02d}"

def j2g(s):
    s = _to_latin(str(s))
    y, m, d = [int(x) for x in s.split("/")]
    return jdatetime.date(y, m, d).togregorian()

def _to_latin(s):
    out = []
    for ch in s:
        i = PERSIAN_DIGITS.find(ch)
        out.append(str(i) if i >= 0 else ("0" if ch == "٠" and False else ch))
    # normalize arabic-indic digits too
    s2 = "".join(out)
    for i, ch in enumerate("٠١٢٣٤٥٦٧٨٩"):
        s2 = s2.replace(ch, str(i))
    s2 = s2.replace("،", ",").replace("٫", "/").replace("-", "/")
    return s2

def norm_expr(cell):
    """Formula fragment: normalize a user-typed jalali date string (digits + separators)."""
    e = f"TRIM({cell})"
    for i, ch in enumerate(PERSIAN_DIGITS + "٠١٢٣٤٥٦٧٨٩"):
        e = f'SUBSTITUTE({e},"{ch}","{i % 10}")'
    for sep in ["-", "‌", " ", "،"]:
        e = f'SUBSTITUTE({e},"{sep}","/")'
    e = e.replace('SUBSTITUTE(SUBSTITUTE(' , 'SUBSTITUTE(SUBSTITUTE(')
    return e

def j2s_formula(tcell):
    """jalali text cell -> excel serial (0 empty, -1 invalid)."""
    n = norm_expr(tcell)
    return (f'=IF({tcell}="",0,'
            f'IFERROR(INDEX(CalS,MATCH({n},CalJ,0)),-1))')

def s2j_formula(Scell):
    """serial cell -> jalali display text."""
    return (f'=IF({Scell}<=0,"",'
            f'IFERROR(INDEX(CalJ,MATCH(ROUND({Scell},0),CalS,0)),"خارج از تقویم"))')

# ----------------------------------------------------------------------------
# styling helpers
# ----------------------------------------------------------------------------
THIN = Side(style="thin", color=GRIDC)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

def font(sz=10, b=False, c="000000", i=False):
    return Font(name=FONT_NAME, size=sz, bold=b, color=c, italic=i)

def pfill(hexc):
    return PatternFill("solid", fgColor=hexc)

def al(h="right", wrap=True, v="center"):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def put(ws, r, c, v=None, *, f=None, fl=None, a=None, bd=True, num=None, unlock=False):
    cell = ws.cell(row=r, column=c)
    if v is not None:
        cell.value = v
    cell.font = f if f is not None else font()
    if fl is not None:
        cell.fill = fl if isinstance(fl, PatternFill) else pfill(fl)
    cell.alignment = a if a is not None else al()
    if bd:
        cell.border = BORDER
    if num is not None:
        cell.number_format = num
    if unlock:
        cell.protection = Protection(locked=False)
    return cell

def fill_range(ws, r1, r2, c1, c2, hexc, bd=True):
    f = pfill(hexc)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = f
            if bd:
                cell.border = BORDER

def merge(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

def banner(ws, title, desc, last_col):
    ws.sheet_view.showGridLines = False
    merge(ws, 1, 1, 1, last_col)
    put(ws, 1, 1, title, f=font(15, True, WHITE), fl=pfill(NAVY), a=al(), bd=False)
    for c in range(2, last_col + 1):
        put(ws, 1, c, None, fl=pfill(NAVY), bd=False)
    merge(ws, 2, 1, 2, last_col)
    put(ws, 2, 1, desc, f=font(9, False, DARKTXT), fl=pfill(PALEBLUE), bd=False)
    for c in range(2, last_col + 1):
        put(ws, 2, c, None, fl=pfill(PALEBLUE), bd=False)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 20

def section(ws, r, c1, c2, text):
    fill_range(ws, r, r, c1, c2, LORANGE, bd=True)
    merge(ws, r, c1, r, c2)
    put(ws, r, c1, text, f=font(11, True, NAVY), fl=pfill(LORANGE), a=al("right"))

def header_row(ws, r, c1, titles, height=32):
    for i, t in enumerate(titles):
        put(ws, r, c1 + i, t, f=font(9.5, True, WHITE), fl=pfill(BLUE), a=CENTER)
    ws.row_dimensions[r].height = height

def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w

def set_dv_list(ws, name, rng):
    dv = DataValidation(type="list", formula1=f"={name}", allow_blank=True, showErrorMessage=True,
                        errorTitle="مقدار نامعتبر", error="از فهرست کشویی انتخاب کنید")
    ws.add_data_validation(dv)
    dv.add(rng)

def protect(ws, pwd):
    p = ws.protection
    p.sheet = True
    p.formatCells = False; p.formatColumns = False; p.formatRows = False
    p.insertColumns = False; p.insertRows = False
    p.deleteColumns = False; p.deleteRows = False
    p.sort = False; p.autoFilter = False
    p.selectLockedCells = False; p.selectUnlockedCells = False
    p.objects = False; p.scenarios = True
    p.password = pwd

# ----------------------------------------------------------------------------
# global config for this build
# ----------------------------------------------------------------------------
PWD_USER = "Borna@1405"           # رمز محافظت پیش‌فرض همه شیت‌ها (در مستندات IT تغییر دهید)
PWD_STRUCTURE = "Borna-Admin-1405"  # رمز ساختار فایل (فقط مدیر پروژه / IT)

CAL_START = jdatetime.date(1403, 1, 1).togregorian()     # 2024-03-20
CAL_END   = jdatetime.date(1409, 1, 1).togregorian() - dt.timedelta(days=1)

GRID_YEAR = 1405                                     # سال تقویمی ستون‌های گانت
GRID_START = jdatetime.date(GRID_YEAR, 1, 1).togregorian()
GRID_END = jdatetime.date(GRID_YEAR + 1, 1, 1).togregorian() - dt.timedelta(days=1)
GRID_DAYS = (GRID_END - GRID_START).days + 1

SHEET_ORDER = ["راهنما", "ورود", "شناسنامه", "امکان‌سنجی", "گانت", "داشبورد",
               "مهندسی_BOM", "مهندسی_مدارک", "بازرگانی", "برنامه‌ریزی", "مالی", "فروش",
               "جمع‌بندی", "تصویب", "هزینه‌واقعی", "ریسک", "تغییرات", "تقویم", "تنظیمات"]

ROLE_CODES = ["INTAKE", "PM", "ENG", "TRD_IN", "TRD_EX", "PLAN", "FIN", "SALES", "CEO", "ADMIN"]
ROLE_NAMES = {"INTAKE": "واحد ثبت سفارش", "PM": "مدیر پروژه", "ENG": "مهندسی",
              "TRD_IN": "بازرگانی داخلی", "TRD_EX": "بازرگانی خارجی", "PLAN": "برنامه‌ریزی و انبار",
              "FIN": "مالی", "SALES": "فروش", "CEO": "مدیریت ارشد", "ADMIN": "مدیر سیستم"}

# دسترسی هر شیت برای هر نقش:  W=ویرایش، R=فقط‌نمایش، -=مخفی
ACCESS = {
    "راهنما":      {r: "R" for r in ROLE_CODES},
    "ورود":        {r: "R" for r in ROLE_CODES},
    "شناسنامه":    {"INTAKE": "W", "PM": "W", "ENG": "R", "FIN": "R", "SALES": "R", "CEO": "R",
                    "TRD_IN": "-", "TRD_EX": "-", "PLAN": "-", "ADMIN": "R"},
    "امکان‌سنجی":  {"PM": "W", "ENG": "W", "FIN": "W", "CEO": "R", "SALES": "R",
                    "INTAKE": "R", "TRD_IN": "R", "TRD_EX": "R", "PLAN": "-", "ADMIN": "R"},
    "گانت":        {"PM": "W", "ENG": "R", "TRD_IN": "R", "TRD_EX": "R", "PLAN": "R",
                    "FIN": "R", "SALES": "R", "CEO": "R", "INTAKE": "R", "ADMIN": "R"},
    "داشبورد":     {"INTAKE": "R", "PM": "R", "ENG": "R", "TRD_IN": "R", "TRD_EX": "R",
                    "PLAN": "R", "FIN": "R", "SALES": "R", "CEO": "R", "ADMIN": "R"},
    "مهندسی_BOM":  {"ENG": "W", "PM": "R", "TRD_IN": "R", "TRD_EX": "R", "PLAN": "R", "FIN": "R",
                    "INTAKE": "-", "SALES": "-", "CEO": "R", "ADMIN": "R"},
    "مهندسی_مدارک": {"ENG": "W", "PM": "R", "PLAN": "R", "INTAKE": "-", "TRD_IN": "-", "TRD_EX": "-",
                     "FIN": "-", "SALES": "-", "CEO": "-", "ADMIN": "R"},
    "بازرگانی":    {"TRD_IN": "W", "TRD_EX": "W", "PM": "R", "PLAN": "R", "FIN": "R",
                    "ENG": "R", "INTAKE": "-", "SALES": "-", "CEO": "R", "ADMIN": "R"},
    "برنامه‌ریزی": {"PLAN": "W", "PM": "R", "TRD_IN": "R", "TRD_EX": "R", "FIN": "R",
                    "ENG": "-", "INTAKE": "-", "SALES": "-", "CEO": "R", "ADMIN": "R"},
    "مالی":        {"FIN": "W", "PM": "R", "CEO": "R", "SALES": "R", "PLAN": "-",
                    "ENG": "-", "TRD_IN": "-", "TRD_EX": "-", "INTAKE": "-", "ADMIN": "R"},
    "فروش":        {"SALES": "W", "FIN": "R", "PM": "R", "CEO": "R", "INTAKE": "R",
                    "ENG": "-", "TRD_IN": "-", "TRD_EX": "-", "PLAN": "-", "ADMIN": "R"},
    "جمع‌بندی":    {"PM": "W", "ENG": "R", "TRD_IN": "R", "TRD_EX": "R", "PLAN": "R", "FIN": "R",
                    "SALES": "R", "INTAKE": "-", "CEO": "R", "ADMIN": "R"},
    "تصویب":       {"CEO": "W", "PM": "R", "INTAKE": "-", "ENG": "-", "TRD_IN": "-", "TRD_EX": "-",
                    "PLAN": "-", "FIN": "-", "SALES": "-", "ADMIN": "R"},
    "هزینه‌واقعی": {"FIN": "W", "PM": "W", "CEO": "R", "ENG": "R", "PLAN": "R",
                    "TRD_IN": "-", "TRD_EX": "-", "SALES": "-", "INTAKE": "-", "ADMIN": "R"},
    "ریسک":        {"PM": "W", "ENG": "R", "TRD_IN": "R", "TRD_EX": "R", "PLAN": "R", "FIN": "R",
                    "SALES": "R", "INTAKE": "R", "CEO": "R", "ADMIN": "R"},
    "تغییرات":     {"PM": "W", "ENG": "R", "TRD_IN": "R", "TRD_EX": "R", "PLAN": "R", "FIN": "R",
                    "SALES": "R", "INTAKE": "R", "CEO": "R", "ADMIN": "R"},
    "تقویم":       {r: "R" for r in ROLE_CODES},
    "تنظیمات":     {**{r: "-" for r in ROLE_CODES}, "ADMIN": "W", "PM": "R", "CEO": "R"},
}

TAB_COLORS = {"راهنما": "808080", "ورود": "1F3864", "شناسنامه": "ED7D31", "امکان‌سنجی": "FFC000",
              "گانت": "2E75B6", "داشبورد": "00B050", "مهندسی_BOM": "548235", "مهندسی_مدارک": "70AD47",
              "بازرگانی": "BF8F00", "برنامه‌ریزی": "7030A0", "مالی": "C00000", "فروش": "255E91",
              "جمع‌بندی": "4472C4", "تصویب": "002060", "هزینه‌واقعی": "943634", "ریسک": "D64541",
              "تغییرات": "8FAADC", "تقویم": "A9A9A9", "تنظیمات": "404040"}

UNITS = ["واحد ثبت سفارش", "مدیر پروژه", "مهندسی", "بازرگانی داخلی", "بازرگانی خارجی",
         "برنامه‌ریزی و انبار", "مالی", "فروش", "مدیریت ارشد"]

USERS = [  # (username, password, full name, role code)
    ("safa",  "Borna-Intake1", "صفا احمدی — واحد بازاریابی/ثبت سفارش", "INTAKE"),
    ("pm1",   "Borna-PM1",     "مهندس بهار رستمی — مدیر پروژه", "PM"),
    ("eng1",  "Borna-Eng1",    "مهندس کاویانی — سرپرست مهندسی", "ENG"),
    ("trd1",  "Borna-Trd1",    "آقای مرادی — بازرگانی داخلی", "TRD_IN"),
    ("trd2",  "Borna-Trd2",    "خانم نوری — بازرگانی خارجی", "TRD_EX"),
    ("plan1", "Borna-Plan1",   "مهندس تقوی — برنامه‌ریزی تولید و انبار", "PLAN"),
    ("fin1",  "Borna-Fin1",    "آقای شریفی — مالی", "FIN"),
    ("sales1","Borna-Sales1",  "آقای کمالی — فروش", "SALES"),
    ("ceo",   "Borna-Ceo1",    "مدیرعامل — مدیریت ارشد", "CEO"),
    ("admin", "Borna-Admin1",  "واحد فناوری اطلاعات", "ADMIN"),
]

# ----------------------------------------------------------------------------
# workbook skeleton
# ----------------------------------------------------------------------------
ap = argparse.ArgumentParser(description="Borna PM workbook builder")
ap.add_argument("--out", default="out/Project_Management_Borna.xlsx")
ap.add_argument("--year", type=int, default=GRID_YEAR, help="سال شمسی ستون‌های گانت")
ap.add_argument("--pwd-user", default=None, help="رمز قفل شیت‌ها (پیش‌فرض: نمونه)")
ap.add_argument("--pwd-structure", default=None, help="رمز قفل ساختار فایل")
args = ap.parse_args()
if args.pwd_user:
    PWD_USER = args.pwd_user
if args.pwd_structure:
    PWD_STRUCTURE = args.pwd_structure
if args.year != GRID_YEAR:
    GRID_START = jdatetime.date(args.year, 1, 1).togregorian()
    GRID_END = jdatetime.date(args.year + 1, 1, 1).togregorian() - dt.timedelta(days=1)
    GRID_DAYS = (GRID_END - GRID_START).days + 1

wb = Workbook()
wb.properties.creator = "Borna Composite — Project Control"
wb.properties.title = "سامانه مدیریت پروژه برنا کامپوزیت ایرانیان"
wb.calculation.fullCalcOnLoad = True

SH = {}
SH["راهنما"] = wb.active
SH["راهنما"].title = "راهنما"
for nm in SHEET_ORDER[1:]:
    SH[nm] = wb.create_sheet(nm)
for nm in SHEET_ORDER:
    ws = SH[nm]
    ws.sheet_view.rightToLeft = True
    ws.sheet_properties.tabColor = TAB_COLORS.get(nm)
# ----------------------------------------------------------------------------
# تقویم شمسی (jalali calendar sheet)
# ----------------------------------------------------------------------------
ws = SH["تقویم"]
CAL_HDR = 7
banner(ws, "تقویم شمسی",
       "جدول تبدیل تاریخ — مبنای سنجش فرم‌ها. بازه: " + jt(CAL_START) + " تا " + jt(CAL_END) +
       " | ستون B کلید ورودی همه فرم‌هاست؛ فقط تاریخ‌های موجود در این جدول معتبرند.", 12)
widths(ws, {"A": 12, "B": 14, "C": 7, "D": 6, "E": 6, "F": 12, "G": 5, "H": 7, "I": 7, "J": 9})
header_row(ws, CAL_HDR, 1, ["سریال میلادی", "تاریخ شمسی", "سال", "ماه", "روز", "روز هفته",
                            "جمعه", "تعطیل رسمی", "روز کاری", "ماه شمسی (متن)"])
r = CAL_HDR + 1
cur = CAL_START
cal_n = 0
while cur <= CAL_END:
    y, m, d = j_of(cur)
    wd = (cur.weekday() + 2) % 7          # 0=شنبه .. 5=پنجشنبه 6=جمعه
    put(ws, r, 1, cur, f=font(8), num="yyyy/mm/dd", fl=pfill(GRAY), a=CENTER)
    put(ws, r, 2, f"{y:04d}/{m:02d}/{d:02d}", f=font(8.5), a=CENTER)
    put(ws, r, 3, y, f=font(8), a=CENTER)
    put(ws, r, 4, m, f=font(8), a=CENTER)
    put(ws, r, 5, d, f=font(8), a=CENTER)
    put(ws, r, 6, WEEKDAYS_P[wd], f=font(8), a=CENTER)
    put(ws, r, 7, 1 if wd == 6 else 0, f=font(8), a=CENTER)
    put(ws, r, 8, f"=IF(COUNTIF(HolsNorm,$B{r})>0,1,0)", f=font(8), a=CENTER)
    put(ws, r, 9, f"=IF(OR($G{r}=1,$H{r}=1),0,1)", f=font(8, True), a=CENTER)
    put(ws, r, 10, f"{m:02d}/{y}", f=font(8), a=CENTER)
    r += 1
    cal_n += 1
    cur += dt.timedelta(days=1)
CAL_LAST = r - 1
ws.auto_filter.ref = f"A{CAL_HDR}:J{CAL_LAST}"
ws.freeze_panes = f"C{CAL_HDR + 1}"
ws.conditional_formatting.add(f"A{CAL_HDR + 1}:J{CAL_LAST}",
    FormulaRule(formula=[f"$G{CAL_HDR + 1}=1"], fill=pfill("FCE4E4")))
ws.conditional_formatting.add(f"A{CAL_HDR + 1}:J{CAL_LAST}",
    FormulaRule(formula=[f"$H{CAL_HDR + 1}=1"], fill=pfill("F4B6B6")))
CAL_S_RANGE = f"'تقویم'!$A${CAL_HDR + 1}:$A${CAL_LAST}"
CAL_J_RANGE = f"'تقویم'!$B${CAL_HDR + 1}:$B${CAL_LAST}"

# ----------------------------------------------------------------------------
# تنظیمات (settings, users, access matrix, holidays, dropdown lists)
# ----------------------------------------------------------------------------
ws = SH["تنظیمات"]
banner(ws, "تنظیمات سامانه", "پارامترهای عمومی، کاربران، ماتریس دسترسی و تعطیلات — "
       "ویرایش این برگه فقط با رمز ساختار فایل ممکن است. پس از دریافت فایل، حتماً رمز عبور همه کاربران را تغییر دهید.", 12)
widths(ws, {"A": 24, "B": 16, "C": 22, "D": 18, "E": 18, "F": 14, "G": 10, "H": 10, "I": 10,
            "J": 10, "K": 10, "L": 10, "M": 14})
P = 5
params = [("واحد پول گزارش‌ها", "میلیون تومان", False, None),
          ("نرخ تبدیل دلار (تومان)", 120000, True, "#,##0"),
          ("نرخ تبدیل یورو (تومان)", 132000, True, "#,##0"),
          ("حد نصاب امتیاز امکان‌سنجی (از ۱۰۰)", 70, True, "0"),
          ("قالب ورود تاریخ در کل فایل", "شمسی به صورت 1405/07/12 (فارسی یا لاتین)", False, None),
          ("کد کشور / تقویم مرجع", "IR — تقویم هجری شمسی (الگوریتم جلالی)", False, None)]
section(ws, P - 1, 1, 6, "پارامترهای عمومی")
for i, (lab, val, inp, num) in enumerate(params):
    put(ws, P + i, 1, lab, f=font(10, True), fl=pfill(GRAY))
    put(ws, P + i, 2, val, f=font(10), fl=pfill(YELLOW) if inp else pfill(WHITE), num=num, unlock=inp)
PARAM_CELL = {k: f"B{P + i}" for i, (k, *_ ) in enumerate(
    [("واحد پول"), ("نرخ دلار"), ("نرخ یورو"), ("حد نصاب"), ("قالب"), ("تقویم")])}
P_USR = 13
section(ws, P_USR - 1, 1, 6, "کاربران و رمز عبور (ورود به سامانه)")
header_row(ws, P_USR, 1, ["نام کاربری", "رمز عبور", "نام و نام خانوادگی", "کد نقش", "عنوان نقش", "بازنشانی اجباری رمز؟"])
for i, (u, p_, nm, role) in enumerate(USERS):
    rr = P_USR + 1 + i
    put(ws, rr, 1, u, f=font(9.5, True), fl=pfill(GRAY))
    put(ws, rr, 2, p_, f=font(9.5), fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 3, nm, f=font(9.5))
    put(ws, rr, 4, role, f=font(9.5), a=CENTER)
    put(ws, rr, 5, ROLE_NAMES[role], f=font(9.5))
    put(ws, rr, 6, "بله" if u != "admin" else "—", f=font(9), a=CENTER, fl=pfill(AMBER), unlock=True)
U_ROWS = (P_USR + 1, P_USR + len(USERS))
put(ws, U_ROWS[1] + 1, 1, "⚠ این رمزها نمونه‌اند؛ هنگام استقرار، توسط مدیر IT تغییر یابند. امنیت اکسل بازدارنده است، نه رمزنگاری داده — برای حفاظت سخت، ماتریس NTFS فایل‌ها را (اسناد docs) اعمال کنید.",
    f=font(8.5, False, DRED), fl=pfill(NOTE_F), bd=True)
merge(ws, U_ROWS[1] + 1, 1, U_ROWS[1] + 1, 6)
P_ACC = U_ROWS[1] + 3
section(ws, P_ACC - 1, 1, 12, "ماتریس دسترسی برگه‌ها (W=ویرایش، R=نمایش، -=مخفی) — توسط ماکرو اعمال می‌شود")
header_row(ws, P_ACC, 1, ["برگه"] + [ROLE_NAMES[r] for r in ROLE_CODES], height=34)
for i, sh in enumerate(SHEET_ORDER):
    rr = P_ACC + 1 + i
    put(ws, rr, 1, sh, f=font(9, True))
    for j, role in enumerate(ROLE_CODES):
        v = ACCESS[sh].get(role, "-")
        put(ws, rr, 2 + j, v, f=font(9, True), a=CENTER,
            fl=pfill({"W": GREEN, "R": "DDEBF7", "-": "E6E6E6"}[v]))
ACC_FIRST = P_ACC + 1
ACC_LAST = P_ACC + len(SHEET_ORDER)
P_HOL = ACC_LAST + 2
section(ws, P_HOL - 1, 1, 6, "تعطیلات رسمی (مبنای روزهای کاری و گانت)")
header_row(ws, P_HOL, 1, ["تاریخ شمسی (ورودی)", "عنوان تعطیلی", "نرمال‌سازی‌شده", "سریال میلادی (خودکار)"], height=22)
HOLS = [("1403/01/01", "nowruz"), ("1403/01/02","nowruz"),("1403/01/03","nowruz"),("1403/01/04","nowruz"),
        ("1403/01/06","Sizdah Bedar"),("1403/03/14","Demise of Imam Khomeini"),("1403/03/15","Uprising of 15 Khordad"),
        ("1403/11/22","Victory of Revolution"),("1403/12/11","Demise of Imam Khomeini (defacto)"),
        ("1403/12/12","Eid al-Fitr 1404")]
HOLIDAYS_TXT = [
    ("1403/01/01", "جشن نوروز"), ("1403/01/02", "جشن نوروز"), ("1403/01/03", "جشن نوروز"),
    ("1403/01/04", "جشن نوروز"), ("1403/01/06", "روز طبیعی / سیزده‌به‌در"),
    ("1403/03/14", "رحلت امام خمینی (ره)"), ("1403/03/15", "قیام ۱۵ خرداد"),
    ("1403/06/20", "اربعین حسینی"), ("1403/11/22", "پیروزی انقلاب اسلامی"),
    ("1403/12/11", "شهادت امام علی (ع)"), ("1403/12/12", "عید فطر"),
    ("1404/01/01", "جشن نوروز"), ("1404/01/02", "جشن نوروز"), ("1404/01/03", "جشن نوروز"),
    ("1404/01/04", "جشن نوروز"), ("1404/01/06", "روز طبیعت"),
    ("1404/03/14", "رحلت امام خمینی (ره)"), ("1404/03/15", "قیام ۱۵ خرداد"),
    ("1404/05/13", "مبعث پیامبر (ص)"), ("1404/11/22", "پیروزی انقلاب اسلامی"),
    ("1404/11/25", "روز جهانی قدس (تقریبی)"),
    ("1405/01/01", "جشن نوروز"), ("1405/01/02", "جشن نوروز"), ("1405/01/03", "جشن نوروز"),
    ("1405/01/04", "جشن نوروز"), ("1405/01/06", "روز طبیعت"),
    ("1405/03/14", "رحلت امام خمینی (ره)"), ("1405/03/15", "قیام ۱۵ خرداد"),
    ("1405/06/20", "تاسوعای حسینی (تقریبی — هر سال به‌روزرسانی شود)"),
    ("1405/06/21", "عاشورای حسینی (تقریبی)"),
    ("1405/11/22", "پیروزی انقلاب اسلامی"),
]
for i, (jd_, ttl) in enumerate(HOLIDAYS_TXT):
    rr = P_HOL + 1 + i
    put(ws, rr, 1, jd_, f=font(9.5), fl=pfill(YELLOW), num="@", a=CENTER, unlock=True)
    put(ws, rr, 2, ttl, f=font(9.5), fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 3, f"={norm_expr('A' + str(rr))}", f=font(8, False, WHITE), a=CENTER)
    put(ws, rr, 4, j2s_formula("$A" + str(rr)), f=font(8), num="yyyy/mm/dd", a=CENTER)
    ws.column_dimensions["C"].hidden = True
HOL_ROWS = (P_HOL + 1, P_HOL + 40)
for rr in range(P_HOL + 1 + len(HOLIDAYS_TXT), HOL_ROWS[1] + 1):
    put(ws, rr, 1, None, f=font(9.5), fl=pfill(YELLOW), num="@", a=CENTER, unlock=True)
    put(ws, rr, 2, None, f=font(9.5), fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 3, f"={norm_expr('A' + str(rr))}", f=font(8, False, WHITE), a=CENTER)
    put(ws, rr, 4, j2s_formula("$A" + str(rr)), f=font(8), num="yyyy/mm/dd", a=CENTER)
P_LIST = HOL_ROWS[1] + 2
section(ws, P_LIST - 1, 1, 6, "فهرست‌های کشویی (منبع اعتبارسنجی داده)")
lists = {
    "ListUnits":   ("واحد مسئول", UNITS),
    "ListIntake":  ("نوع ورودی", ["سفارش مشتری", "ایده داخلی", "طرح توسعه", "درخواست تغییر محصول"]),
    "ListProj":    ("وضعیت پرونده", ["در جریان بررسی", "مصوب مشروط", "مصوب", "معلّق", "رد شده"]),
    "ListDoc":     ("وضعیت مدرک", ["منتظر دریافت", "دریافت‌شده", "به‌روزرسانی لازم", "مردود"]),
    "ListRiskSt":  ("وضعیت ریسک", ["باز", "پایش", "بسته"]),
    "ListRiskStr": ("راهبرد ریسک", ["اجتناب", "کاهش/تسکین", "انتقال", "پذیرش"]),
    "ListCur":     ("واحد پول قلم", ["تومان", "دلار", "یورو"]),
    "ListPhase":   ("فاز پروژه", ["ثبت و امکان‌سنجی", "مهندسی", "تأمین", "تولید", "کیفیت و تست", "تحویل و پس از فروش"]),
    "ListChgSt":   ("وضعیت تغییر", ["درخواست جدید", "تأییدشده", "رد شده", "اجراشده"]),
    "ListTaskSt":  ("وضعیت تسک", ["آغازنشده", "در جریان", "تکمیل‌شده", "در تأخیر", "عقب‌افتاده"]),
    "score5":      ("امتیاز ۱ تا ۵", ["1", "2", "3", "4", "5"]),
}
row0 = P_LIST + 1
LIST_RANGES = {}
for i, (nm, (lab, vals)) in enumerate(lists.items()):
    rr = row0 + i
    put(ws, rr, 1, lab, f=font(9.5, True), fl=pfill(GRAY))
    for j, v in enumerate(vals):
        put(ws, rr, 2 + j, v, f=font(9), a=CENTER, fl=pfill(GRAY), bd=True)
    LIST_RANGES[nm] = (rr, 2, 2 + len(vals) - 1)
put(ws, row0 + len(lists) + 1, 1, "راهنمای ماکرو: پس از نصب ماکرو، برای ورود کلیدهای Alt+F8 و اجرای Borna_Login را بزنید. خروج: Borna_Logout. ورود اضطراری مدیر سیستم: Borna_Admin.",
    f=font(9, True, NAVY), fl=pfill(NOTE_F))
merge(ws, row0 + len(lists) + 1, 1, row0 + len(lists) + 1, 12)

# ----------------------------------------------------------------------------
# شناسنامه (intake / project charter)
# ----------------------------------------------------------------------------
ws = SH["شناسنامه"]
banner(ws, "شناسنامه پروژه — ثبت سفارش / ایده / طرح",
       "واحد دریافت‌کننده سفارش (بازاریابی/فروش/هر واحد) این برگه را تکمیل و به مدیریت پروژه ارجاع می‌کند. تاریخ‌ها شمسی.", 14)
widths(ws, {"A": 21, "B": 15, "C": 15, "D": 12, "E": 21, "F": 15, "G": 15, "H": 12,
            "I": 10, "J": 10, "K": 10, "L": 10, "M": 10, "N": 10})
section(ws, 4, 1, 14, "الف) مشخصات درخواست")
rows_a = [
    ("کد پروژه", "BC-P-1405-01", True, "@"), ("تاریخ ثبت درخواست", jt(jdatetime.date(1405, 4, 5).togregorian()), True, "@"),
    ("عنوان پروژه", "طراحی، ساخت و تحویل مخزن پلیمری FRP حجم 250m³", True, None),
    ("واحد/شخص ثبت‌کننده", "واحد بازاریابی", True, None),
    ("نوع ورودی", None, True, None), ("اولویت", "بالا", True, None),
    ("کارفرما / مشتری", "پتروشیمی نمونه — واحد نگهدداری", True, None),
    ("مدیر پروژه مسئول", "مهندس رستمی", True, None),
    ("تاریخ مورد نظر کارفرما", jt(jdatetime.date(1405, 11, 20).togregorian()), True, "@"),
    ("برآورد اولیه زمان (روز کاری)", 120, True, "0"),
    ("برآورد اولیه فروش (م.ت.)", 14800, True, "#,##0"),
    ("بودجه مجاز سازمان (م.ت.)", 12000, True, "#,##0"),
]
r0 = 5
for i, (lab, v, inp, num) in enumerate(rows_a):
    rr = r0 + i
    put(ws, rr, 1, lab, f=font(10, True), fl=pfill(GRAY))
    if num == "@":
        put(ws, rr, 2, v, f=font(10), fl=pfill(YELLOW), num="@", unlock=True, a=al("center"))
        merge(ws, rr, 2, rr, 3)
        put(ws, rr, 3, None, fl=pfill(YELLOW), unlock=True)
    else:
        put(ws, rr, 2, v, f=font(10), fl=pfill(YELLOW), unlock=True, num=num)
        merge(ws, rr, 2, rr, 3)
        put(ws, rr, 3, None, fl=pfill(YELLOW), unlock=True)
        if num: ws.cell(row=rr, column=2).number_format = num
    put(ws, rr, 5, "", fl=pfill(WHITE), bd=True)
put(ws, r0 + 4, 2, None, fl=pfill(YELLOW), unlock=True)  # نوع ورودی value
set_dv_list(ws, "ListIntake", f"B{r0 + 4}")
put(ws, r0 + 5, 2, None, fl=pfill(YELLOW), unlock=True)
# right column: auto status feed
auto = [("وضعیت پرونده (خودکار از برگه تصویب)", "='تصویب'!$B$12"),
        ("امتیاز امکان‌سنجی (خودکار)", "=IF('امکان‌سنجی'!$F$17=0,\"—\",'امکان‌سنجی'!$F$17)"),
        ("پیشرفت واقعی پروژه (خودکار)", "=IF('گانت'!$J$45=\"\",0,'گانت'!$J$45)"),
        ("پیشرفت برنامه‌ای امروز (خودکار)", "=IF('گانت'!$Q$45=\"\",0,'گانت'!$Q$45)")]
for i, (lab, f_) in enumerate(auto):
    rr = r0 + i
    put(ws, rr, 6, lab, f=font(9.5, True, NAVY), fl=pfill("DDEBF7"))
    put(ws, rr, 7, f_, f=font(10, True), fl=pfill("DDEBF7"), num="0.0%" if "پیشرفت" in lab else None)
    merge(ws, rr, 7, rr, 8)
section(ws, 18, 1, 14, "ب) شرح مختصر و مبنای درخواست")
put(ws, 19, 1, "شرح", f=font(10, True), fl=pfill(GRAY))
merge(ws, 19, 2, 21, 14)
put(ws, 19, 2, "کارفرما در نامه ۱۴۲/۰۴۰۱ درخواست ساخت مخزن FRP با حجم مفید ۲۵۰ مترمکعب برای نگهداری محلول اسیدی، همراه با تست هیدرواستاتیک، گواهی NDT و آموزش بهره‌بردار نموده است. "
      "تجهیزات چرخشی رزیون و قالب‌های بزرگ نیازمند تأمین خارجی است. تحویل حداکثر تا پایان دی ۱۴۰۵ مشمول وجه التزام روزی ۱۵۰ م.ت.",
    f=font(9.5), fl=pfill(YELLOW), unlock=True)
fill_range(ws, 20, 21, 2, 14, YELLOW)
section(ws, 23, 1, 14, "ج) گردش ارجاع (امضاها)")
header_row(ws, 24, 1, ["مرحله", "واحد / شخص", "تاریخ (شمسی)", "امضا", "یادداشت", "", "", "", "", "", "", "", "", ""], height=22)
flow = [("ثبت درخواست", "واحد بازاریابی"), ("تحویل به مدیریت پروژه", "دفتر مدیر پروژه"),
        ("بازگشت به واحد متقاضی (در صورت رد/اصلاح)", "واحد متقاضی")]
for i, (a, b) in enumerate(flow):
    rr = 25 + i
    put(ws, rr, 1, a, f=font(9.5)); put(ws, rr, 2, b, f=font(9.5))
    put(ws, rr, 3, None, fl=pfill(YELLOW), num="@", unlock=True, a=CENTER)
    put(ws, rr, 4, None, fl=pfill(YELLOW), unlock=True, a=CENTER)
    merge(ws, rr, 5, rr, 14)
    put(ws, rr, 5, None, fl=pfill(YELLOW), unlock=True)

# ----------------------------------------------------------------------------
# امکان‌سنجی (feasibility)
# ----------------------------------------------------------------------------
ws = SH["امکان‌سنجی"]
banner(ws, "امکان‌سنجی پروژه (Feasibility / Five-Case)",
       "تدوین توسط مدیریت پروژه با همکاری واحدهای تخصصی. مبنای ارزیابی: Five-Case خزانه بریتانیا + چک‌لیست PMI. امتیاز ۱ تا ۵؛ نتیجه در شناسنامه و داشبورد منعکس می‌شود.", 10)
widths(ws, {"A": 6, "B": 30, "C": 16, "D": 9, "E": 11, "F": 10, "G": 42, "H": 10, "I": 10, "J": 10})
header_row(ws, 6, 1, ["ردیف", "معیار", "دسته", "وزن ٪", "امتیاز (۱-۵)", "وزنی (از ۱۰۰)", "مستند / مبنای امتیاز"])
crit = [
    ("تقاضا و بازار هدف", "بازار", 15, 5, "نمونه‌کار مشابه ۱۴۰۴؛ مکاتبه پتروشیمی"),
    ("مزیت رقابتی و ریسک قیمت", "بازار", 10, 4, "تولید داخل؛ رقابت دو تأمین‌کننده وارداتی"),
    ("ظرفیت فنی و مهندسی (طراحی، قالب، لی‌آپ)", "فنی", 15, 4, "نیاز به قالب جدید؛ دانش فنی موجود"),
    ("تجهیزات و زیرساخت (تزریق رزی، پخت)", "فنی", 10, 3, "پمپ تزریق نیازمند تأمین خارجی"),
    ("زنجیره تأمین مواد (رزین/الیاف/شیمیایی)", "تأمین", 10, 4, "موجودی رزین کافی؛ الیاف از تأمین‌کننده معتبر"),
    ("ریسک اجرایی و HSE", "ریسک", 10, 3, "کار در ارتفاع و مواد شیمیایی؛ مجوز HSE لازم"),
    ("بازگشت سرمایه و نقدینگی", "مالی", 15, 4, "حاشیه سود هدف ≥۱۸٪؛ پیش‌پرداخت ۴۰٪"),
    ("انطباق با قوانین و استاندارد (ISO/ASME)", "قانونی", 5, 4, "استفاده از ASME RTP-1 برای مخازن FRP"),
    ("تعهم نیروی انسانی", "سازمانی", 5, 4, "تیم ۱۲ نفره؛ دو نفر پیمانکاری"),
    ("فشار زمانی تحویل", "زمان", 5, 3, "تحویل ۶ ماهه فشرده ولی امکان‌پذیر"),
]
for i, (nm, cat, w, sc, doc) in enumerate(crit):
    rr = 7 + i
    put(ws, rr, 1, i + 1, a=CENTER, f=font(9.5))
    put(ws, rr, 2, nm, f=font(9.5))
    put(ws, rr, 3, cat, a=CENTER, f=font(9.5))
    put(ws, rr, 4, w, a=CENTER, f=font(9.5), num="0", fl=pfill(GRAY))
    put(ws, rr, 5, sc, a=CENTER, f=font(10, True), fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 6, f"=ROUND($E{rr}*100*$D{rr}/100/5,1)", a=CENTER, f=font(9.5, True), num="0.0", fl=pfill(GRAY))
    put(ws, rr, 7, doc, f=font(8.5), fl=pfill(YELLOW), unlock=True)
set_dv_list(ws, "score5", "E7:E16")
put(ws, 17, 2, "جمع", f=font(10, True), fl=pfill(LBLUE))
put(ws, 17, 3, "", fl=pfill(LBLUE))
put(ws, 17, 4, "=SUM(D7:D16)", f=font(10, True), a=CENTER, num="0", fl=pfill(LBLUE))
put(ws, 17, 5, "", fl=pfill(LBLUE))
put(ws, 17, 6, "=ROUND(SUM(F7:F16),1)", f=font(12, True, NAVY), a=CENTER, num="0.0", fl=pfill(LBLUE))
put(ws, 17, 7, "وزن‌ها باید ۱۰۰ و امتیاز کل از ۱۰۰ باشد", f=font(8.5, False, DARKTXT), fl=pfill(LBLUE))
put(ws, 19, 2, "نتیجه", f=font(10, True), fl=pfill(GRAY))
merge(ws, 19, 3, 19, 7)
put(ws, 19, 3, '=IF($F$17=0,"امتیازدهی کامل نشده",IF($F$17>=ParamThresh,"✓ مجاز به ادامه — ورود به فاز برنامه‌ریزی و گانت","× نیازمند بازنگری طرح یا کاهش دامنه"))',
    f=font(11, True), fl=pfill("E2EFDA"))
put(ws, 20, 2, "تاریخ ارزیابی", f=font(10, True), fl=pfill(GRAY))
put(ws, 20, 3, jt(jdatetime.date(1405, 4, 22).togregorian()), num="@", f=font(9.5), a=CENTER, fl=pfill(YELLOW), unlock=True)
put(ws, 20, 4, "بازنگری (در میانه پروژه)", f=font(9.5), fl=pfill(GRAY))
merge(ws, 20, 5, 20, 7)
put(ws, 20, 5, "در صورت تغییر دامنه، امتیازدهی و گانت الزاماً بازنگری می‌شود (مسئول: مدیر پروژه).", f=font(8.5), fl=pfill(NOTE_F))
for cc in (5, 6, 7):
    ws.cell(row=20, column=cc).fill = pfill(NOTE_F)
ws.conditional_formatting.add("F17", FormulaRule(formula=["$F$17>=ParamThresh"], fill=pfill(GREEN)))
ws.conditional_formatting.add("F17", FormulaRule(formula=["AND($F$17>0,$F$17<ParamThresh)"], fill=pfill(RED)))
ws.conditional_formatting.add("D17", FormulaRule(formula=["$D$17<>100"], fill=pfill(RED)))
# ----------------------------------------------------------------------------
# گانت (schedule / Gantt)
# ----------------------------------------------------------------------------
ws = SH["گانت"]
NT_FIRST, NT_LAST, NT_TOT = 8, 44, 45          # task rows + total row
GC = 20                                         # first grid column (T)
GLAST = GC + GRID_DAYS - 1
LC = GC + 8                                     # banner span
banner(ws, "گانت پروژه — برنامه‌ریزی و کنترل زمان",
       "ورودی ستون‌های زرد: تاریخ شمسی (قالب 1405/07/01)، درصد پیشرفت و یادداشت. محاسبات سریال/وضعیت خودکار است. "
       "ستون‌های گانت بر اساس تقویم تنظیم می‌شوند؛ برای سال دیگر اسکریپت سازنده را با --year اجرا کنید.", LC)
widths(ws, {"A": 4.5, "B": 8, "C": 40, "D": 14, "E": 8, "F": 11.5, "G": 11.5, "H": 11.5, "I": 11.5,
            "J": 8.5, "K": 3, "L": 3, "M": 3, "N": 3, "O": 6.5, "P": 6.5, "Q": 8, "R": 13, "S": 24})
for c in range(GC, GLAST + 1):
    ws.column_dimensions[gcl(c)].width = 2.35
ws.row_dimensions[11].height = 13.5
ws.row_dimensions[12].height = 10.5

kpi = [("امروز (شمسی)", '=IFERROR(INDEX(CalJ,MATCH(TODAY(),CalS,0)),"خارج از بازه تقویم")'),
       ("پیشرفت واقعی", f"=$J${NT_TOT}"),
       ("پیشرفت برنامه‌ای امروز", f"=$Q${NT_TOT}"),
       ("فعالیت در تأخیر/عقب‌افتاده", f'=COUNTIF($R${NT_FIRST}:$R${NT_LAST},"در تأخیر")+COUNTIF($R${NT_FIRST}:$R${NT_LAST},"عقب‌افتاده")')]
put(ws, 3, 2, "شاخص‌ها:", f=font(9, True, NAVY), bd=False)
col = 3
for lab, f_ in kpi:
    put(ws, 3, col, lab, f=font(8, False, DARKTXT), fl=pfill(GRAY), a=CENTER)
    merge(ws, 3, col, 3, col + 1)
    put(ws, 4, col, f_, f=font(10, True, NAVY), fl=pfill("DDEBF7"), a=CENTER,
        num="0.0%" if "پیشرفت" in lab else None)
    merge(ws, 4, col, 4, col + 1)
    for cc in (col, col + 1):
        ws.cell(row=3, column=cc).fill = pfill(GRAY)
        ws.cell(row=4, column=cc).fill = pfill("DDEBF7")
    col += 3

hdrs = ["ردیف", "کد WBS", "شرح فعالیت / بسته کاری", "واحد مسئول", "پیش‌نیاز",
        "شروع مبنایی", "پایان مبنایی", "شروع واقعی", "پایان واقعی/پیش‌بینی", "پیشرفت ٪",
        "sF", "sG", "sH", "sI", "مدت (WD)", "باقی‌مانده", "برنامه‌ای امروز", "وضعیت", "یادداشت"]
header_row(ws, NT_FIRST - 1, 1, hdrs, height=40)
for c in range(12, 15):
    ws.cell(row=NT_FIRST - 1, column=c).font = font(6, True, "C0C0C0")

for c in range(GC, GLAST + 1):
    ws.cell(row=NT_FIRST - 1, column=c)
    ws.column_dimensions[gcl(c)]

# month band (row 6) + day numbers (row 7) + hidden serial row (row... ) — use rows: 6 band, 7 days
ws.row_dimensions[6].height = 15
ws.row_dimensions[7].height = 12
SER_ROW = 7  # day numbers live on row 7; serials live on hidden row 6?  -- handled below with dedicated rows:

# serials row = row 5 (kept tiny), months band row 4, days row 5? finalize: row 4=months, row 5=days, row 6=serial(hidden)
# we already painted header at row NT_FIRST-1 (=7). Rows 4,5,6 belong to the calendar band.
# clear placeholder row values written by kpi merges on row 3/4 cols 3..8 if overlapping: rows 4 used only cols C..; band starts col GC so safe.
months_days = []
cur = GRID_START
while cur <= GRID_END:
    y, m, d = j_of(cur)
    months_days.append((cur, y, m, d))
    cur += dt.timedelta(days=1)
# month bands
band_specs = []
start_c = GC
prev_ym = None
for i, (gday, y, m, d) in enumerate(months_days):
    ym = (y, m)
    c = GC + i
    if prev_ym is not None and ym != prev_ym:
        band_specs.append((start_c, c - 1, prev_ym))
        start_c = c
    prev_ym = ym
    # day numbers row 5
    put(ws, 5, c, d, f=font(6.5, b=(d == 1)),
        a=CENTER, bd=False, fl=pfill(SOFTGRAY) if (gday.weekday() + 2) % 7 == 6 else None)
    # serial row 6 hidden height
    put(ws, 6, c, ser(gday), f=font(1, False, WHITE), bd=False, a=CENTER)
    # header cells fill
    put(ws, 4, c, None, bd=False)
band_specs.append((start_c, GC + len(months_days) - 1, prev_ym))
for (c1, c2, (y, m)) in band_specs:
    merge(ws, 4, c1, 4, c2)
    alt = "2F5597" if (m % 2) else "1F3864"
    put(ws, 4, c1, MONTHS_P[m - 1], f=font(7.5, True, WHITE), fl=pfill(alt), a=CENTER)
    for cc in range(c1, c2 + 1):
        ws.cell(row=4, column=cc).fill = pfill(alt)
ws.row_dimensions[6].height = 2.5
# month first-day separators on row 5 (thick left border)
for (c1, c2, _ym) in band_specs:
    cell = ws.cell(row=5, column=c1)
    cell.border = Border(left=Side(style="medium", color=NAVY), right=THIN, top=THIN, bottom=THIN)
for i in range(GC, GLAST + 1):
    cc = ws.cell(row=4, column=i)
    cc.border = Border(left=Side(style="medium", color=NAVY) if any(i == s for s, e, ym in band_specs) else None,
                       bottom=None)

G = "گانت"
grng = f"{gcl(GC)}{NT_FIRST}:{gcl(GLAST)}{NT_LAST}"
def gf(inner):
    """replace leading grid cell ref marker @ with proper relative column on first col of range"""
    return inner.replace("@", gcl(GC) + "$6")
grid_rules = [
    ("=AND($M{r}>0,@>=$M{r},@<=$M{r}+($N{r}-$M{r})*$J{r})", LGREEN, True),          # completed part of actual
    ("=AND($M{r}>0,$N{r}>$L{r},@>$L{r},@<=$N{r})", "C00000", True),  # overdue tail vs baseline
    ("=AND($J{r}<1,$M{r}>0,@>$M{r}+($N{r}-$M{r})*$J{r},@<=MAX($N{r},$L{r}))", LORANGE, True),  # remaining forecast
    ("=AND($M{r}<=0,@>=$K{r},@<=$L{r},$K{r}>0)", BARPLAN, True),                      # not-started baseline
    ("=@=TODAY()", "F8CBAD", False),                                                 # today column
    ("=WEEKDAY(@,2)=6", SOFTGRAY, False),                                            # fridays
]
for fml, hexc, stop in grid_rules:
    ws.conditional_formatting.add(grng, FormulaRule(formula=[fml.format(r=NT_FIRST).replace("$J{", "$J{").replace("@", gcl(GC) + "$6")],
        fill=pfill(hexc), stopIfTrue=stop))
# header rules: today in day-number row
ws.conditional_formatting.add(f"{gcl(GC)}5:{gcl(GLAST)}5",
    FormulaRule(formula=[f"{gcl(GC)}$6=TODAY()"], font=font(8, True, DRED), fill=pfill("FFD966")))
# status colors
rr = f"R{NT_FIRST}:R{NT_LAST}"
ws.conditional_formatting.add(rr, FormulaRule(formula=[f'ISNUMBER(SEARCH("تکمیل",$R{NT_FIRST}))'], fill=pfill(GREEN), font=font(9, True, DGREEN)))
ws.conditional_formatting.add(rr, FormulaRule(formula=[f'OR(ISNUMBER(SEARCH("تأخیر",$R{NT_FIRST})),ISNUMBER(SEARCH("عقب",$R{NT_FIRST})),ISNUMBER(SEARCH("نامعتبر",$R{NT_FIRST})))'], fill=pfill(RED), font=font(9, True, DRED)))
ws.conditional_formatting.add(rr, FormulaRule(formula=[f'ISNUMBER(SEARCH("جریان",$R{NT_FIRST}))'], fill=pfill(AMBER), font=font(9, True, DAMBER)))

# ---------------- demo tasks ----------------
# (code, title, unit, pred, F_start, G_end, H_actS, I_actE, progress, note, is_phase)
DEMO = [
    ("F0", "فاز ۰ — ورودی، امکان‌سنجی و تصویب", "", "", "", "", "", "", None, "", True),
    ("F0.1", "ثبت سفارش/ایده و صدور شناسنامه", "واحد ثبت سفارش", "—", "1405/04/01", "1405/04/05", "1405/04/01", "1405/04/05", 1.0, "نمونه‌کار مشابه: مخزن ۱۰۰m³ پارس‌ودا", False),
    ("F0.2", "تهیه پیش‌نویس گانت و ساختار WBS توسط مدیر پروژه", "مدیر پروژه", "F0.1", "1405/04/06", "1405/04/10", "1405/04/06", "1405/04/11", 1.0, "WBS ۴ سطحی، مبنای ۵ Case", False),
    ("F0.3", "امکان‌سنجی و امتیازدهی معیارها", "مدیر پروژه", "F0.2", "1405/04/12", "1405/04/22", "1405/04/13", "1405/04/22", 1.0, "همکاری مهندسی/مالی — برگه امکان‌سنجی", False),
    ("F0.4", "جلسه کمیته تصویب (گیت ۱)", "مدیریت ارشد", "F0.3", "1405/04/25", "1405/04/26", "1405/04/25", "1405/04/26", 1.0, "مصوب مشروط: بازنگری پس از BOM نهایی", False),
    ("F1", "فاز ۱ — مهندسی", "", "", "", "", "", "", None, "", True),
    ("F1.1", "دریافت نقشه‌ها و مشخصات فنی از کارفرما", "مهندسی", "F0.4", "1405/05/01", "1405/05/08", "1405/05/01", "1405/05/09", 1.0, "۱۴ نقشه / rev B", False),
    ("F1.2", "تدوین Test Plan و روش‌های آزمون", "مهندسی", "F1.1", "1405/05/09", "1405/05/20", "1405/05/10", "1405/05/22", 1.0, "ASTM D2992 / ASME RTP-1", False),
    ("F1.3", "تهیه BOM اولیه و محاسبه نفرساعت", "مهندسی", "F1.2", "1405/05/21", "1405/05/30", "1405/05/21", "1405/05/30", 1.0, "BOM rev0 ارسال به بازرگانی", False),
    ("F1.4", "طراحی قالب و نهایی‌سازی BOM", "مهندسی", "F1.3", "1405/05/31", "1405/06/10", "1405/05/31", "1405/06/14", 0.9, "قالب اسپیل ۲۵۰m³ — ۹۰٪", False),
    ("F2", "فاز ۲ — تأمین (بازرگانی)", "", "", "", "", "", "", None, "", True),
    ("F2.1", "استعلام و قیمت‌گذاری داخلی (رزین، الیاف، شیمیایی)", "بازرگانی داخلی", "F1.3", "1405/06/03", "1405/06/16", "1405/06/03", "1405/06/20", 0.75, "۳ تأمین‌کننده رزین ایزو", False),
    ("F2.2", "استعلام خارجی (پمپ رزیون، تجهیزات تزریق)", "بازرگانی خارجی", "F1.3", "1405/06/04", "1405/06/22", "1405/06/06", "1405/07/15", 0.45, "تأخیر در L/C — ریسک R-02 فعال", False),
    ("F2.3", "ارزیابی و انتخاب تأمین‌کننده / قرارداد", "بازرگانی داخلی", "F2.1", "1405/06/23", "1405/06/30", "1405/06/24", "1405/07/04", 0.6, "ماتریس ارزیابی تکمیل شد", False),
    ("F2.4", "صدور سفارش خرید و تخصیص ارز/اسناد", "مالی", "F2.3", "1405/06/28", "1405/07/12", "1405/07/01", "", 0.3, "منتظر تخصیص ارز نیمایی", False),
    ("F2.5", "پایش تحویل و رسید انبار اقلام", "برنامه‌ریزی و انبار", "F2.4", "1405/07/01", "1405/07/25", "1405/07/01", "", 0.1, "رسید الیاف در سامانه", False),
    ("F3", "فاز ۳ — تولید", "", "", "", "", "", "", None, "", True),
    ("F3.1", "ساخت قالب و ابزارها", "برنامه‌ریزی و انبار", "F1.4", "1405/07/05", "1405/07/28", "", "", None, "پیمانکشی — پیش‌فاکتور تأیید شد", False),
    ("F3.2", "لایه‌نشانی بدنه و پخت رزین", "برنامه‌ریزی و انبار", "F2.5", "1405/08/01", "1405/08/25", "", "", None, "", False),
    ("F3.3", "رینگ‌های تقویتی و نازل‌ها", "برنامه‌ریزی و انبار", "F2.5", "1405/08/20", "1405/09/10", "", "", None, "", False),
    ("F3.4", "مونتاژ نهایی و اتصالات", "برنامه‌ریزی و انبار", "F3.3", "1405/09/11", "1405/09/25", "", "", None, "", False),
    ("F4", "فاز ۴ — کیفیت و تست", "", "", "", "", "", "", None, "", True),
    ("F4.1", "تست هیدرواستاتیک و NDT", "مهندسی", "F3.4", "1405/09/26", "1405/10/06", "", "", None, "ناظر ثالث مشتری", False),
    ("F4.2", "ترمیم، ژل‌کوت نهایی و بسته‌بندی", "برنامه‌ریزی و انبار", "F4.1", "1405/10/07", "1405/10/18", "", "", None, "", False),
    ("F5", "فاز ۵ — مالی، فروش و تحویل", "", "", "", "", "", "", None, "", True),
    ("F5.1", "تدوین بهای تمام‌شده و گزارش انحراف بودجه (چرخه ماهانه)", "مالی", "F2.1", "1405/07/01", "1405/11/20", "1405/07/01", "", 0.15, "به‌روزرسانی ماهانه", False),
    ("F5.2", "تعیین قیمت نهایی فروش و شرایط پرداخت", "فروش", "F1.3", "1405/06/15", "1405/06/28", "1405/06/16", "", 0.5, "بر اساس بهای تقریبی rev1", False),
    ("F5.3", "جمع‌بندی مدیر پروژه و تصویب نهایی مدیریت ارشد", "مدیر پروژه", "F5.2", "1405/07/02", "1405/07/06", "1405/07/01", "1405/07/01", 1.0, "مصوب مشروط ثبت شد (برگه تصویب)", False),
    ("F5.4", "حمل و تحویل کارگاهی", "برنامه‌ریزی و انبار", "F4.2", "1405/10/20", "1405/10/31", "", "", None, "", False),
    ("F5.5", "نصب، راه‌اندازی و آموزش بهره‌بردار", "مدیر پروژه", "F5.4", "1405/11/03", "1405/11/12", "", "", None, "", False),
    ("F5.6", "صورت‌جلسه تحویل و صدور گارانتی ۱۸ ماهه", "مدیر پروژه", "F5.5", "1405/11/13", "1405/11/15", "", "", None, "", False),
]
DEMO_ROWS = {c[0]: NT_FIRST + i for i, c in enumerate(DEMO)}
assert len(DEMO) <= NT_LAST - NT_FIRST + 1, "demo tasks exceed task rows"
for i, (code, title, unit, pred, fs, ge, hs, he, prog, note, is_phase) in enumerate(DEMO):
    rr = NT_FIRST + i
    rowfill = PHASE if is_phase else None
    put(ws, rr, 1, i + 1, a=CENTER, f=font(8.5, False, DARKTXT), fl=rowfill)
    put(ws, rr, 2, code, f=font(9, is_phase, NAVY if is_phase else "000000"), a=CENTER, fl=rowfill)
    put(ws, rr, 3, title, f=font(9.5 if is_phase else 9, is_phase), fl=rowfill)
    put(ws, rr, 4, unit if unit else None, f=font(9), a=CENTER, fl=pfill(YELLOW) if not is_phase else rowfill, unlock=not is_phase)
    put(ws, rr, 5, pred, f=font(8.5), a=CENTER, fl=pfill(YELLOW) if not is_phase else rowfill, unlock=not is_phase)
    for cc, val in ((6, fs), (7, ge), (8, hs), (9, he)):
        put(ws, rr, cc, val if val else None, f=font(9), num="@", a=CENTER,
            fl=pfill(YELLOW) if not is_phase else rowfill, unlock=not is_phase)
    put(ws, rr, 10, prog if prog is not None else None, f=font(9, True), num="0%", a=CENTER,
        fl=pfill(YELLOW) if not is_phase else rowfill, unlock=not is_phase)
    # hidden serials + derived (only for task rows and blank spare rows)
    for cc, src in ((11, 6), (12, 7), (13, 8), (14, 9)):
        colL = gcl(src)
        put(ws, rr, cc, j2s_formula(f"${colL}{rr}"), f=font(6), num=";;;", a=CENTER, fl=rowfill)
    put(ws, rr, 15, f"=IF(AND($K{rr}>0,$L{rr}>0),NETWORKDAYS.INTL($K{rr},$L{rr},\"{WEEKEND_CODE}\",Hols),0)", f=font(8.5), a=CENTER, num="0", fl=rowfill)
    put(ws, rr, 16, f"=IF($L{rr}<=0,\"—\",IF($J{rr}>=1,\"\",IF(TODAY()>$L{rr},0,NETWORKDAYS.INTL(MAX(TODAY(),$K{rr}),$L{rr},\"{WEEKEND_CODE}\",Hols))))", f=font(8.5), a=CENTER, num="0", fl=rowfill)
    put(ws, rr, 17, f"=IF(OR($K{rr}<=0,$L{rr}<=0),0,MEDIAN(0,(TODAY()-$K{rr}+1)/($L{rr}-$K{rr}+1),1))", f=font(6), num=";;;", a=CENTER, fl=rowfill)
    put(ws, rr, 18, f'=IF(OR($F{rr}="",$G{rr}=""),"",IF($K{rr}<0,"⚠ تاریخ نامعتبر",IF($M{rr}<=0,IF(TODAY()>$L{rr},"عقب‌افتاده","آغازنشده"),IF($J{rr}>=1,IF(OR(AND($N{rr}>$L{rr},$N{rr}>0),TODAY()>$L{rr}),"تکمیل با تأخیر","تکمیل‌شده"),IF(OR(AND($N{rr}>$L{rr},$N{rr}>0),TODAY()>$L{rr}),"در تأخیر","در جریان")))))',
        f=font(8.5, True), a=CENTER, fl=rowfill)
    put(ws, rr, 19, note, f=font(8.5, False, DARKTXT), fl=pfill(YELLOW) if not is_phase else rowfill, unlock=not is_phase)
    ws.row_dimensions[rr].height = 15.5
for rr in range(NT_FIRST + len(DEMO), NT_LAST + 1):
    for cc in range(1, 20):
        if cc in (11, 12, 13, 14, 17):
            continue
        fl_ = pfill(YELLOW) if cc in (4, 5, 6, 7, 8, 9, 10, 19) else pfill(WHITE)
        put(ws, rr, cc, None, f=font(9), fl=fl_, unlock=cc in (4, 5, 6, 7, 8, 9, 10, 19), num="@" if cc in (6, 7, 8, 9) else None)
    for cc, src in ((11, 6), (12, 7), (13, 8), (14, 9)):
        put(ws, rr, cc, j2s_formula(f"${gcl(src)}{rr}"), f=font(6), num=";;;", a=CENTER)
    put(ws, rr, 15, f"=IF(AND($K{rr}>0,$L{rr}>0),NETWORKDAYS.INTL($K{rr},$L{rr},\"{WEEKEND_CODE}\",Hols),0)", f=font(8.5), a=CENTER, num="0")
    put(ws, rr, 16, f"=IF($L{rr}<=0,\"—\",IF($J{rr}>=1,\"\",IF(TODAY()>$L{rr},0,NETWORKDAYS.INTL(MAX(TODAY(),$K{rr}),$L{rr},\"{WEEKEND_CODE}\",Hols))))", f=font(8.5), a=CENTER, num="0")
    put(ws, rr, 17, f"=IF(OR($K{rr}<=0,$L{rr}<=0),0,MEDIAN(0,(TODAY()-$K{rr}+1)/($L{rr}-$K{rr}+1),1))", f=font(6), num=";;;", a=CENTER)
    put(ws, rr, 18, f'=IF(OR($F{rr}="",$G{rr}=""),"",IF($K{rr}<0,"⚠ تاریخ نامعتبر",IF($M{rr}<=0,IF(TODAY()>$L{rr},"عقب‌افتاده","آغازنشده"),IF($J{rr}>=1,IF(OR(AND($N{rr}>$L{rr},$N{rr}>0),TODAY()>$L{rr}),"تکمیل با تأخیر","تکمیل‌شده"),IF(OR(AND($N{rr}>$L{rr},$N{rr}>0),TODAY()>$L{rr}),"در تأخیر","در جریان")))))',
        f=font(8.5, True), a=CENTER)
    ws.row_dimensions[rr].height = 15.5
# total row
put(ws, NT_TOT, 3, "جمع کل (وزن‌دهی با مدت روزکاری)", f=font(10, True), fl=pfill(LBLUE))
for cc in (1, 2): put(ws, NT_TOT, cc, None, fl=pfill(LBLUE))
for cc in range(4, 20): put(ws, NT_TOT, cc, None, fl=pfill(LBLUE))
put(ws, NT_TOT, 10, f"=IF(SUM($O${NT_FIRST}:$O${NT_LAST})=0,\"\",SUMPRODUCT($O${NT_FIRST}:$O${NT_LAST},$J${NT_FIRST}:$J${NT_LAST})/SUM($O${NT_FIRST}:$O${NT_LAST}))", f=font(11, True, NAVY), a=CENTER, num="0.0%", fl=pfill(LBLUE))
put(ws, NT_TOT, 15, f"=SUM($O${NT_FIRST}:$O${NT_LAST})", f=font(10, True), a=CENTER, num="0", fl=pfill(LBLUE))
put(ws, NT_TOT, 17, f"=IF(SUM($O${NT_FIRST}:$O${NT_LAST})=0,\"\",SUMPRODUCT($O${NT_FIRST}:$O${NT_LAST},$Q${NT_FIRST}:$Q${NT_LAST})/SUM($O${NT_FIRST}:$O${NT_LAST}))", f=font(11, True, NAVY), a=CENTER, num="0.0%", fl=pfill(LBLUE))
# progress numeric guard: empty J cells treated as 0 in SUMPRODUCT (text!) — J blank rows hold None => SUMPRODUCT errors.
# ensure J inputs default 0 text-free: numeric validation below + note; spare rows J left blank intentionally → wrap SUMPRODUCT with N():
ws.cell(row=NT_TOT, column=10).value = f"=IF(SUM($O${NT_FIRST}:$O${NT_LAST})=0,\"\",SUMPRODUCT($O${NT_FIRST}:$O${NT_LAST},$J${NT_FIRST}:$J${NT_LAST})/SUM($O${NT_FIRST}:$O${NT_LAST}))"
ws.cell(row=NT_TOT, column=17).value = f"=IF(SUM($O${NT_FIRST}:$O${NT_LAST})=0,\"\",SUMPRODUCT($O${NT_FIRST}:$O${NT_LAST},$Q${NT_FIRST}:$Q${NT_LAST})/SUM($O${NT_FIRST}:$O${NT_LAST}))"
# legend
lr = NT_TOT + 2
legend = [("سبز", "پیشرفت واقعی انجام‌شده", LGREEN), ("نارنجی", "باقی‌مانده تا پایان پیش‌بینی", LORANGE),
          ("قرمز", "تأخیر نسبت به مبنای", "C00000"), ("آبی", "برنامه مبنایی (آغازنشده)", BARPLAN)]
put(ws, lr, 2, "راهنمای رنگ:", f=font(9, True), bd=False)
for i, (nm, desc, hexc) in enumerate(legend):
    put(ws, lr, 3 + i * 4, "", fl=pfill(hexc))
    put(ws, lr, 4 + i * 4, desc, f=font(8.5, False, DARKTXT), bd=False)
# DV + protections ranges
dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="100",
                    errorTitle="درصد نامعتبر", error="عدد بین ۰ تا ۱۰۰ وارد کنید", showErrorMessage=True)
ws.add_data_validation(dv); dv.add(f"J{NT_FIRST}:J{NT_LAST}")
set_dv_list(ws, "ListUnits", f"D{NT_FIRST}:D{NT_LAST}")
for cc in ("K", "L", "M", "N", "Q"):
    ws.column_dimensions[cc].hidden = True
ws.freeze_panes = f"F{NT_FIRST}"

# ----------------------------------------------------------------------------
# داشبورد (dashboard)
# ----------------------------------------------------------------------------
ws = SH["داشبورد"]
banner(ws, "داشبورد مدیریتی — وضعیت پرونده",
       "تمام مقادیر از سایر برگه‌ها محاسبه می‌شوند (بدون ورودی). برای تازه‌سازی: Ctrl+Alt+F9.", 18)
widths(ws, {gcl(i): 11.5 for i in range(1, 19)})
ws.column_dimensions["A"].width = 18
GT = f"'{G}'"
def gref(col):
    return f"{GT}!${col}${NT_FIRST}:${col}${NT_LAST}"
cards1 = [("وضعیت پرونده", "='تصویب'!$B$12", None),
          ("امتیاز امکان‌سنجی", "='امکان‌سنجی'!$F$17", "0.0"),
          ("پیشرفت واقعی کل", f"=IF({GT}!$J${NT_TOT}=\"\",0,{GT}!$J${NT_TOT})", "0.0%"),
          ("پیشرفت برنامه‌ای امروز", f"=IF({GT}!$Q${NT_TOT}=\"\",0,{GT}!$Q${NT_TOT})", "0.0%"),
          ("فعالیت در تأخیر (باز)", f"=COUNTIF({gref('R')},\"در تأخیر\")+COUNTIF({gref('R')},\"عقب‌افتاده\")", "0"),
          ("اقلام BOM قیمت‌گیری‌شده", "=IF(COUNTA('مهندسی_BOM'!$B$8:$B$31)=0,0,COUNTA('بازرگانی'!$B$8:$B$31)/COUNTA('مهندسی_BOM'!$B$8:$B$31))", "0%")]
cards2 = [("بهای تمام‌شده (م.ت.)", "=IF('مالی'!$C$8=\"\",0,'مالی'!$C$8)", "#,##0"),
          ("بودجه مصوب (م.ت.)", "=IF('مالی'!$B$19=\"\",0,'مالی'!$B$19)", "#,##0"),
          ("هزینه واقعی تا امروز (م.ت.)", "=IF('هزینه‌واقعی'!$H$30=\"\",0,'هزینه‌واقعی'!$H$30)", "#,##0"),
          ("CPI (اثر هزینه)", "=IF('هزینه‌واقعی'!$H$30=0,1,ROUND('مالی'!$B$19*IF('گانت'!$J$45=\"\",0,'گانت'!$J$45)/('هزینه‌واقعی'!$H$30),2))", "0.00"),
          ("SPI (اثر زمان)", "=IF(IF('گانت'!$Q$45=\"\",0,'گانت'!$Q$45)=0,1,ROUND(IF('گانت'!$J$45=\"\",0,'گانت'!$J$45)/'گانت'!$Q$45,2))", "0.00"),
          ("ریسک‌های باز", "=COUNTIF('ریسک'!$K$8:$K$27,\"باز\")", "0")]
cards3 = [("قیمت نهایی فروش (م.ت.)", "=IF('فروش'!$B$12=0,0,'فروش'!$B$12)", "#,##0"),
          ("حاشیه سود (٪)", "=IF('فروش'!$B$12=0,0,('فروش'!$B$12-'فروش'!$B$6)/'فروش'!$B$12)", "0.0%"),
          ("اقلام بحرانی تأمین", "=COUNTIF('برنامه‌ریزی'!$L$8:$L$31,\"بحرانی\")", "0"),
          ("تغییرات تأییدنشده", "=COUNTIF('تغییرات'!$H$8:$H$22,\"درخواست جدید\")", "0"),
          ("تاریخ امروز (شمسی)", '=IFERROR(INDEX(CalJ,MATCH(TODAY(),CalS,0)),"—")', None),
          ("مدیر پروژه", "='شناسنامه'!$B$12", None)]
for ci, cards in enumerate((cards1, cards2, cards3)):
    for k, (lab, f_, num) in enumerate(cards):
        c0 = 1 + (k % 6) * 3
        r_ = 5 + ci * 4
        merge(ws, r_, c0, r_, c0 + 2); merge(ws, r_ + 1, c0, r_ + 1, c0 + 2)
        put(ws, r_, c0, lab, f=font(8.5, False, WHITE), fl=pfill(BLUE), a=CENTER)
        put(ws, r_ + 1, c0, f_, f=font(15, True, NAVY), fl=pfill(WHITE), a=CENTER, num=num)
        fill_range(ws, r_ + 1, r_ + 1, c0, c0 + 2, WHITE)
        ws.row_dimensions[r_ + 1].height = 26
ws.conditional_formatting.add("A15:R15",
    FormulaRule(formula=["$A$15>0.05"], fill=pfill(RED)))

r_units = 19
put(ws, r_units - 1, 1, "پیشرفت واحدهای درگیر (از گانت)", f=font(11, True, NAVY), fl=pfill(LORANGE))
fill_range(ws, r_units - 1, r_units - 1, 1, 5, LORANGE)
header_row(ws, r_units, 1, ["واحد", "فعالیت", "پیشرفت واقعی ٪", "پیشرفت برنامه‌ای ٪", "در تأخیر"], height=26)
for i, u in enumerate(UNITS[1:-1]):
    rr = r_units + 1 + i
    put(ws, rr, 1, u, f=font(9.5, True))
    put(ws, rr, 2, f"=COUNTIF({gref('D')},$A{rr})", a=CENTER, f=font(9.5))
    put(ws, rr, 3, f"=IFERROR(SUMPRODUCT({gref('O')}*({gref('D')}=$A{rr})*{gref('J')})/SUMPRODUCT(({gref('D')}=$A{rr})*{gref('O')})/1,\"—\")", a=CENTER, f=font(9.5, True), num="0.0%")
    put(ws, rr, 4, f"=IFERROR(SUMPRODUCT({gref('O')}*({gref('D')}=$A{rr})*{gref('Q')})/SUMPRODUCT(({gref('D')}=$A{rr})*{gref('O')})/1,\"—\")", a=CENTER, f=font(9.5), num="0.0%")
    put(ws, rr, 5, f"=COUNTIFS({gref('D')},$A{rr},{gref('R')},\"در تأخیر\")+COUNTIFS({gref('D')},$A{rr},{gref('R')},\"عقب‌افتاده\")", a=CENTER, f=font(9.5))
u_last = r_units + len(UNITS[1:-1])

# S-curve helper block
r_s = u_last + 3
put(ws, r_s - 1, 1, "منحنی S — پیشرفت تجمعی (٪)", f=font(11, True, NAVY), fl=pfill(LORANGE))
fill_range(ws, r_s - 1, r_s - 1, 1, 14, LORANGE)
put(ws, r_s, 1, "ماه (پایان)", f=font(8.5, True), fl=pfill(GRAY), a=CENTER)
months_ = [(GRID_YEAR, m) for m in range(1, 13)]
for k, (yy, mm) in enumerate(months_):
    last_d = jdatetime.date(yy, mm + 1, 1).togregorian() - dt.timedelta(days=1) if mm < 12 else jdatetime.date(yy + 1, 1, 1).togregorian() - dt.timedelta(days=1)
    put(ws, r_s + 1, 2 + k, MONTHS_P[mm - 1], f=font(8, True), a=CENTER, fl=pfill(GRAY))
    put(ws, r_s + 2, 2 + k, ser(last_d), f=font(1, False, WHITE), a=CENTER, num=";;;")
    colL = gcl(2 + k)
    put(ws, r_s + 3, 2 + k,
        f"=IFERROR(100*(SUMPRODUCT({gref('O')},({colL}${r_s + 2}>={gref('K')})*({colL}${r_s + 2}<{gref('L')})*(({colL}${r_s + 2}-{gref('K')}+1)/({gref('L')}-{gref('K')}+1)))+SUMPRODUCT({gref('O')},({colL}${r_s + 2}>={gref('L')})*({gref('L')}>0)))/SUM({gref('O')}),\"\")",
        f=font(8), a=CENTER, num="0")
    put(ws, r_s + 4, 2 + k,
        f"=IFERROR(100*SUMPRODUCT({gref('O')},({gref('M')}>0)*({gref('M')}<={colL}${r_s + 2}),{gref('J')})/SUM({gref('O')}),\"\")",
        f=font(8), a=CENTER, num="0")
put(ws, r_s + 3, 1, "برنامه‌ای", f=font(8.5, True), fl=pfill(GRAY), a=CENTER)
put(ws, r_s + 4, 1, "واقعی", f=font(8.5, True), fl=pfill(GRAY), a=CENTER)

# charts
def mkbar(title, data_ref, cat_ref, color="2F5597", percent=False):
    ch = BarChart(); ch.type = "col"; ch.style = 10; ch.title = title
    ch.add_data(data_ref, titles_from_data=False)
    ch.set_categories(cat_ref)
    ch.legend = None; ch.gapWidth = 60
    s = ch.series[0]
    from openpyxl.chart.marker import DataPoint
    s.graphicalProperties.solidFill = color
    if percent:
        ch.y_axis.numFmt = "0%"
    ch.y_axis.majorGridlines = None
    return ch

ch = mkbar("پیشرفت واقعی واحدها (٪)", Reference(SH["داشبورد"], min_col=3, min_row=r_units + 1, max_row=u_last),
           Reference(SH["داشبورد"], min_col=1, min_row=r_units + 1, max_row=u_last))
ch.y_axis.delete = False; ch.x_axis.delete = False
ch.y_axis.numFmt = "0%"; ch.height = 8.6; ch.width = 17.5
ch.dLbls = None
ws.add_chart(ch, f"G{r_units}")

ch2 = LineChart(); ch2.title = "منحنی S — برنامه‌ای در برابر واقعی (٪)"; ch2.style = 12
ch2.add_data(Reference(SH["داشبورد"], min_col=1, min_row=r_s + 3, max_col=13, max_row=r_s + 4), titles_from_data=True, from_rows=True)
ch2.set_categories(Reference(SH["داشبورد"], min_col=2, min_row=r_s + 1, max_col=13, max_row=r_s + 1))
ch2.height = 8.2; ch2.width = 17.5
ch2.y_axis.numFmt = "0"
ch2.y_axis.delete = False; ch2.x_axis.delete = False
for si, col_ in ((0, "2F5597"), (1, LGREEN)):
    ch2.series[si].graphicalProperties.line.solidFill = col_
    ch2.series[si].graphicalProperties.line.width = 22000
    ch2.series[si].smooth = False
ws.add_chart(ch2, f"A{r_s + 6}")

ch3 = BarChart(); ch3.type = "col"; ch3.title = "بودجه در برابر هزینه واقعی (م.ت.)"; ch3.style = 10
ch3.add_data(Reference(SH["مالی"], min_col=2, min_row=12, max_col=3, max_row=18), titles_from_data=True)
ch3.set_categories(Reference(SH["مالی"], min_col=1, min_row=13, max_row=18))
ch3.height = 8.2; ch3.width = 13.5
ch3.y_axis.delete = False; ch3.x_axis.delete = False
ch3.series[0].graphicalProperties.solidFill = "2F5597"
ch3.series[1].graphicalProperties.solidFill = "C00000"
ws.add_chart(ch3, f"I{r_s + 6}")

ch4 = PieChart(); ch4.title = "وضعیت ریسک‌ها"
ch4.add_data(Reference(SH["داشبورد"], min_col=3, min_row=44, max_row=45), titles_from_data=False)
ch4.set_categories(Reference(SH["داشبورد"], min_col=1, min_row=44, max_row=45))
ch4.height = 8; ch4.width = 8.5
ws.add_chart(ch4, f"A48")
put(ws, 43, 1, "ریسک — جمع‌بندی", f=font(11, True, NAVY), fl=pfill(LORANGE)); fill_range(ws, 43, 43, 1, 5, LORANGE)
put(ws, 44, 1, "باز", f=font(9), fl=pfill(GRAY))
put(ws, 44, 3, "=COUNTIF('ریسک'!$K$8:$K$27,\"باز\")", a=CENTER, f=font(9, True))
put(ws, 45, 1, "بسته/پایش", f=font(9), fl=pfill(GRAY))
put(ws, 45, 3, "=COUNTIF('ریسک'!$K$8:$K$27,\"بسته\")+COUNTIF('ریسک'!$K$8:$K$27,\"پایش\")", a=CENTER, f=font(9, True))

r_ck = 50
put(ws, r_ck - 1, 1, "چک‌لیست گردش فرآیند (خودکار بر اساس تکمیل برگه‌ها)", f=font(11, True, NAVY), fl=pfill(LORANGE))
fill_range(ws, r_ck - 1, r_ck - 1, 1, 5, LORANGE)
header_row(ws, r_ck, 1, ["مرحله", "واحد متولی", "وضعیت ثبت"], height=22)
checklist = [
    ("ثبت سفارش/شناسنامه", "واحد ثبت سفارش", f"=IF(COUNTA('شناسنامه'!$B$5)>0,\"✓ ثبت شد\",\"—\")"),
    ("امکان‌سنجی", "مدیر پروژه", "=IF('امکان‌سنجی'!$F$17>0,\"✓ انجام شد\",\"—\")"),
    ("گانت اولیه", "مدیر پروژه", f"=IF(COUNTA({GT}!$C${NT_FIRST}:$C${NT_LAST})>0,\"✓ تدوین شد\",\"—\")"),
    ("مدارک مهندسی", "مهندسی", "=IF(COUNTA('مهندسی_مدارک'!$B$8:$B$31)>0,\"✓ دریافت شد\",\"—\")"),
    ("BOM و نفرساعت", "مهندسی", "=IF(COUNTA('مهندسی_BOM'!$B$8:$B$31)>0,\"✓ تهیه شد\",\"—\")"),
    ("قیمت‌گذاری تأمین", "بازرگانی", "=IF(COUNTA('بازرگانی'!$B$8:$B$31)>0,\"✓ قیمت‌گذاری شد\",\"—\")"),
    ("نیازسنجی تولید و انبار", "برنامه‌ریزی", "=IF(COUNTA('برنامه‌ریزی'!$A$8:$A$31)>0,\"✓ انجام شد\",\"—\")"),
    ("بهای تمام‌شده و بودجه", "مالی", "=IF('مالی'!$C$8>0,\"✓ تدوین شد\",\"—\")"),
    ("قیمت فروش و پرداخت", "فروش", "=IF('فروش'!$B$12>0,\"✓ تعیین شد\",\"—\")"),
    ("جمع‌بندی مدیر پروژه", "مدیر پروژه", "=IF(COUNTA('جمع‌بندی'!$A$14)>0,\"✓ ارسال شد\",\"—\")"),
    ("تصمیم مدیریت ارشد", "مدیریت ارشد", "=IF('تصویب'!$B$12<>\"\",\"✓ ثبت شد\",\"در انتظار\")"),
]
for i, (a, b, f_) in enumerate(checklist):
    rr = r_ck + 1 + i
    put(ws, rr, 1, a, f=font(9.5))
    put(ws, rr, 2, b, f=font(9), a=CENTER)
    put(ws, rr, 3, f_, f=font(9.5, True), a=CENTER)
ws.conditional_formatting.add(f"C{r_ck + 1}:C{r_ck + len(checklist)}",
    FormulaRule(formula=[f'ISNUMBER(SEARCH("✓",$C{r_ck + 1}))'], fill=pfill(GREEN), font=font(9.5, True, DGREEN)))
ws.freeze_panes = "A5"
# ----------------------------------------------------------------------------
# مهندسی_BOM (BOM + man-hours)
# ----------------------------------------------------------------------------
ws = SH["مهندسی_BOM"]
banner(ws, "مهندسی — لیست مواد (BOM) و نفرساعت",
       "BOM مبنای کار بازرگانی، برنامه‌ریزی و مالی است. ستون K وضعیت قیمت‌گذاری را از برگه بازرگانی به‌صورت خودکار نشان می‌دهد. مقادیر بر مبنای یک محصول نهایی.", 12)
widths(ws, {"A": 5, "B": 13, "C": 30, "D": 30, "E": 8, "F": 10, "G": 9, "H": 10, "I": 13, "J": 11, "K": 15, "L": 12})
section(ws, 6, 1, 12, "الف) BOM — بر مبنای یک دستگاه محصول نهایی")
header_row(ws, 7, 1, ["ردیف", "کد فنی", "شرح قلم", "مشخصات / استاندارد", "واحد", "مقدار پایه",
                      "ضایعات ٪", "مقدار نهایی", "کلاس", "منبع", "قیمت‌گذاری بازرگانی", "مقدار سفارش"])
BOM_FIRST, BOM_LAST = 8, 31
BOM_ITEMS = [
    ("RM-RES-ISO", "رزین ایزوفتالیک", "ویسکوزیته پایین، استایرن ۴۵-۵۰٪، IRAM 7583", "کیلوگرم", 5200, 5, "مواد اولیه", "داخلی"),
    ("RM-GF-MAT", "فایبرگلاس MAT ۴۵۰", "E-Glass 450g/m² — ASTM D5767", "کیلوگرم", 2100, 8, "مواد اولیه", "داخلی"),
    ("RM-GF-ROV", "رووینگ ۲۴۰۰tex", "E-Glass 2400tex — for filament winding", "کیلوگرم", 1650, 6, "مواد اولیه", "داخلی"),
    ("RM-CH-PER", "کاتالیزور MEKP", "متیل اتیل کتون پر اکساید ۹٪", "کیلوگرم", 95, 3, "شیمیایی", "داخلی"),
    ("RM-CH-ACC", "شتاب‌دهنده کبالت", "اکتوات کبالت ۶٪", "لیتر", 28, 2, "شیمیایی", "داخلی"),
    ("RM-GEL", "ژل‌کوت اورتالیک", "مقاوم شیمیایی — رنگ خاکستری RAL7035", "کیلوگرم", 340, 6, "مواد اولیه", "وارداتی"),
    ("EQ-PUMP", "پمپ تزریق رزیون", "دوجزئی، دبی ۱۲L/min، فشار ۲۵bar", "عدد", 1, 0, "تجهیزات", "وارداتی"),
    ("EQ-MIX", "میکسر صنعتی ۷۵۰ لیتری", "استیل ۳۰۴، موتور ۵.۵kW", "عدد", 1, 0, "تجهیزات", "داخلی"),
    ("CW-FLG-16", "فلنج کربن استیل DN150", "ASME B16.5 Class150 RF", "عدد", 6, 0, "متفرقه", "داخلی"),
    ("CW-FLG-24", "فلنج کربن استیل DN200", "ASME B16.5 Class150 RF", "عدد", 2, 0, "متفرقه", "داخلی"),
    ("CW-MAN", "مردآب (Manway) ۵۰۰mm", "فلزی با درب و اورینگ EPDM", "عدد", 1, 0, "متفرقه", "داخلی"),
    ("CW-NIP", "نازل خروج/ورود و شیرآف", "DN50/DN80، فلنجی", "عدد", 14, 0, "متفرقه", "داخلی"),
    ("CW-SEN", "سنسور سطح و دما", "راداری + PT100 با ترانسمیتر", "عدد", 4, 0, "ابزار دقیق", "وارداتی"),
    ("AW-PNT", "پرایمر و پوشش نهایی UV", "ژل‌کوت پوششی مقاوم UV", "کیلوگرم", 120, 5, "شیمیایی", "داخلی"),
    ("PK-PKG", "بسته‌بندی و مهاربندی حمل", "وارق آستر + تسمه‌بندی صنعتی", "بسته", 1, 0, "بسته‌بندی", "داخلی"),
    ("SV-CAL", "کالیبراسیون و تست هیدرواستاتیک", "خدمات ثالث — ISO 17025", "دفعه", 1, 0, "خدمات", "داخلی"),
]
for i in range(BOM_LAST - BOM_FIRST + 1):
    rr = BOM_FIRST + i
    if i < len(BOM_ITEMS):
        code, desc, spec, unit, qty, waste, cls, src = BOM_ITEMS[i]
        put(ws, rr, 1, i + 1, a=CENTER, f=font(8.5))
        put(ws, rr, 2, code, f=font(9, True, NAVY), a=CENTER, fl=pfill(YELLOW), unlock=True)
        put(ws, rr, 3, desc, f=font(9), fl=pfill(YELLOW), unlock=True)
        put(ws, rr, 4, spec, f=font(8.5), fl=pfill(YELLOW), unlock=True)
        put(ws, rr, 5, unit, a=CENTER, f=font(9), fl=pfill(YELLOW), unlock=True)
        put(ws, rr, 6, qty, a=CENTER, f=font(9), num="#,##0", fl=pfill(YELLOW), unlock=True)
        put(ws, rr, 7, waste, a=CENTER, f=font(9), num="0", fl=pfill(YELLOW), unlock=True)
        put(ws, rr, 8, f"=IF($F{rr}=\"\",\"\",ROUND($F{rr}*(1+$G{rr}/100),1))", a=CENTER, f=font(9, True), num="#,##0.0", fl=pfill(GRAY))
        put(ws, rr, 9, cls, a=CENTER, f=font(9), fl=pfill(YELLOW), unlock=True)
        put(ws, rr, 10, src, a=CENTER, f=font(9), fl=pfill(YELLOW), unlock=True)
        put(ws, rr, 11, f"=IF($B{rr}=\"\",\"\",IF(COUNTIF('بازرگانی'!$B$8:$B$31,$B{rr})>0,\"✓ قیمت‌گذاری شد\",\"در انتظار قیمت\"))",
            a=CENTER, f=font(8.5), fl=pfill("DDEBF7"))
        put(ws, rr, 12, f"=IF($H{rr}=\"\",\"\",ROUND($H{rr}*$B$5,1))", a=CENTER, f=font(9), num="#,##0.0", fl=pfill(GRAY))
    else:
        for cc in (1, 2, 3, 4, 5, 6, 7, 9, 10):
            put(ws, rr, cc, None, fl=pfill(YELLOW), unlock=True,
                num="#,##0" if cc in (6, 7) else None, f=font(9))
        put(ws, rr, 8, f"=IF($F{rr}=\"\",\"\",ROUND($F{rr}*(1+$G{rr}/100),1))", a=CENTER, f=font(9, True), num="#,##0.0")
        put(ws, rr, 11, f"=IF($B{rr}=\"\",\"\",IF(COUNTIF('بازرگانی'!$B$8:$B$31,$B{rr})>0,\"✓ قیمت‌گذاری شد\",\"در انتظار قیمت\"))", a=CENTER, f=font(8.5))
        put(ws, rr, 12, f"=IF($H{rr}=\"\",\"\",ROUND($H{rr}*$B$5,1))", a=CENTER, f=font(9), num="#,##0.0")
put(ws, 5, 1, "تعداد محصول سفارشی", f=font(9.5, True), fl=pfill(GRAY))
put(ws, 5, 2, 1, a=CENTER, f=font(10, True), fl=pfill(YELLOW), unlock=True, num="0")
put(ws, 5, 3, "← ضریب کل BOM (ستون L)؛ برنامه‌ریزی تولید این ستون را می‌خواند", f=font(8.5, False, DARKTXT), bd=False)
tr = BOM_LAST + 1
fill_range(ws, tr, tr, 1, 12, LBLUE)
put(ws, tr, 3, "جمع اقلام BOM", f=font(10, True), fl=pfill(LBLUE))
put(ws, tr, 6, f"=COUNTA($B${BOM_FIRST}:$B${BOM_LAST})", a=CENTER, f=font(10, True), num="0", fl=pfill(LBLUE))
put(ws, tr, 11, f"=TEXT(COUNTIF($K${BOM_FIRST}:$K${BOM_LAST},\"*قیمت‌گذاری*\"),\"0\")&\" از \"&TEXT(COUNTA($B${BOM_FIRST}:$B${BOM_LAST}),\"0\")",
    a=CENTER, f=font(9, True, NAVY), fl=pfill(LBLUE))
ws.conditional_formatting.add(f"K{BOM_FIRST}:K{BOM_LAST}",
    FormulaRule(formula=[f'ISNUMBER(SEARCH("✓",$K{BOM_FIRST}))'], fill=pfill(GREEN), font=font(8.5, True, DGREEN)))
ws.conditional_formatting.add(f"K{BOM_FIRST}:K{BOM_LAST}",
    FormulaRule(formula=[f'ISNUMBER(SEARCH("انتظار",$K{BOM_FIRST}))'], fill=pfill(AMBER), font=font(8.5, True, DAMBER)))

# man-hours table
mh0 = tr + 3
section(ws, mh0 - 1, 1, 12, "ب) نفرساعت و زمان‌سنجی (مبنای حق‌الزحمه مستقیم نیروی کار)")
header_row(ws, mh0, 1, ["فعالیت", "تعداد نفر", "ساعت/روز", "روز کاری", "نفرساعت",
                       "نرخ ساعتی (تومان)", "هزینه کل (م.ت.)", "", "", "", "", ""], height=26)
MHH = [
    ("طراحی و نقشه‌کشی", 2, 8, 12), ("محاسبات سازه و لایه‌چین", 1, 6, 8),
    ("ساخت قالب و الگو", 3, 8, 10), ("لایه‌نشانی و تولید بدنه", 6, 8, 18),
    ("مونتاژ نازل و اتصالات", 3, 8, 6), ("کنترل کیفیت و تست", 2, 8, 5),
    ("بسته‌بندی، حمل و نصب", 4, 8, 4),
]
for i, (nm, ppl, hr, days) in enumerate(MHH):
    rr = mh0 + 1 + i
    put(ws, rr, 1, nm, f=font(9.5))
    put(ws, rr, 2, ppl, a=CENTER, f=font(9.5), fl=pfill(YELLOW), unlock=True, num="0")
    put(ws, rr, 3, hr, a=CENTER, f=font(9.5), fl=pfill(YELLOW), unlock=True, num="0")
    put(ws, rr, 4, days, a=CENTER, f=font(9.5), fl=pfill(YELLOW), unlock=True, num="0")
    put(ws, rr, 5, f"=$B{rr}*$C{rr}*$D{rr}", a=CENTER, f=font(9.5, True), num="#,##0", fl=pfill(GRAY))
    put(ws, rr, 6, 250000 + i * 15000, a=CENTER, f=font(9.5), num="#,##0", fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 7, f"=ROUND($E{rr}*$F{rr}/1000000,1)", a=CENTER, f=font(9.5, True), num="#,##0.0", fl=pfill(GRAY))
mh_rows_last = mh0 + len(MHH)
tr2 = mh_rows_last + 1
fill_range(ws, tr2, tr2, 1, 7, LBLUE)
put(ws, tr2, 1, "جمع نفرساعت / هزینه نیروی کار", f=font(10, True), fl=pfill(LBLUE))
put(ws, tr2, 5, f"=SUM(E{mh0 + 1}:E{mh_rows_last})", a=CENTER, f=font(10, True), num="#,##0", fl=pfill(LBLUE))
put(ws, tr2, 7, f"=SUM(G{mh0 + 1}:G{mh_rows_last})", a=CENTER, f=font(10, True, NAVY), num="#,##0.0", fl=pfill(LBLUE))
MH_TOTAL = f"'مهندسی_BOM'!$G${tr2}"
ws.auto_filter.ref = f"A7:L{BOM_LAST}"

# ----------------------------------------------------------------------------
# مهندسی_مدارک (document register)
# ----------------------------------------------------------------------------
ws = SH["مهندسی_مدارک"]
banner(ws, "مهندسی — رجیستر مدارک پروژه",
       "نقشه‌ها، Test Plan، مشخصات فنی و گواهی‌ها. هر مدرک ورودی از کارفرما/مشاور در اینجا ثبت و شماره‌گذاری می‌شود.", 10)
widths(ws, {"A": 5, "B": 15, "C": 38, "D": 14, "E": 8, "F": 12, "G": 18, "H": 13, "I": 30, "J": 10})
header_row(ws, 6, 1, ["ردیف", "کد مدرک", "عنوان مدرک", "نوع", "Revision", "تاریخ (شمسی)",
                      "منبع/تحویل‌دهنده", "وضعیت", "مسیر فایل / شبکه", "پیوست"])
DOCS = [
    ("BC-ENG-DRW-001", "نقشه جامع مخزن 250m³", "نقشه", "C", "1405/05/02", "کارفرما — مهندس مشاور", "دریافت‌شده", "\\\\BORNASRV\\Projects\\BC-P-1405-01\\Engineering\\Drawings"),
    ("BC-ENG-DRW-002", "نقشه لایه‌چینی (Lay-up Schedule)", "نقشه", "B", "1405/05/06", "مهندسی برنا", "دریافت‌شده", "\\\\BORNASRV\\Projects\\BC-P-1405-01\\Engineering\\Layup"),
    ("BC-ENG-TPL-001", "طرح قالب اسپیل و مونتاژ", "نقشه", "A", "1405/06/08", "مهندسی برنا", "به‌روزرسانی لازم", "\\\\BORNASRV\\Projects\\BC-P-1405-01\\Engineering\\Mold"),
    ("BC-ENG-TP-001", "Test Plan — آزمون هیدرواستاتیک و نشت", "Test Plan", "A", "1405/05/21", "مهندسی + QC", "دریافت‌شده", "\\\\BORNASRV\\Projects\\BC-P-1405-01\\Quality\\TP"),
    ("BC-ENG-SPC-001", "مشخصات فنی رزین و الیاف", "مشخصات", "D", "1405/05/25", "بازرگانی/تأمین‌کننده", "دریافت‌شده", "\\\\BORNASRV\\Projects\\BC-P-1405-01\\Procurement\\Specs"),
    ("BC-ENG-CAL-001", "محاسبات سازه ASME RTP-1", "محاسبات", "B", "1405/06/12", "مهندس مشاور", "منتظر دریافت", "\\\\BORNASRV\\Projects\\BC-P-1405-01\\Engineering\\Calcs"),
    ("BC-ENG-HSE-001", "ماتریس ایمنی کار در ارتفاع و مواد شیمیایی", "HSE", "A", "1405/06/18", "HSE", "دریافت‌شده", "\\\\BORNASRV\\Projects\\BC-P-1405-01\\HSE"),
    ("BC-ENG-CRT-001", "گواهی‌های آزمون مواد (MTC)", "گواهی", "—", "", "تأمین‌کنندگان", "منتظر دریافت", ""),
]
DOC_FIRST, DOC_LAST = 8, 27
for i in range(DOC_LAST - DOC_FIRST + 1):
    rr = DOC_FIRST + i
    if i < len(DOCS):
        code, ttl, typ, rev, dtx, src, st, path = DOCS[i]
        put(ws, rr, 1, i + 1, a=CENTER, f=font(8.5))
        put(ws, rr, 2, code, f=font(9, True, NAVY), a=CENTER)
        put(ws, rr, 3, ttl, f=font(9))
        put(ws, rr, 4, typ, a=CENTER, f=font(9))
        put(ws, rr, 5, rev, a=CENTER, f=font(9))
        put(ws, rr, 6, dtx, a=CENTER, f=font(9), num="@")
        put(ws, rr, 7, src, f=font(8.5))
        put(ws, rr, 8, st, a=CENTER, f=font(9, True))
        put(ws, rr, 9, path, f=font(8, False, DARKTXT))
        put(ws, rr, 10, "—", a=CENTER, f=font(8.5))
    else:
        for cc in range(1, 11):
            put(ws, rr, cc, None, fl=pfill(YELLOW), unlock=True, num="@" if cc == 6 else None, f=font(9))
set_dv_list(ws, "ListDoc", f"H{DOC_FIRST}:H{DOC_LAST}")
ws.conditional_formatting.add(f"H{DOC_FIRST}:H{DOC_LAST}",
    FormulaRule(formula=[f'$H{DOC_FIRST}="دریافت‌شده"'], fill=pfill(GREEN), font=font(9, True, DGREEN)))
ws.conditional_formatting.add(f"H{DOC_FIRST}:H{DOC_LAST}",
    FormulaRule(formula=[f'OR($H{DOC_FIRST}="منتظر دریافت",$H{DOC_FIRST}="مردود")'], fill=pfill(RED), font=font(9, True, DRED)))
ws.conditional_formatting.add(f"H{DOC_FIRST}:H{DOC_LAST}",
    FormulaRule(formula=[f'$H{DOC_FIRST}="به‌روزرسانی لازم"'], fill=pfill(AMBER), font=font(9, True, DAMBER)))
ws.auto_filter.ref = f"A6:J{DOC_LAST}"

# ----------------------------------------------------------------------------
# بازرگانی — قیمت‌گذاری اقلام BOM
# ----------------------------------------------------------------------------
ws = SH["بازرگانی"]
banner(ws, "بازرگانی داخلی و خارجی — قیمت‌گذاری اقلام BOM",
       "شرح/واحد/مقدار از BOM مهندسی با کد فنی لینک می‌شود (فیلدهای طوسی را ویرایش نکنید). قیمت واحد به پول قلم؛ نرخ ارز از برگه تنظیمات. جمع نهایی به میلیون تومان به واحد مالی ارسال می‌شود.", 20)
widths(ws, {"A": 5, "B": 13, "C": 26, "D": 8, "E": 11, "F": 12, "G": 20, "H": 8, "I": 13, "J": 11,
            "K": 9, "L": 12, "M": 14, "N": 15, "O": 11, "P": 12, "Q": 12, "R": 9, "S": 12, "T": 13})
header_row(ws, 7, 1, ["ردیف", "کد فنی قلم", "شرح (از BOM)", "واحد", "مقدار نهایی", "نوع (داخلی/خارجی)",
                      "تأمین‌کننده / کشور", "ارز", "قیمت واحد", "نرخ ارز (تومان)", "عوارض+گمرک ٪",
                      "حمل و بیمه (تومان)", "قیمت تمام‌شده واحد (تومان)", "جمع (تومان)",
                      "تاریخ استعلام", "موعد تحویل تأمین‌کننده", "سریال موعد", "سریال استعلام", "Lead (روز)", "وضعیت"])
TRD_FIRST, TRD_LAST = 8, 31
TRD_DEMO = [
    ("RM-RES-ISO", "داخلی", "شرکت پترو رزین اصفهان", "تومان", 185000, 0, 0, "1405-06-05", "1405-06-28"),
    ("RM-GF-MAT", "داخلی", "ایران کامپوزیت رشت", "تومان", 152000, 0, 0, "1405-06-06", "1405-07-05"),
    ("RM-GF-ROV", "داخلی", "ایران کامپوزیت رشت", "تومان", 138000, 0, 25000000, "1405-06-06", "1405-07-05"),
    ("RM-CH-PER", "داخلی", "شیمی‌پارس", "تومان", 2400000, 0, 0, "1405-06-08", "1405-06-20"),
    ("RM-CH-ACC", "داخلی", "شیمی‌پارس", "تومان", 3100000, 0, 0, "1405-06-08", "1405-06-20"),
    ("RM-GEL", "خارجی", "Polypaint — ترکیه", "یورو", 4.6, 21, 38000000, "1405-06-09", "1405-07-28"),
    ("EQ-PUMP", "خارجی", "Graco — آمریکا/واسط هند", "دلار", 21500, 23, 145000000, "1405-06-10", "1405-08-02"),
    ("EQ-MIX", "داخلی", "سازنده ماشین‌سازی تبریز", "تومان", 1280000000, 0, 42000000, "1405-06-12", "1405-07-22"),
    ("CW-FLG-16", "داخلی", "آهن‌پایه تهران", "تومان", 9600000, 0, 0, "1405-06-13", "1405-07-01"),
    ("CW-FLG-24", "داخلی", "آهن‌پایه تهران", "تومان", 14200000, 0, 0, "1405-06-13", "1405-07-01"),
    ("CW-MAN", "داخلی", "فولاد ساز اراک", "تومان", 268000000, 0, 12000000, "1405-06-15", "1405-07-18"),
    ("CW-NIP", "داخلی", "لوله‌سازان اصفهان", "تومان", 21500000, 0, 0, "1405-06-15", "1405-07-10"),
    ("CW-SEN", "خارجی", "VEGA — آلمان (کارگزار)", "یورو", 1180, 17, 46000000, "1405-06-17", "1405-08-15"),
    ("AW-PNT", "داخلی", "رنگ‌سازی یزد", "تومان", 1650000, 0, 0, "1405-06-18", "1405-07-02"),
    ("PK-PKG", "داخلی", "بسته‌بندی برنا (خودکفا)", "تومان", 285000000, 0, 0, "1405-06-20", "1405-08-05"),
    ("SV-CAL", "داخلی", "آزمایشگاه ثالث — ISI", "تومان", 640000000, 0, 0, "1405-06-22", "1405-09-20"),
]
for i in range(TRD_LAST - TRD_FIRST + 1):
    rr = TRD_FIRST + i
    demo = TRD_DEMO[i] if i < len(TRD_DEMO) else None
    put(ws, rr, 1, i + 1, a=CENTER, f=font(8.5))
    inp = pfill(YELLOW); calcf = pfill(GRAY)
    put(ws, rr, 2, demo[0] if demo else None, a=CENTER, f=font(8.5, True, NAVY), fl=inp, unlock=True)
    put(ws, rr, 3, f"=IF($B{rr}=\"\",\"\",IFERROR(INDEX('مهندسی_BOM'!$C$8:$C$31,MATCH($B{rr},'مهندسی_BOM'!$B$8:$B$31,0)),\"⚠ کد نامعتبر\"))", f=font(9), fl=calcf)
    put(ws, rr, 4, f"=IF($B{rr}=\"\",\"\",IFERROR(INDEX('مهندسی_BOM'!$E$8:$E$31,MATCH($B{rr},'مهندسی_BOM'!$B$8:$B$31,0)),\"\"))", a=CENTER, f=font(9), fl=calcf)
    put(ws, rr, 5, f"=IF($B{rr}=\"\",\"\",IFERROR(INDEX('مهندسی_BOM'!$H$8:$H$31,MATCH($B{rr},'مهندسی_BOM'!$B$8:$B$31,0)),\"\"))", a=CENTER, f=font(9, True), num="#,##0.0", fl=calcf)
    put(ws, rr, 6, demo[1] if demo else None, a=CENTER, f=font(9), fl=inp, unlock=True)
    put(ws, rr, 7, demo[2] if demo else None, f=font(9), fl=inp, unlock=True)
    put(ws, rr, 8, demo[3] if demo else None, a=CENTER, f=font(9), fl=inp, unlock=True)
    put(ws, rr, 9, demo[4] if demo else None, a=CENTER, f=font(9, True), num="#,##0.00", fl=inp, unlock=True)
    put(ws, rr, 10, f"=IF($H{rr}=\"\",\"\",IF($H{rr}=\"تومان\",1,IF($H{rr}=\"دلار\",FX_USD,FX_EUR)))", a=CENTER, f=font(9), num="#,##0", fl=calcf)
    put(ws, rr, 11, (demo[5] if demo else 0), a=CENTER, f=font(9), num="0", fl=inp, unlock=True)
    put(ws, rr, 12, (demo[6] if demo else None), a=CENTER, f=font(9), num="#,##0", fl=inp, unlock=True)
    put(ws, rr, 13, f"=IF(OR($I{rr}=\"\",$J{rr}=\"\"),\"\",ROUND($I{rr}*$J{rr}*(1+$K{rr}/100),0))", a=CENTER, f=font(9, True), num="#,##0", fl=calcf)
    put(ws, rr, 14, f"=IF($M{rr}=\"\",\"\",$M{rr}*$E{rr}+IF($L{rr}=\"\",0,$L{rr}))", a=CENTER, f=font(9, True), num="#,##0", fl=calcf)
    put(ws, rr, 15, demo[7] if demo else None, a=CENTER, f=font(9), num="@", fl=inp, unlock=True)
    put(ws, rr, 16, demo[8] if demo else None, a=CENTER, f=font(9), num="@", fl=inp, unlock=True)
    put(ws, rr, 17, j2s_formula(f"$P{rr}") + "", a=CENTER, f=font(6), num=";;;")
    put(ws, rr, 18, j2s_formula(f"$O{rr}"), a=CENTER, f=font(6), num=";;;")
    put(ws, rr, 19, f"=IF(AND($Q{rr}>0,$R{rr}>0),$Q{rr}-$R{rr},\"\")", a=CENTER, f=font(8.5), num="0", fl=calcf)
    put(ws, rr, 20, f"=IF($N{rr}=\"\",\"در انتظار قیمت\",IF(AND($Q{rr}>0,$Q{rr}<TODAY()),\"⚠ موعد گذشته\",IF(AND($Q{rr}>0,$Q{rr}-TODAY()<=14),\"نزدیک موعد\",\"در جریان\")))", a=CENTER, f=font(8.5, True))
# wait: column 17=Q serial promised, 18=R serial inquiry — Q>TODAY compare uses Q — ok names:
for cc, col in ((17, "Q"), (18, "R")):
    pass
trr = TRD_LAST + 1
fill_range(ws, trr, trr, 1, 20, LBLUE)
put(ws, trr, 3, "جمع کل اقلام قیمت‌گذاری‌شده (تومان)", f=font(10, True), fl=pfill(LBLUE))
put(ws, trr, 14, f"=SUM($N${TRD_FIRST}:$N${TRD_LAST})", a=CENTER, f=font(10, True), num="#,##0", fl=pfill(LBLUE))
put(ws, trr, 15, "میلیون تومان →", f=font(9, False, DARKTXT), bd=False)
put(ws, trr, 16, f"=ROUND($N${trr}/1000000,1)", a=CENTER, f=font(12, True, NAVY), num="#,##0.0", fl=pfill(LBLUE))
TRD_TOTAL_MT = f"'بازرگانی'!$P${trr}"
# FX block
rfx = trr + 2
put(ws, rfx, 2, "نرخ ارز مورد استفاده (از تنظیمات — ویرایش در تنظیمات):", f=font(9, True), fl=pfill(GRAY))
put(ws, rfx, 4, "دلار:", f=font(9), a=CENTER, bd=False)
put(ws, rfx, 5, "=FX_USD", f=font(9, True), num="#,##0", fl=pfill(GRAY), a=CENTER)
put(ws, rfx, 6, "یورو:", f=font(9), a=CENTER, bd=False)
put(ws, rfx, 7, "=FX_EUR", f=font(9, True), num="#,##0", fl=pfill(GRAY), a=CENTER)
put(ws, rfx + 1, 2, "مستندسازی: حداقل دو پیش‌فاکتور معتبر برای هر قلام (بالای ۵۰۰ م.ت.) الزامی است — فایل‌ها در مسیر شبکه پروژه.", f=font(8.5, False, DARKTXT), bd=False)
merge(ws, rfx + 1, 2, rfx + 1, 12)
set_dv_list(ws, "ListCur", f"H{TRD_FIRST}:H{TRD_LAST}")
for st, hex_ in (("در جریان", "DDEBF7"), ("نزدیک موعد", AMBER), ("⚠ موعد گذشته", RED), ("در انتظار قیمت", GRAY)):
    ws.conditional_formatting.add(f"T{TRD_FIRST}:T{TRD_LAST}",
        FormulaRule(formula=[f'$T{TRD_FIRST}="{st}"'], fill=pfill(hex_)))
ws.freeze_panes = f"C{TRD_FIRST}"
ws.auto_filter.ref = f"A7:T{TRD_LAST}"

# ----------------------------------------------------------------------------
# برنامه‌ریزی و انبار (planning / inventory)
# ----------------------------------------------------------------------------
ws = SH["برنامه‌ریزی"]
banner(ws, "برنامه‌ریزی تولید و انبارها — نیازسنجی و زمان‌بندی تأمین",
       "نیاز کل = مقدار نهایی BOM × تعداد سفارش. نیاز خالص پس از کسر موجودی و در‌راه محاسبه می‌شود؛ تاریخ درخواست از «نیاز − Lead تأمین‌کننده» به‌صورت خودکار (شمسی) استخراج می‌گردد.", 13)
widths(ws, {"A": 13, "B": 26, "C": 11, "D": 10, "E": 9, "F": 11, "G": 9, "H": 12, "I": 10, "J": 8,
            "K": 13, "L": 13, "M": 22})
header_row(ws, 6, 1, ["کد فنی", "شرح (از BOM)", "نیاز کل", "موجودی انبار", "در راه", "نیاز خالص",
                      "ذخیره ایمنی", "تاریخ نیاز (شمسی)", "سریال نیاز", "Lead (روز)", "تاریخ درخواست خرید",
                      "وضعیت تأمین", "یادداشت / سفارش"])
put(ws, 4, 1, "تعداد سفارش تولید", f=font(9.5, True), fl=pfill(GRAY))
put(ws, 4, 2, "='مهندسی_BOM'!$B$5", a=CENTER, f=font(10, True), num="0", fl=pfill("DDEBF7"))
PLAN_FIRST, PLAN_LAST = 8, 31
PLAN_DEMO = {
    "RM-RES-ISO": (3400, 1900, 300, "1405-07-20", "PO-1405-071 تأیید شد"),
    "RM-GF-MAT":  (1700, 650, 250, "1405-07-25", "رسید تا ۷/۵"),
    "RM-GF-ROV":  (1300, 900, 200, "1405-07-25", ""),
    "RM-CH-PER":  (78, 40, 12, "1405-07-12", "مصرف فاز تولید"),
    "EQ-PUMP":    (1, 0, 0, "1405-07-30", "وابسته به تخصیص ارز"),
    "CW-MAN":     (1, 0, 0, "1405-08-15", ""),
    "CW-SEN":     (4, 1, 0, "1405-09-01", "کارگزار اعلام ۱۱ هفته"),
}
for i in range(PLAN_LAST - PLAN_FIRST + 1):
    rr = PLAN_FIRST + i
    code = BOM_ITEMS[i][0] if i < len(BOM_ITEMS) else None
    put(ws, rr, 1, code, a=CENTER, f=font(8.5, True, NAVY), fl=(pfill(GRAY) if code else pfill(YELLOW)),
        unlock=(code is None))
    put(ws, rr, 2, f"=IF($A{rr}=\"\",\"\",IFERROR(INDEX('مهندسی_BOM'!$C$8:$C$31,MATCH($A{rr},'مهندسی_BOM'!$B$8:$B$31,0)),\"⚠\"))", f=font(9), fl=pfill(GRAY))
    put(ws, rr, 3, f"=IF($A{rr}=\"\",\"\",IFERROR(INDEX('مهندسی_BOM'!$L$8:$L$31,MATCH($A{rr},'مهندسی_BOM'!$B$8:$B$31,0)),0))", a=CENTER, f=font(9), num="#,##0.0", fl=pfill(GRAY))
    inv = PLAN_DEMO.get(code, (None, None, None, None, None)) if code else (None,) * 5
    put(ws, rr, 4, inv[0], a=CENTER, f=font(9), num="#,##0.0", fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 5, inv[1], a=CENTER, f=font(9), num="#,##0.0", fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 6, f"=IF($A{rr}=\"\",\"\",MAX(0,IF($C{rr}=\"\",0,$C{rr})-IF($D{rr}=\"\",0,$D{rr})-IF($E{rr}=\"\",0,$E{rr})))", a=CENTER, f=font(9, True), num="#,##0.0", fl=pfill(GRAY))
    put(ws, rr, 7, inv[2] if code else 0, a=CENTER, f=font(9), num="#,##0.0", fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 8, inv[3], a=CENTER, f=font(9), num="@", fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 9, j2s_formula(f"$H{rr}"), a=CENTER, f=font(6), num=";;;")
    put(ws, rr, 10, f"=IF($A{rr}=\"\",\"\",IFERROR(INDEX('بازرگانی'!$S${TRD_FIRST}:$S${TRD_LAST},MATCH($A{rr},'بازرگانی'!$B${TRD_FIRST}:$B${TRD_LAST},0)),\"\"))", a=CENTER, f=font(9), num="0", fl=pfill(GRAY))
    put(ws, rr, 11, f"=IF(OR($I{rr}<=0,$J{rr}=\"\"),\"—\",IFERROR(INDEX(CalJ,MATCH($I{rr}-ROUND($J{rr},0),CalS,0)),\"خارج از تقویم\"))", a=CENTER, f=font(9, True, NAVY), fl=pfill("DDEBF7"))
    put(ws, rr, 12, f"=IF($A{rr}=\"\",\"\",IF(AND($F{rr}>0,$I{rr}>0,$I{rr}-TODAY()<=14),\"بحرانی\",IF($F{rr}<=0,\"تامین‌شده\",IF($E{rr}>0,\"در راه\",\"درخواست خرید\"))))", a=CENTER, f=font(8.5, True))
    put(ws, rr, 13, inv[4], f=font(8.5), fl=pfill(YELLOW), unlock=True)
trp = PLAN_LAST + 1
fill_range(ws, trp, trp, 1, 13, LBLUE)
put(ws, trp, 2, "جمع اقلام نیازمند سفارش", f=font(10, True), fl=pfill(LBLUE))
put(ws, trp, 6, f"=SUMIF($L${PLAN_FIRST}:$L${PLAN_LAST},\"درخواست خرید\",$F${PLAN_FIRST}:$F${PLAN_LAST})", a=CENTER, f=font(9, True), num="#,##0.0", fl=pfill(LBLUE))
put(ws, trp, 12, f"=COUNTIF($L${PLAN_FIRST}:$L${PLAN_LAST},\"بحرانی\")&\" بحرانی\"", a=CENTER, f=font(10, True, DRED), fl=pfill(LBLUE))
for st, hex_ in (("بحرانی", RED), ("درخواست خرید", AMBER), ("در راه", "DDEBF7"), ("تأمین‌شده", GREEN)):
    ws.conditional_formatting.add(f"L{PLAN_FIRST}:L{PLAN_LAST}",
        FormulaRule(formula=[f'$L{PLAN_FIRST}="{st}"'], fill=pfill(hex_)))
ws.freeze_panes = f"A{PLAN_FIRST}"
ws.auto_filter.ref = f"A6:M{PLAN_LAST}"
# ----------------------------------------------------------------------------
# مالی (cost / budget)
# ----------------------------------------------------------------------------
ws = SH["مالی"]
banner(ws, "مالی — بهای تمام‌شده و بودجه پروژه",
       "مواد از جمع برگه بازرگانی و نیروی کار از جمع نفرساعت مهندسی به‌صورت خودکار کشیده می‌شود. واحد: میلیون تومان.", 10)
widths(ws, {"A": 26, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16, "G": 12, "H": 12, "I": 12, "J": 12})
section(ws, 4, 1, 6, "الف) تشکیل بهای تمام‌شده (م.ت.)")
cost_rows = [
    ("مواد و تجهییزات خریداری‌شده (BOM)", "='بازرگانی'!$P$32", False),
    ("نیروی کار مستقیم (نفرساعت مهندسی/تولید)", "='مهندسی_BOM'!$G$43", False),
    ("سربار + اداری + ذخیره ریسک (٪)", 26, True),
]
for i, (lab, v, inp) in enumerate(cost_rows):
    rr = 5 + i
    put(ws, rr, 1, lab, f=font(10, True), fl=pfill(GRAY))
    merge(ws, rr, 1, rr, 2)
    put(ws, rr, 3, v, f=font(10, True if not inp else 10), a=CENTER, num="#,##0.0", fl=pfill(YELLOW) if inp else pfill(GRAY), unlock=inp)
put(ws, 8, 1, "جمع بهای تمام‌شده پروژه (م.ت.)", f=font(11, True, NAVY), fl=pfill(LBLUE))
merge(ws, 8, 1, 8, 2)
put(ws, 8, 3, "=ROUND((C5+C6)*(1+C7/100),1)", f=font(13, True, NAVY), a=CENTER, num="#,##0.0", fl=pfill(LBLUE))
FIN_TOTAL = "'مالی'!$C$8"
section(ws, 11, 1, 6, "ب) بودجه فازبندی‌شده و مصرف واقعی")
header_row(ws, 12, 1, ["فاز", "بودجه مصوب (م.ت.)", "هزینه واقعی", "انحراف", "٪ مصرف"], height=22)
PHASES = ["ثبت و امکان‌سنجی", "مهندسی", "تأمین", "تولید", "کیفیت و تست", "تحویل و پس از فروش"]
BUDGET = [180, 950, 8600, 2350, 320, 480]
for i, ph in enumerate(PHASES):
    rr = 13 + i
    put(ws, rr, 1, ph, f=font(9.5, True))
    put(ws, rr, 2, BUDGET[i], a=CENTER, f=font(10), num="#,##0", fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 3, f"=IFERROR(SUMIFS('هزینه‌واقعی'!$H:$H,'هزینه‌واقعی'!$C:$C,$A{rr}),0)", a=CENTER, f=font(10), num="#,##0.0", fl=pfill(GRAY))
    put(ws, rr, 4, f"=ROUND($C{rr}-$B{rr},1)", a=CENTER, f=font(9.5, True), num="#,##0.0;[Red]-#,##0.0", fl=pfill(GRAY))
    put(ws, rr, 5, f"=IF($B{rr}=0,0,$C{rr}/$B{rr})", a=CENTER, f=font(9.5), num="0.0%", fl=pfill(GRAY))
put(ws, 19, 1, "جمع", f=font(10, True), fl=pfill(LBLUE))
put(ws, 19, 2, "=SUM(B13:B18)", a=CENTER, f=font(11, True, NAVY), num="#,##0", fl=pfill(LBLUE))
put(ws, 19, 3, "=SUM(C13:C18)", a=CENTER, f=font(11, True, NAVY), num="#,##0.0", fl=pfill(LBLUE))
put(ws, 19, 4, "=ROUND(C19-B19,1)", a=CENTER, f=font(10, True), num="#,##0.0;[Red]-#,##0.0", fl=pfill(LBLUE))
put(ws, 19, 5, "=IF(B19=0,0,C19/B19)", a=CENTER, f=font(10, True), num="0.0%", fl=pfill(LBLUE))
ws.conditional_formatting.add("B19", FormulaRule(formula=["$C$8>$B$19"], fill=pfill(RED)))
ws.conditional_formatting.add("B19", FormulaRule(formula=["AND($C$8>0,$C$8<=$B$19)"], fill=pfill(GREEN)))
section(ws, 21, 1, 6, "ج) ملاحظات مالی")
put(ws, 22, 1, "مستندات پیوست (پیش‌فاکتورها، اسناد گمرکی، ریز دستمزد) در مسیر شبکه پروژه بایگانی شود. "
                "انحراف بیش از ۱۰٪ هر فاز، بازنگری الزامی گانت و امکان‌سنجی را در پی دارد (خط‌مشی گیت‌ها).",
    f=font(9), fl=pfill(NOTE_F))
merge(ws, 22, 1, 22, 10)
fill_range(ws, 22, 22, 1, 10, NOTE_F, bd=False)
ws.row_dimensions[22].height = 30

# fix finance total cell reference used by dashboard/فروش: keep at C8 but dashboard used B8 — alias B8
put(ws, 5, 6, "نمایش کلی", f=font(8, False, DARKTXT), bd=False)

# ----------------------------------------------------------------------------
# فروش (pricing)
# ----------------------------------------------------------------------------
ws = SH["فروش"]
banner(ws, "فروش — قیمت نهایی و شرایط پرداخت",
       "مبانی: بهای تمام‌شده از مالی؛ برآورد اولیه از شناسنامه. قیمت‌ها میلیون تومان.", 12)
widths(ws, {"A": 24, "B": 15, "C": 15, "D": 14, "E": 14, "F": 12, "G": 12, "H": 12, "I": 12, "J": 12, "K": 12, "L": 12})
rows_s = [
    ("بهای تمام‌شده (از مالی)", "='مالی'!$C$8", "#,##0.0", False),
    ("حاشیه سود هدف (٪)", 18, "0", True),
    ("قیمت فروش قبل از مالیات", "=ROUND($B$5*(1+$B$6/100),1)", "#,##0.0", False),
    ("مالیات بر ارزش افزوده (٪)", 10, "0", True),
    ("تخفیف ویژه (٪)", 2, "0", True),
    ("قیمت مشتری (برآورد شناسنامه)", "='شناسنامه'!$B$15", "#,##0", False),
    ("قیمت نهایی فروش با مالیات (م.ت.)", "=ROUND($B$7*(1+$B$8/100)*(1-$B$9/100),1)", "#,##0.0", False),
    ("قیمت نهایی بدون مالیات (م.ت.)", "=ROUND($B$11/(1+$B$8/100),1)", "#,##0.0", False),
    ("حاشیه نسبت به برآورد مشتری", "=IF($B$10=0,\"—\",IF($B$12<=$B$10,\"✓ در محدوده برآورد\",\"⚠ بالاتر از برآورد — نیازمند مذاکره\"))", None, False),
]
for i, (lab, v, num, inp) in enumerate(rows_s):
    rr = 5 + i
    put(ws, rr, 1, lab, f=font(10, True), fl=pfill(GRAY))
    put(ws, rr, 2, v, f=font(11 if rr in (11, 12) else 10, True), a=CENTER, num=num,
        fl=pfill(YELLOW) if inp else (pfill("E2EFDA") if rr >= 11 else pfill(GRAY)))
section(ws, 15, 1, 9, "شرایط پرداخت")
header_row(ws, 16, 1, ["مرحله پرداخت", "درصد", "مبلغ (م.ت.)", "تاریخ سررسید (شمسی)", "وضعیت", "رسید/چک", "توضیح"], height=22)
pay = [("پیش‌پرداخت قرارداد", 40, "1405-07-08"), ("تحویل ۵۰٪ اقلام تأمین", 25, "1405-08-25"),
       ("تحویل کالا در محل کارفرما", 30, "1405-11-10"), ("ضمانت‌نامه حسن انجام کار (۱۸ ماه)", 5, "1406-05-01")]
for i, (nm, pc, dtx) in enumerate(pay):
    rr = 17 + i
    put(ws, rr, 1, nm, f=font(9.5))
    put(ws, rr, 2, pc, a=CENTER, f=font(9.5), num="0", fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 3, f"=ROUND($B$11*$B{rr}/100,1)", a=CENTER, f=font(9.5, True), num="#,##0.0", fl=pfill(GRAY))
    put(ws, rr, 4, dtx, a=CENTER, f=font(9.5), num="@", fl=pfill(YELLOW), unlock=True)
    nrm = norm_expr("$D" + str(rr))
    put(ws, rr, 5, f"=IF($D{rr}=\"\",\"—\",IF(IFERROR(MATCH(" + nrm + ",CalJ,0),0)=0,\"⚠ تاریخ نامعتبر\",IF(INDEX(CalS,MATCH(" + nrm + ",CalJ,0))>TODAY(),\"نرسیده\",\"سررسید\")))", a=CENTER, f=font(9, True))


    put(ws, rr, 6, None, a=CENTER, f=font(9), fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 7, None, f=font(9), fl=pfill(YELLOW), unlock=True)
    merge(ws, rr, 7, rr, 9)
    put(ws, rr, 8, None, fl=pfill(YELLOW), unlock=True)
    put(ws, rr, 9, None, fl=pfill(YELLOW), unlock=True)
put(ws, 21, 1, "جمع درصدها", f=font(9.5, True), fl=pfill(GRAY))
put(ws, 21, 2, "=IF(SUM(B17:B20)=100,\"✓\",\"⚠ باید ۱۰۰ شود\")", a=CENTER, f=font(10, True))
ws.conditional_formatting.add("E17:E20", FormulaRule(formula=['$E17="سررسید"'], fill=pfill(AMBER)))

# ----------------------------------------------------------------------------
# جمع‌بندی مدیر پروژه (PM consolidation)
# ----------------------------------------------------------------------------
ws = SH["جمع‌بندی"]
banner(ws, "جمع‌بندی مدیر پروژه — بسته گزارش به مدیریت ارشد",
       "شاخص‌های کلیدی خودکارند؛ یادداشت و جمع‌بندی مدیر پروژه وارد می‌شود. پس از تکمیل، برگه «تصویب» برای نظر مدیریت ارشد باز است.", 9)
widths(ws, {"A": 30, "B": 16, "C": 18, "D": 16, "E": 16, "F": 16, "G": 16, "H": 16, "I": 16})
header_row(ws, 5, 1, ["شاخص کلیدی", "مقدار", "وضعیت", "منبع", "", "", "", "", ""], height=22)
pm_rows = [
    ("امتیاز امکان‌سنجی (از ۱۰۰)", "=IF('امکان‌سنجی'!$F$17=0,\"—\",'امکان‌سنجی'!$F$17)", "=IF('امکان‌سنجی'!$F$17>=ParamThresh,\"✓\",\"⚠\")", "برگه امکان‌سنجی"),
    ("پیشرفت واقعی زمان", f"=IF('گانت'!$J$45=\"\",\"—\",'گانت'!$J$45)", None, "گانت"),
    ("اختلاف با برنامه (پیشرفت)", "=IF(OR('گانت'!$J$45=\"\",'گانت'!$Q$45=\"\"),\"—\",'گانت'!$J$45-'گانت'!$Q$45)", "=IF(OR('گانت'!$J$45=\"\",'گانت'!$Q$45=\"\"),\"—\",IF('گانت'!$J$45-'گانت'!$Q$45>=-0.005,\"✓ مطابق برنامه\",\"⚠ عقب‌تر از برنامه\"))", "گانت"),
    ("مصرف بودجه", "=IF('مالی'!$B$19=0,\"—\",'مالی'!$C$19/'مالی'!$B$19)", "=IF('مالی'!$C$19-'مالی'!$B$19<=0,\"✓\",\"⚠ فراتر از بودجه\")", "مالی"),
    ("ریسک‌های بحرانی باز", "=COUNTIFS('ریسک'!$G$8:$G$27,\"بحرانی\",'ریسک'!$K$8:$K$27,\"باز\")", None, "ریسک"),
    ("قیمت نهایی فروش (م.ت.)", "=IF('فروش'!$B$12=0,\"—\",'فروش'!$B$12)", None, "فروش"),
]
for i, (lab, v, st, src) in enumerate(pm_rows):
    rr = 6 + i
    put(ws, rr, 1, lab, f=font(9.5, True))
    put(ws, rr, 2, v, a=CENTER, f=font(10, True, NAVY), num="0.0%" if "پیشرفت" in lab or "مصرف" in lab else "#,##0.0;;0", fl=pfill(GRAY))
    if st:
        put(ws, rr, 3, st, a=CENTER, f=font(9.5, True), fl=pfill(GRAY))
    else:
        put(ws, rr, 3, "—", a=CENTER, f=font(9.5), fl=pfill(GRAY))
    put(ws, rr, 4, src, a=CENTER, f=font(8.5, False, DARKTXT), fl=pfill(GRAY))
put(ws, 13, 1, "جمع‌بندی مدیر پروژه", f=font(10, True), fl=pfill(GRAY))
put(ws, 14, 1, None, fl=pfill(YELLOW), unlock=True)
merge(ws, 14, 1, 16, 9)
fill_range(ws, 14, 16, 1, 9, YELLOW)
put(ws, 14, 1, "وضعیت زمانی ۶٪ عقب از برنامه است که عمدتاً از تأخیر L/C تجهیزات خارجی (F2.2) ناشی می‌شود؛ جبران با موازی‌سازی عملیات قالب‌سازی و لایه‌نشانی پیش‌بینی شد. "
      "هزینه تأمین در سقف بودجه است؛ پیشنهاد: پیش‌پرداخت رزین به تأمین‌کننده مورد تأیید برای قفل قیمت. پرونده برای تصویب نهایی به مدیریت ارشد ارسال می‌شود.",
    f=font(9.5), a=al("right"), fl=pfill(YELLOW), unlock=True)
put(ws, 17, 1, "تاریخ ارسال به مدیریت ارشد", f=font(9.5, True), fl=pfill(GRAY))
put(ws, 17, 2, "1405/07/01", a=CENTER, f=font(9.5), num="@", fl=pfill(YELLOW), unlock=True)

# ----------------------------------------------------------------------------
# تصویب مدیریت ارشد
# ----------------------------------------------------------------------------
ws = SH["تصویب"]
banner(ws, "تصمیم مدیریت ارشد (Stage-Gate)",
       "مدیریت ارشد پس بررسی بسته جمع‌بندی، یکی از وضعیت‌های «مصوب / مصوب مشروط / نیازمند اصلاح / رد شده» را ثبت می‌کند. نتیجه خودکار در شناسنامه و داشبورد منعکس می‌شود.", 8)
widths(ws, {"A": 24, "B": 20, "C": 18, "D": 18, "E": 16, "F": 16, "G": 14, "H": 14})
d_rows = [("کد پروژه", "='شناسنامه'!$B$5", None),
          ("امتیاز امکان‌سنجی", "=IF('امکان‌سنجی'!$F$17=0,\"—\",'امکان‌سنجی'!$F$17)", "0.0"),
          ("پیشرفت واقعی", "='گانت'!$J$45", "0.0%"),
          ("بهای تمام‌شده (م.ت.)", "='مالی'!$C$8", "#,##0.0"),
          ("قیمت فروش پیشنهادی (م.ت.)", "='فروش'!$B$12", "#,##0.0"),
          ("جمع‌بندی مدیر پروژه", "='جمع‌بندی'!$A$14", None)]
for i, (lab, v, num) in enumerate(d_rows):
    rr = 5 + i
    put(ws, rr, 1, lab, f=font(10, True), fl=pfill(GRAY))
    put(ws, rr, 2, v, f=font(10, True, NAVY), a=CENTER, num=num, fl=pfill("DDEBF7"))
    merge(ws, rr, 2, rr, 3)
put(ws, 12, 1, "تصمیم مدیریت ارشد", f=font(11, True, WHITE), fl=pfill(NAVY))
put(ws, 12, 2, "مصوب مشروط", f=font(12, True, DGREEN), fl=pfill(YELLOW), unlock=True, a=CENTER)
set_dv_list(ws, "ListProj", "B12")
put(ws, 13, 1, "تاریخ تصمیم (شمسی)", f=font(10, True), fl=pfill(GRAY))
put(ws, 13, 2, "1405/07/01", a=CENTER, f=font(10), num="@", fl=pfill(YELLOW), unlock=True)
put(ws, 14, 1, "نظر/شرط مدیریت", f=font(10, True), fl=pfill(GRAY))
merge(ws, 14, 2, 14, 8)
put(ws, 14, 2, "مصوب مشروط: با شرط تهاتر ۵٪ تخیفیف در ازای کاهش دوره گارانتی به ۱۲ ماه و تکمیل بازنگری ریسک R-02 تا پایان مهر.",
    f=font(9.5), fl=pfill(YELLOW), unlock=True)
for cc in range(3, 9):
    ws.cell(row=14, column=cc).fill = pfill(YELLOW)
    ws.cell(row=14, column=cc).protection = Protection(locked=False)
put(ws, 15, 1, "امضای مدیرعامل", f=font(10, True), fl=pfill(GRAY))
merge(ws, 15, 2, 15, 3)
put(ws, 15, 2, None, fl=pfill(YELLOW), unlock=True)
ws.conditional_formatting.add("B12", FormulaRule(formula=['OR($B$12="مصوب",$B$12="مصوب مشروط")'], fill=pfill(GREEN), font=font(12, True, DGREEN)))
ws.conditional_formatting.add("B12", FormulaRule(formula=['OR($B$12="رد شده",$B$12="نیازمند اصلاح")'], fill=pfill(RED)))

# ----------------------------------------------------------------------------
# هزینه‌واقعی (actual cost ledger)
# ----------------------------------------------------------------------------
ws = SH["هزینه‌واقعی"]
banner(ws, "هزینه‌های واقعی پروژه (Turn-over / Actual Cost Ledger)",
       "ثبت بهای واقعی هر فاز به تفکیک واحد — مبنای محاسبه CPI و انحراف بودجه. واحد: میلیون تومان.", 10)
widths(ws, {"A": 5, "B": 12, "C": 20, "D": 34, "E": 14, "F": 16, "G": 12, "H": 13, "I": 14, "J": 16})
header_row(ws, 6, 1, ["ردیف", "تاریخ (شمسی)", "فاز", "شرح هزینه / سند", "مستند حسابداری",
                      "واحد ثبت‌کننده", "کارشناس", "مبلغ (م.ت.)", "وضعیت پرداخت", "محل بایگانی"])
AC_FIRST, AC_LAST = 8, 29
AC_DEMO = [
    ("1405-04-05", "ثبت و امکان‌سنجی", "برآورد اولیه و جلسه کمیته", "PP-1405/221", "مدیر پروژه", 12.5),
    ("1405-05-21", "مهندسی", "ساعت طراحی نقشه‌ها و محاسبات", "PP-1405/244", "مهندسی", 86.4),
    ("1405-06-05", "تأمین", "رزین ایزوفتالیک (۵۰٪)", "INV-1405/118", "بازرگانی داخلی", 612),
    ("1405-06-10", "تأمین", "الیاف MAT/ROV", "INV-1405/122", "بازرگانی داخلی", 441),
    ("1405-06-18", "تأمین", "پیش‌پرداخت ژل‌کوت وارداتی", "TTR-1405/19", "بازرگانی خارجی", 96.8),
    ("1405-06-25", "تأمین", "ارزیابی تأمین‌کننده و نمونه", "PP-1405/260", "بازرگانی داخلی", 18.2),
    ("1405-07-01", "مهندسی", "طراحی قالب — ویرایش B", "PP-1405/271", "مهندسی", 54.1),
    ("1405-07-10", "تأمین", "نازل و فلنج‌ها", "INV-1405/131", "بازرگانی داخلی", 108.4),
    ("1405-07-18", "تولید", "دستمزد آماده‌سازی خط", "PR-1405/77", "برنامه‌ریزی و انبار", 62.3),
    ("1405-07-25", "تأمین", "پمپ رزیون (۵۰٪ اول)", "TTR-1405/24", "بازرگانی خارجی", 1380.5),
]
for i in range(AC_LAST - AC_FIRST + 1):
    rr = AC_FIRST + i
    if i < len(AC_DEMO):
        d, ph, ttl, doc, unit, amt = AC_DEMO[i]
        put(ws, rr, 1, i + 1, a=CENTER, f=font(8.5))
        put(ws, rr, 2, d, a=CENTER, f=font(9), num="@")
        put(ws, rr, 3, ph, f=font(9), a=CENTER)
        put(ws, rr, 4, ttl, f=font(9))
        put(ws, rr, 5, doc, a=CENTER, f=font(9))
        put(ws, rr, 6, unit, a=CENTER, f=font(9))
        put(ws, rr, 7, "—", a=CENTER, f=font(9))
        put(ws, rr, 8, amt, a=CENTER, f=font(9.5, True), num="#,##0.0")
        put(ws, rr, 9, "پرداخت‌شده", a=CENTER, f=font(9))
        put(ws, rr, 10, "\\\\BORNASRV\\Projects\\...\\Finance", f=font(8, False, DARKTXT))
    else:
        for cc in range(1, 11):
            put(ws, rr, cc, None, fl=pfill(YELLOW), unlock=True, num="@" if cc == 2 else ("#,##0.0" if cc == 8 else None), f=font(9))
set_dv_list(ws, "ListPhase", f"C{AC_FIRST}:C{AC_LAST}")
set_dv_list(ws, "ListUnits", f"F{AC_FIRST}:F{AC_LAST}")
tr_ = AC_LAST + 1
fill_range(ws, tr_, tr_, 1, 10, LBLUE)
put(ws, tr_, 3, "جمع هزینه واقعی", f=font(10, True), fl=pfill(LBLUE))
put(ws, tr_, 8, f"=ROUND(SUM(H{AC_FIRST}:H{AC_LAST}),1)", a=CENTER, f=font(12, True, NAVY), num="#,##0.0", fl=pfill(LBLUE))
AC_TOTAL = "'هزینه‌واقعی'!$H$30"

# ----------------------------------------------------------------------------
# ریسک‌ها (risk register)
# ----------------------------------------------------------------------------
ws = SH["ریسک"]
banner(ws, "ثبت ریسک پروژه",
       "امتیاز = احتمال × اثر (هرکدام ۱ تا ۵). سطح‌بندی خودکار و رنگ‌آمیزی شرطی. بازبینی حداقل ماهانه توسط مدیر پروژه — مرجع: PMI PMBOK 7 و ISO 31000.", 14)
widths(ws, {"A": 8, "B": 40, "C": 12, "D": 9, "E": 9, "F": 9, "G": 11, "H": 12, "I": 34, "J": 14, "K": 11, "L": 12, "M": 12, "N": 10})
header_row(ws, 6, 1, ["کد", "شرح ریسک", "دسته", "احتمال", "اثر", "امتیاز", "سطح", "راهبرد",
                      "اقدام کاهش / پاسخ", "مسئول", "وضعیت", "آخرین بازبینی", "بازبینی بعد", ""])
RK_FIRST, RK_LAST = 8, 27
RK_DEMO = [
    ("R-01", "نوسان قیمت رزین و الیاف در بازار داخلی", "تأمین", 4, 4, "کاهش", "قفل قیمت با خرید فوری و قرارداد دوجانبه", "بازرگانی داخلی"),
    ("R-02", "تأخیر تخصیص ارز/گشایش اعتبارات اسنادی", "تأمین", 4, 5, "کاهش", "استعلام از کارگزار، جایگزین داخلی پمپ، بازنگری گانت", "بازرگانی خارجی"),
    ("R-03", "کیفیت قالب اسپیل (انحراف ابعادی)", "فنی", 3, 4, "کاهش", "کنترل سه‌بعدی قبل از لایه‌نشانی، Tolerance review", "مهندسی"),
    ("R-04", "تعارض برنامه خط تولید با پروژه", "سازمانی", 3, 3, "پذیرش", "برنامه‌ریزی بار خط، شیفت اضافه", "برنامه‌ریزی و انبار"),
    ("R-05", "حوادث HSE کار با رزین/کار در ارتفاع", "ایمنی", 2, 5, "اجتناب", "مجوز کار، PPE، آموزش اجباری، بیمه مسئولیت", "مدیر پروژه"),
    ("R-06", "ادعای وجه التزام تاخیر توسط کارفرما", "قرارداد", 3, 4, "کاهش", "صورت‌جلسه مکتوب تعلیقات، مستندسازی علل تأخیر در گانت", "مدیر پروژه"),
    ("R-07", "افزایش نرخ ارز و هزینه حمل بین‌المللی", "مالی", 3, 3, "انتقال", "بیمه باربری + قرارداد CIF + ذخیره ۵٪", "مالی"),
    ("R-08", "مستندسازی ناکافی مدارک تحویلی", "کیفیت", 2, 3, "کاهش", "چک‌لیست تحویل مدارک، ممیزی داخلی قبل از گیت ۳", "مهندسی"),
]
for i in range(RK_LAST - RK_FIRST + 1):
    rr = RK_FIRST + i
    if i < len(RK_DEMO):
        code, ttl, cat, p_, imp, strg, act, owner = RK_DEMO[i]
        put(ws, rr, 1, code, a=CENTER, f=font(9, True, NAVY))
        put(ws, rr, 2, ttl, f=font(9))
        put(ws, rr, 3, cat, a=CENTER, f=font(9))
        put(ws, rr, 4, p_, a=CENTER, f=font(9, True), fl=pfill(YELLOW), unlock=True, num="0")
        put(ws, rr, 5, imp, a=CENTER, f=font(9, True), fl=pfill(YELLOW), unlock=True, num="0")
        put(ws, rr, 6, f"=IF(OR($D{rr}=\"\",$E{rr}=\"\"),\"\",$D{rr}*$E{rr})", a=CENTER, f=font(10, True), num="0", fl=pfill(GRAY))
        put(ws, rr, 7, f"=IF($F{rr}=\"\",\"\",IF($F{rr}>=16,\"بحرانی\",IF($F{rr}>=9,\"بالا\",\"متوسط\")))", a=CENTER, f=font(9, True), fl=pfill(GRAY))
        put(ws, rr, 8, strg, a=CENTER, f=font(9), fl=pfill(YELLOW), unlock=True)
        put(ws, rr, 9, act, f=font(8.5), fl=pfill(YELLOW), unlock=True)
        put(ws, rr, 10, owner, a=CENTER, f=font(9))
        put(ws, rr, 11, "باز", a=CENTER, f=font(9, True), fl=pfill(YELLOW), unlock=True)
        put(ws, rr, 12, "1405/06/28", a=CENTER, f=font(9), num="@")
        put(ws, rr, 13, "1405/07/28", a=CENTER, f=font(9), num="@", fl=pfill(YELLOW), unlock=True)
    else:
        for cc in range(1, 14):
            put(ws, rr, cc, None, fl=pfill(YELLOW), unlock=True,
                num="0" if cc in (4, 5) else ("@" if cc in (12, 13) else None), f=font(9))
for fml, hex_ in ((f'$F{RK_FIRST}>=16', RED), (f'$F{RK_FIRST}>=9', AMBER)):
    ws.conditional_formatting.add(f"F{RK_FIRST}:F{RK_LAST}",
        FormulaRule(formula=[f"AND({fml},ISNUMBER($F{RK_FIRST}))"], fill=pfill(hex_)))
for rng_, txt, hex_ in ((f"K{RK_FIRST}:K{RK_LAST}", "باز", RED), (f"G{RK_FIRST}:G{RK_LAST}", "بحرانی", "F8CBAD")):
    ws.conditional_formatting.add(rng_, FormulaRule(formula=[f'ISNUMBER(SEARCH("{txt}",${rng_[0]}{RK_FIRST}))'], fill=pfill(hex_)))
set_dv_list(ws, "ListRiskSt", f"K{RK_FIRST}:K{RK_LAST}")
set_dv_list(ws, "ListRiskStrategy", f"H{RK_FIRST}:H{RK_LAST}")
ws.freeze_panes = f"B{RK_FIRST}"
ws.auto_filter.ref = f"A6:M{RK_LAST}"

# ----------------------------------------------------------------------------
# تغییرات (change requests)
# ----------------------------------------------------------------------------
ws = SH["تغییرات"]
banner(ws, "ثبت و کنترل تغییرات پروژه (CR)",
       "هر تغییر دامنه/زمان/هزینه در این جدول ثبت می‌شود. تغییرات تأییدشده: به‌روزرسانی گانت، BOM و بودجه الزامی است.", 10)
widths(ws, {"A": 9, "B": 12, "C": 16, "D": 44, "E": 11, "F": 12, "G": 26, "H": 12, "I": 12, "J": 16})
header_row(ws, 6, 1, ["کد CR", "تاریخ", "واحد متقاضی", "شرح تغییر", "اثر زمان (روز)", "اثر هزینه (م.ت.)",
                      "بررسی مدیر پروژه", "وضعیت", "تأیید مدیریت", "بایگانی"])
CR_FIRST, CR_LAST = 8, 22
CR_DEMO = [
    ("CR-01", "1405/05/25", "مهندسی", "افزایش ضخامت لایه‌ها در ناحیه نازل‌ها (+۱.۲mm)", 2, 128.5, "توجیه فنی دارد — اثر زمان قابل جذب", "تأییدشده", "—"),
    ("CR-02", "1405/06/12", "کارفرما", "تغییر محل خروجی از کف به دیواره", 6, 214.0, "نیازمند بازنگری نقشه و تست — در گیت بعدی", "درخواست جدید", "—"),
    ("CR-03", "1405/06/29", "بازرگانی خارجی", "جایگزینی پمپ با برند معادل هندی", 0, -385.0, "صرفه‌جویی؛ تایید فنی اخذ شد", "تأییدشده", "—"),
]
for i in range(CR_LAST - CR_FIRST + 1):
    rr = CR_FIRST + i
    if i < len(CR_DEMO):
        code, dtx, unit, ttl, timp, cimp, rev, st, app = CR_DEMO[i]
        put(ws, rr, 1, code, a=CENTER, f=font(9, True, NAVY))
        put(ws, rr, 2, dtx, a=CENTER, f=font(9), num="@")
        put(ws, rr, 3, unit, a=CENTER, f=font(9))
        put(ws, rr, 4, ttl, f=font(9))
        put(ws, rr, 5, timp, a=CENTER, f=font(9), num="0")
        put(ws, rr, 6, cimp, a=CENTER, f=font(9, True), num="#,##0.0;[Red]-#,##0.0")
        put(ws, rr, 7, rev, f=font(8.5))
        put(ws, rr, 8, st, a=CENTER, f=font(9, True))
        put(ws, rr, 9, app, a=CENTER, f=font(9))
        put(ws, rr, 10, "—", a=CENTER, f=font(9))
    else:
        for cc in range(1, 11):
            put(ws, rr, cc, None, fl=pfill(YELLOW), unlock=True,
                num="@" if cc == 2 else ("#,##0.0;[Red]-#,##0.0" if cc == 6 else None), f=font(9))
set_dv_list(ws, "ListChgSt", f"H{CR_FIRST}:H{CR_LAST}")
ws.conditional_formatting.add(f"H{CR_FIRST}:H{CR_LAST}", FormulaRule(formula=[f'$H{CR_FIRST}="تأییدشده"'], fill=pfill(GREEN), font=font(9, True, DGREEN)))
ws.conditional_formatting.add(f"H{CR_FIRST}:H{CR_LAST}", FormulaRule(formula=[f'$H{CR_FIRST}="رد شده"'], fill=pfill(RED), font=font(9, True, DRED)))
ws.auto_filter.ref = f"A6:J{CR_LAST}"

# ----------------------------------------------------------------------------
# ورود (login)
# ----------------------------------------------------------------------------
ws = SH["ورود"]
banner(ws, "ورود به سامانه — احراز هویت نقش کاربری",
       "این برگه دروازه ورود هر کاربر است. پس از نصب ماکرو (اسناد INSTALL)، فقط برگه‌های مجاز برای نقش شما باز می‌شود.", 8)
widths(ws, {"A": 22, "B": 26, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14})
section(ws, 4, 1, 3, "اطلاعات ورود")
put(ws, 5, 1, "نام کاربری", f=font(11, True), fl=pfill(GRAY))
put(ws, 5, 2, None, fl=pfill(YELLOW), unlock=True, f=font(11, True), a=CENTER)
set_dv_list(ws, "ListUsers", "B5")
put(ws, 6, 1, "رمز عبور", f=font(11, True), fl=pfill(GRAY))
put(ws, 6, 2, None, fl=pfill(YELLOW), unlock=True, f=font(11, True), a=CENTER)
put(ws, 5, 3, "نقش فعال", f=font(10, True), fl=pfill(GRAY))
put(ws, 6, 3, "مهمان — فقط همین برگه و راهنما", f=font(10, True, DRED), fl=pfill(GRAY), a=CENTER)
section(ws, 8, 1, 8, "راهنمای اجرا (پس از نصب ماکرو)")
steps = ["۱) نام کاربری و رمز را وارد کنید (جدول کاربران در برگه تنظیمات).",
         "۲) کلیدهای Alt+F8 → ماژول Borna_Login → Run.",
         "۳) برای خروج: Alt+F8 → Borna_Logout. ورود اضطراری مدیر سیستم: Borna_Admin.",
         "۴) در صورت نصب نبودن ماکرو: فایل بدون محدودیت دید باز می‌شود؛ ویرایش همچنان با قفل شیت محدود است."]
for i, s_ in enumerate(steps):
    put(ws, 9 + i, 1, s_, f=font(9.5), fl=pfill(NOTE_F), bd=True)
    merge(ws, 9 + i, 1, 9 + i, 8)
    for cc in range(2, 9):
        ws.cell(row=9 + i, column=cc).fill = pfill(NOTE_F)
r_roles = 15
put(ws, r_roles - 1, 1, "خلاصه سطوح دسترسی (مطابق ماتریس تنظیمات)", f=font(11, True, NAVY), fl=pfill(LORANGE))
fill_range(ws, r_roles - 1, r_roles - 1, 1, 8, LORANGE)
header_row(ws, r_roles, 1, ["نقش", "برگه‌های قابل ویرایش", "برگه‌های فقط‌نمایش", "مخفی"], height=20)
def role_view(role):
    w = [s for s in SHEET_ORDER if ACCESS[s].get(role) == "W"]
    r_ = [s for s in SHEET_ORDER if ACCESS[s].get(role) == "R"]
    h = [s for s in SHEET_ORDER if ACCESS[s].get(role, "-") == "-"]
    return w, r_, h
for i, role in enumerate(ROLE_CODES):
    rr = r_roles + 1 + i
    w, rv, hv = role_view(role)
    put(ws, rr, 1, ROLE_NAMES[role], f=font(9.5, True), a=CENTER)
    put(ws, rr, 2, "، ".join(x for x in w) or "—", f=font(8.5), fl=pfill(GREEN))
    put(ws, rr, 3, "، ".join(x for x in rv) or "—", f=font(8.5), fl=pfill("DDEBF7"))
    merge(ws, rr, 3, rr, 4)
    put(ws, rr, 4, None, fl=pfill("DDEBF7"))
    put(ws, rr, 5, "، ".join(x for x in hv) or "—", f=font(8.5, False, DRED), fl=pfill(RED))
    merge(ws, rr, 5, rr, 8)
    for cc in range(6, 9):
        ws.cell(row=rr, column=cc).fill = pfill(RED)
    ws.row_dimensions[rr].height = 30
put(ws, r_roles + len(ROLE_CODES) + 2, 1,
    "نکته امنیتی: قفل اکسل در سطح بازدارندگی است. برای حفاظت واقعی داده‌ها، فایل در پوشه شبکه با NTFS "
    "و ماتریس دسترسی واحدها مستقر شود — راهنما: docs/IT-DEPLOYMENT.md.",
    f=font(9, False, DRED), fl=pfill(NOTE_F))
merge(ws, r_roles + len(ROLE_CODES) + 2, 1, r_roles + len(ROLE_CODES) + 2, 8)
for cc in range(2, 9):
    ws.cell(row=r_roles + len(ROLE_CODES) + 2, column=cc).fill = pfill(NOTE_F)

# ----------------------------------------------------------------------------
# راهنما (guide)
# ----------------------------------------------------------------------------
ws = SH["راهنما"]
banner(ws, "سامانه مدیریت پروژه برنا کامپوزیت ایرانیان — راهنما (v1.4 شمسی)",
       "الگوی عمومی برای هر پروژه (سفارش/ایده/طرح) با گردش واحدی: ثبت → امکان‌سنجی → مهندسی → تأمین → برنامه‌ریزی → مالی → فروش → تصویب. واحد پول: میلیون تومان.", 10)
widths(ws, {"A": 16, "B": 20, "C": 18, "D": 18, "E": 16, "F": 16, "G": 14, "H": 14, "I": 14, "J": 14})
G2 = "0, ".join if False else None
guide_lines = [
    ("H1", "۱) گردش کار حاکم بر فایل"),
    ("T", "واحد بازار یابی/فروش (دریافت سفارش، ایده یا طرح) ← تکمیل «شناسنامه» ← مدیریت پروژه: «امکان‌سنجی» + «گانت» + تدوین الزامات با همکاری واحدها ← "
          "«مهندسی»: نقشه‌ها، Test Plan، BOM، نفرساعت ← «بازرگانی داخلی/خارجی»: قیمت‌گذاری اقلام BOM ← «برنامه‌ریزی تولید و انبار»: نیازسنجی و زمان تأمین ← "
          "«مالی»: بهای تمام‌شده و بودجه ← «فروش»: قیمت نهایی و شرایط پرداخت ← جمع‌بندی مدیر پروژه ← تصمیم مدیرعامل در «تصویب». "
          "در تمام مراحل، مدیر پروژه گانت و امکان‌سنجی را بازنگری و به‌روزرسانی می‌کند."),
    ("H1", "۲) شروع سریع"),
    ("T", "۱. برای هر پروژه یک کپی از این فایل در پوشه شبکه بسازید: \\\\BORNASRV\\Projects\\<کد پروژه>\\ — ۲. «شناسنامه» را تکمیل کنید — ۳. داده‌های نمونه (مخزن FRP) را با داده واقعی جایگزین کنید؛ "
          "سلول‌های زرد ورودی و سلول‌های طوسی خودکارند — ۴. در «گانت» تاریخ‌های شمسی را با قالب 1405/07/12 وارد کنید — ۵. در «تنظیمات» کاربران و رمزها را اصلاح کنید — "
          "۶. ماکروی دسترسی را طبق docs/INSTALL.md نصب کنید تا حالت نرم‌افزاری (ورود/خروج نقش) فعال شود."),
    ("H1", "۳) قوانین تاریخ شمسی"),
    ("T", "مبنای تبدیل، برگه «تقویم» است (سریال میلادی ↔ رشته شمسی). تاریخ‌ها فقط داخل بازه تقویم معتبرند؛ در صورت «خارج از تقویم» یا «تاریخ نامعتبر»، قالب ورودی را بررسی کنید. "
          "رقم فارسی یا لاتین فرقی نمی‌کند؛ جداکننده / ، - و فاصله پذیرفته می‌شود. روزهای تعطیل رسمی در «تنظیمات» قابل ویرایش‌اند و در محاسبه روزهای کاری اعمال می‌شوند (جمعه تعطیل)."),
    ("H1", "۴) رنگ‌بندی"),
    ("T", "زرد = ورودی کاربر | طوسی/آبی کم‌رنگ = فرمول (ویرایش نکنید) | سبز = تکمیل/موفق | نارنجی = در جریان/هشدار | قرمز = تأخیر/انحراف/ریسک بحرانی."),
    ("H1", "۵) نقش شیت‌ها"),
]
sh_desc = [
    ("راهنما", "راهنما و استانداردها", "همه"),
    ("ورود", "دروازه احراز هویت نقش", "همه"),
    ("شناسنامه", "ثبت سفارش/ایده + وضعیت پرونده", "واحد ثبت سفارش / مدیر پروژه"),
    ("امکان‌سنجی", "امتیازدهی Five-Case و گیت ۱", "مدیر پروژه (با همکاری واحدها)"),
    ("گانت", "WBS، مبنای زمان، پیشرفت واقعی، بازنگری", "مدیر پروژه"),
    ("داشبورد", "KPI، منحنی S، نمودارها، چک‌لیست گردش", "همه"),
    ("مهندسی_BOM", "BOM + نفرساعت و زمان‌سنجی", "مهندسی"),
    ("مهندسی_مدارک", "رجیستر نقشه‌ها و Test Plan", "مهندسی"),
    ("بازرگانی", "قیمت‌گذاری اقلام BOM، ارز و گمرک، زمان تأمین", "بازرگانی داخلی/خارجی"),
    ("برنامه‌ریزی", "نیاز خالص، موجودی/در‌راه، تاریخ درخواست خرید", "برنامه‌ریزی و انبار"),
    ("مالی", "بهای تمام‌شده و بودجه فازی", "مالی"),
    ("فروش", "قیمت نهایی و شرایط پرداخت", "فروش"),
    ("جمع‌بندی", "بسته گزارش مدیر پروژه به مدیرعامل", "مدیر پروژه"),
    ("تصویب", "تصمیم گیت نهایی (Stage-Gate)", "مدیریت ارشد"),
    ("هزینه‌واقعی", "دفتر هزینه واقعی (مبنای CPI)", "مالی / مدیر پروژه"),
    ("ریسک", "ثبت، امتیاز و پایش ریسک", "مدیر پروژه + واحدها"),
    ("تغییرات", "ثبت و کنترل تغییرات (CR)", "مدیر پروژه"),
    ("تقویم", "جدول تبدیل تاریخ (منبع فرمول‌ها)", "محرکه — ویرایش نشود"),
    ("تنظیمات", "پارامترها، کاربران، ماتریس دسترسی، تعطیلات", "مدیر سیستم / مدیر پروژه"),
]
header_row(ws, 14, 1, ["برگه", "کاربرد", "واحد مسئول"], height=20)
for i, (a, b, c) in enumerate(sh_desc):
    rr = 15 + i
    put(ws, rr, 1, a, f=font(9.5, True, NAVY))
    put(ws, rr, 2, b, f=font(9))
    merge(ws, rr, 2, rr, 4)
    put(ws, rr, 3, None); put(ws, rr, 4, None)
    put(ws, rr, 5, c, a=CENTER, f=font(9))
    merge(ws, rr, 5, rr, 8)
    for cc in (6, 7, 8): ws.cell(row=rr, column=cc)
r2 = 15 + len(sh_desc) + 2
put(ws, r2, 1, "۶) مراجع و استانداردهای طراحی", f=font(12, True, NAVY), fl=pfill(LORANGE), bd=True)
merge(ws, r2, 1, r2, 8)
fill_range(ws, r2, r2, 1, 8, LORANGE)
refs = [
    "PMI — A Guide to the PMBOK® Guide (7th ed.) و Practice Standard for Scheduling (3rd ed.): ساختار WBS، گیت‌ها، کنترل زمان.",
    "GAO-16-32G, Schedule Assessment Guide: چهار آزمون برنامه (کامل/ساختار درست/منبع‌دار/قابل‌اتکا) — مبنای طراحی ستون‌های گانت و وضعیت‌ها.",
    "HM Treasury — The Five-Case Guide (Strategic/Economic/Commercial/Financial/Delivery): مبنای برگه امکان‌سنجی.",
    "ISO 21502:2020 و ISO 31000:2018: حاکمیت پروژه و مدیریت ریسک (ثبت ریسک/اقدام/بازبینی).",
    "ASME RTP-1 و ASTM D2992 (مرجع فنی نمونه نمایشی مخازن FRP — مرتبط با ماهیت صنعتی شرکت).",
]
for i, s_ in enumerate(refs):
    rr = r2 + 1 + i
    put(ws, rr, 1, s_, f=font(9), fl=pfill(GRAY), bd=True)
    merge(ws, rr, 1, rr, 8)
r3 = r2 + len(refs) + 2
put(ws, r3, 1, "۷) محدودیت‌ها و امنیت", f=font(12, True, NAVY), fl=pfill(LORANGE), bd=True)
merge(ws, r3, 1, r3, 8)
fill_range(ws, r3, r3, 1, 8, LORANGE)
lim = [
    "قفل شیت/ساختار و رمز کاربران در اکسل، بازدارنده است نه رمزنگاری؛ داده در فرمت xlsx روی پوشه مشترک برای هر کسی که فایل را بخواند قابل استخراج است.",
    "برای حفاظت واقعی: استقرار فایل در Share با NTFS + ماتریس دسترسی واحدها (اسناد/IT-DEPLOYMENT.md)، قفل VBA Project، و در صورت نیاز جداسازی فایل هر واحد.",
    "پیش‌نمایش‌های تحت وب، تقویم شمسی/فرمت‌های فارسی را ممکن است ناقص نشان دهند؛ مرجع، Excel دسکتاپ است.",
    "فایل را با Excel 2016 به بالا باز کنید؛ محاسبات اختیاری منحنی S، برآورد خطی است و جایگزین کنترل هزینه کامل (EVM) محسوب نمی‌شود.",
    "بازه ستون‌های گانت ثابت و هم‌تراز با سال 1405 است؛ برای سال دیگر فایل را با اسکریپت build/build_workbook.py --year 1406 بازسازی کنید.",
]
for i, s_ in enumerate(lim):
    rr = r3 + 1 + i
    put(ws, rr, 1, "• " + s_, f=font(9), fl=pfill("FFF7F3"), bd=True)
    merge(ws, rr, 1, rr, 8)
    ws.row_dimensions[rr].height = 26
r4 = r3 + len(lim) + 2
put(ws, r4, 1, "۸) بازسازی و توسعه (برای تیم IT/سیستم‌ها)", f=font(12, True, NAVY), fl=pfill(LORANGE), bd=True)
merge(ws, r4, 1, r4, 8)
fill_range(ws, r4, r4, 1, 8, LORANGE)
put(ws, r4 + 1, 1, "منبع فایل یک اسکریپت Python است (build/build_workbook.py). پس از تغییر، بازسازی: "
    "python3 build/build_workbook.py --out out/Project_Management_Borna.xlsx — سپس نسخه‌گذاری و مستندسازی تغییرات الزامی است. ماژول VBA نقش‌محور در vba/modBorna.bas است.",
    f=font(9), fl=pfill(GRAY), bd=True)
merge(ws, r4 + 1, 1, r4 + 2, 8)
ws.row_dimensions[r4 + 1].height = 34

# guide rows placement above tables? keep simple: prepend flow text at rows 4..10
for i, (kind, txt) in enumerate(guide_lines[:9]):
    rr = 4 + i
    if kind == "H1":
        put(ws, rr, 1, txt, f=font(11, True, NAVY), fl=pfill(LORANGE), bd=True)
        merge(ws, rr, 1, rr, 8)
        fill_range(ws, rr, rr, 1, 8, LORANGE)
    else:
        put(ws, rr, 1, txt, f=font(9), fl=pfill(GRAY), bd=True)
        merge(ws, rr, 1, rr + 1, 8) if False else merge(ws, rr, 1, rr, 8)
        ws.row_dimensions[rr].height = 46
# ----------------------------------------------------------------------------
# named ranges, validation wiring, protection, save
# ----------------------------------------------------------------------------
SET = SH["تنظیمات"]
def add_name(nm, ref):
    wb.defined_names.add(DefinedName(nm, attr_text=ref))
    # also register as list name for DV (which uses the name itself)

add_name("CalS", CAL_S_RANGE)
add_name("CalJ", CAL_J_RANGE)
add_name("Hols", f"'تنظیمات'!$D${P_HOL + 1}:$D${HOL_ROWS[1]}")
add_name("HolsNorm", f"'تنظیمات'!$C${P_HOL + 1}:$C${HOL_ROWS[1]}")
add_name("ParamThresh", "'تنظیمات'!$B$8")
add_name("FX_USD", "'تنظیمات'!$B$6")
add_name("FX_EUR", "'تنظیمات'!$B$7")
add_name("ListUsers", f"'تنظیمات'!$A${U_ROWS[0]}:$A${U_ROWS[1]}")
add_name("BomCodes", "'مهندسی_BOM'!$B$8:$B$31")
for nm, (rr, c1, c2) in LIST_RANGES.items():
    add_name(nm, f"'تنظیمات'!${gcl(c1)}${rr}:${gcl(c2)}${rr}")

# data-validation list wiring for horizontal lists (DV needs vertical-ish single row; Excel accepts row ranges)
# fix: score5 DV references
# (ListScore5 already defined through LIST_RANGES as "score5")

# ---- workbook structure protection (tabs/visibility) ----
wb.security = WorkbookProtection(lockStructure=True, workbookPassword=PWD_STRUCTURE)

# ---- sheet-level protection: locked formulas, open yellow inputs ----
for nm in SHEET_ORDER:
    protect(SH[nm], PWD_USER)

# active sheet = راهنما
for nm in SHEET_ORDER:
    SH[nm].sheet_view.tabSelected = (nm == "راهنما")
SH["راهنما"].sheet_view.showGridLines = False
wb.active = 0

import os
os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)

# ---- final integrity guard: balanced formula parens/quotes ----
bad_cells = []
for _ws in wb.worksheets:
    for _row in _ws.iter_rows():
        for _c in _row:
            _v = _c.value
            if isinstance(_v, str) and _v.startswith("="):
                if _v.count("(") != _v.count(")") or _v.count('"') % 2:
                    bad_cells.append(f"{_ws.title}!{_c.coordinate}")
assert not bad_cells, f"unbalanced formulas: {bad_cells[:10]}"
print("formula integrity OK")

wb.save(args.out)
print("saved:", args.out)
print("sheets:", len(SHEET_ORDER), "| grid cols:", GRID_DAYS, "| cal rows:", cal_n)
