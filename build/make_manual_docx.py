# -*- coding: utf-8 -*-
"""ساخت «دستورالعمل اجرا و استقرار — سامانه برنا v1.4» به‌صورت فایل Word راست‌به‌چپ."""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1F, 0x4E, 0x79)
BLUE = RGBColor(0x2F, 0x55, 0x97)
ORANGE = RGBColor(0xC5, 0x5A, 0x11)
GRAY = RGBColor(0x40, 0x40, 0x40)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FA = "Tahoma"

PD = "۰۱۲۳۴۵۶۷۸۹"
def fa(n):
    return "".join(PD[int(d)] if d.isdigit() else d for d in str(n))

def _rtl_par(p):
    pPr = p._p.get_or_add_pPr()
    b = OxmlElement("w:bidi"); b.set(qn("w:val"), "1"); pPr.append(b)

def _fix_run(r, size=11, bold=False, color=GRAY, italic=False):
    r.font.name = FA; r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color
    rPr = r._r.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rPr.append(rf)
    rf.set(qn("w:ascii"), FA); rf.set(qn("w:hAnsi"), FA); rf.set(qn("w:cs"), FA)
    rtl = OxmlElement("w:rtl"); rtl.set(qn("w:val"), "0"); rPr.append(rtl)
    szCs = OxmlElement("w:szCs"); szCs.set(qn("w:val"), str(int(size * 2))); rPr.append(szCs)

def P(doc, text="", size=11, bold=False, color=GRAY, align="right", space=4, indent=0):
    p = doc.add_paragraph()
    _rtl_par(p)
    p.alignment = {"right": WD_ALIGN_PARAGRAPH.RIGHT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "left": WD_ALIGN_PARAGRAPH.LEFT}[align]
    pf = p.paragraph_format; pf.space_before = Pt(2); pf.space_after = Pt(space)
    if indent: pf.right_indent = Cm(indent)
    r = p.add_run(text); _fix_run(r, size, bold, color)
    return p

def H(doc, num, text, level=1):
    if level == 1:
        p = P(doc, f"{num}. {text}" if num else text, size=15, bold=True, color=NAVY, space=2)
        pPr = p._p.get_or_add_pPr()
        pbdr = OxmlElement("w:pBdr"); bot = OxmlElement("w:bottom")
        bot.set(qn("w:val"), "single"); bot.set(qn("w:sz"), "8"); bot.set(qn("w:color"), "2F5597")
        bot.set(qn("w:space"), "4"); pbdr.append(bot); pPr.append(pbdr)
        p.paragraph_format.space_before = Pt(14)
    else:
        P(doc, text, size=12, bold=True, color=BLUE, space=2)

def NOTE(doc, text, kind="warn"):
    t = doc.add_table(rows=1, cols=1); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fill = {"warn": "FFF2CC", "ok": "E2EFDA", "info": "DEEAF6"}[kind]
    c = t.cell(0, 0)
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:fill"), fill)
    c._tc.get_or_add_tcPr().append(shd)
    p = c.paragraphs[0]; _rtl_par(p); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run(text); _fix_run(r, 10, bold=(kind != "info"), color=ORANGE if kind == "warn" else GRAY)
    P(doc, "", size=2)

def TABLE(doc, headers, rows, widths=None, fsize=10):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers)); t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = t._tbl.tblPr
    bv = OxmlElement("w:bidiVisual"); tblPr.append(bv)
    for j, h in enumerate(headers):
        c = t.cell(0, j); c.text = ""
        p = c.paragraphs[0]; _rtl_par(p); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); _fix_run(r, fsize, bold=True, color=WHITE)
        shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:fill"), "2F5597")
        c._tc.get_or_add_tcPr().append(shd)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = t.cell(i + 1, j); c.text = ""
            p = c.paragraphs[0]; _rtl_par(p)
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if j == 0 or len(str(v)) > 25 else WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(v)); _fix_run(r, fsize)
        if i % 2 == 1:
            for j in range(len(headers)):
                shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:fill"), "F2F7FC")
                t.cell(i + 1, j)._tc.get_or_add_tcPr().append(shd)
    if widths:
        for j, w in enumerate(widths):
            for i in range(len(rows) + 1):
                t.cell(i, j).width = Cm(w)
    P(doc, "", size=4)
    return t

def STEP(doc, n, title, body=""):
    p = P(doc, "", size=1, space=0)
    p2 = doc.add_paragraph(); _rtl_par(p2); p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.paragraph_format.space_before = Pt(6); p2.paragraph_format.space_after = Pt(1)
    r1 = p2.add_run(f"گام {fa(n)} — "); _fix_run(r1, 11.5, bold=True, color=BLUE)
    r2 = p2.add_run(title); _fix_run(r2, 11.5, bold=True, color=GRAY)
    if body:
        P(doc, body, size=10.5, space=2, indent=0.6)

def CODE(doc, text):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pPr = p._p.get_or_add_pPr()
    b = OxmlElement("w:bidi"); b.set(qn("w:val"), "0"); pPr.append(b)
    r = p.add_run(text); _fix_run(r, 9.5, color=RGBColor(0x1B, 0x3B, 0x5C))
    r.font.name = "Consolas"
    rPr = r._r.get_or_add_rPr(); rf = rPr.find(qn("w:rFonts"))
    rf.set(qn("w:ascii"), "Consolas"); rf.set(qn("w:hAnsi"), "Consolas")
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:fill"), "F2F2F2")
    pPr.append(shd)
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(6)

def main(out_path):
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = FA; st.font.size = Pt(11)
    for s in doc.sections:
        s.page_width, s.page_height = Cm(21), Cm(29.7)
        s.top_margin = s.bottom_margin = Cm(1.9); s.left_margin = s.right_margin = Cm(2)
        sectPr = s._sectPr; b = OxmlElement("w:bidi"); sectPr.append(b)
    # فوتر با شماره صفحه
    fp = doc.sections[0].footer.paragraphs[0]; _rtl_par(fp); fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = fp.add_run(); _fix_run(fr, 9)
    fld = OxmlElement("w:fldSimple"); fld.set(qn("w:instr"), "PAGE")
    rr = OxmlElement("w:r"); tt = OxmlElement("w:t"); tt.text = "۱"; rr.append(tt); fld.append(rr)
    fp._p.append(fld)

    # ===== سربرگ =====
    t = doc.add_table(rows=2, cols=1); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = t.cell(0, 0)
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:fill"), "1F4E79")
    c._tc.get_or_add_tcPr().append(shd)
    p = c.paragraphs[0]; _rtl_par(p); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("سامانهٔ مدیریت پروژه — برنا کامپوزیت ایرانیان"); _fix_run(r, 18, bold=True, color=WHITE)
    c2 = t.cell(1, 0)
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:fill"), "2F5597")
    c2._tc.get_or_add_tcPr().append(shd)
    p = c2.paragraphs[0]; _rtl_par(p); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("دستورالعمل اجرا و استقرار  •  نسخهٔ ۱.۴  •  تهیه: مهر ۱۴۰۵"); _fix_run(r, 12, bold=True, color=WHITE)
    P(doc, "", size=6)

    P(doc, "این دستورالعمل نشان می‌دهد بستهٔ BornaPM_v1.4.zip را چطور اجرا کنید: از تست ۱۰ دقیقه‌ای روی یک "
            "سیستم، تا استقرار کامل روی شبکهٔ داخلی با دسترسی واقعی هر واحد.", size=11, color=GRAY)

    # ===== ۱ نمای کلی =====
    H(doc, "۱", "نمای کلی — دو مدل اجرا، یک فایل")
    TABLE(doc,
          ["", "مدل A — تک‌فایل (سریع)", "مدل B — شبکه‌ای کامل (توصیهٔ نهایی)"],
          [["فایل‌های موردنیاز", "Project_Management_Borna.xlsx فقط", "فایل مادر + ۷ فایل واحد (پوشهٔ out/split)"],
           ["نحوهٔ دسترسی", "صفحهٔ ورود + قفل شیت‌ها (ماکرو)", "فایل جدا برای هر واحد + ACL ویندوزی (دیوار واقعی)"],
           ["زمان استقرار", "حدود ۱۰ دقیقه", "نیم‌روز با واحد IT"],
           ["مناسب برای", "شروع کار، تست، تیم‌های کوچک", "استقرار رسمی شرکتی، چند پروژهٔ هم‌زمان"]],
          widths=[3.6, 6.3, 6.3])
    NOTE(doc, "اگر تازه شروع می‌کنید: هفتهٔ اول با مدل A کار کنید و هم‌زمان IT پوشه‌ها و گروه‌های مدل B را "
              "آماده کند؛ مهاجرت بین دو مدل فقط با کپی‌کردن فایل‌ها انجام می‌شود، داده‌ای از دست نمی‌رود.", "info")

    # ===== ۲ محتویات بسته =====
    H(doc, "۲", "محتویات بسته")
    TABLE(doc,
          ["مسیر", "چیست", "برای چه کسی"],
          [["out/Project_Management_Borna.xlsx", "فایل مادر — ۱۹ شیت کامل با فرمول، داشبورد، تقویم شمسی و دمو", "مدیر پروژه (PMO)"],
           ["out/split/01_INTAKE … 07_EXEC", "فرمِ واحد هر بخش (فقط ورودی خودش + اسنپ‌شات مخفی بالادستی)", "کاربران همان واحد"],
           ["vba/modBorna.bas", "صفحهٔ ورود و تخصیص خودکار دسترسی (ماکرو فایل مادر)", "فقط برای فایل مادر"],
           ["vba/modBornaSplit.bas", "چرخهٔ جمع‌وانتشار داده بین فایل مادر و فایل واحدها", "فقط برای فایل مادر"],
           ["vba/ThisWorkbook.txt", "کدهای رویداد باز/بسته‌شدن فایل", "فقط برای فایل مادر"],
           ["docs/INSTALL.md — IT-DEPLOYMENT.md", "توضیحات کامل نصب و استقرار شبکه", "PMO و IT"],
           ["tools/deploy_acl.ps1", "اسکریپت ساخت گروه‌های AD و سطح دسترسی پوشه‌ها", "IT"],
           ["build/*.py + requirements.txt", "کد سازندهٔ فایل‌ها (برای سفارشی‌سازی و رمزهای اختصاصی)", "در صورت نیاز"]],
          widths=[6.2, 7.4, 2.6], fsize=9.5)

    # ===== ۳ تست سریع =====
    H(doc, "۳", "گام صفر — تست ۱۰ دقیقه‌ای (قبل از هر تصمیمی)")
    STEP(doc, 1, "دانلود و Extract", "فایل BornaPM_v1.4.zip را باز (Extract All) کنید. به اینترنت نیاز فقط همین یک‌بار است؛ بعد از آن همه‌چیز آفلاین کار می‌کند.")
    STEP(doc, 2, "باز‌کردن فایل مادر", "فایل out/Project_Management_Borna.xlsx را باز کنید. اکسل ۲۰۱۶ به‌بعد یا Microsoft 365 روی ویندوز — هیچ افزونه یا نصبی لازم نیست.")
    STEP(doc, 3, "گردش دمو را ببینید", "شیت «شناسنامه» اطلاعات پروژهٔ نمونهٔ مخزن FRP 150m³ را نشان می‌دهد؛ در «داشبورد» نمودار S-Curve و در «گانت» میله‌های برنامه/تحقق شمسی را نگاه کنید. پایین هر شیت ورودی، راهنمای یک‌خطی نوشته شده است.")
    STEP(doc, 4, "قفل‌ها را امتحان کنید", "روی یک سلول طوسی در شیت «مالی» دابل‌کلیک کنید — اکسل اجازهٔ ویرایش نمی‌دهد؛ سلول‌های زرد آزادند. این یعنی مدل قفل‌گذاری درست کار می‌کند.")
    NOTE(doc, "ورودی‌ها فقط سلول‌های زرد هستند؛ ستون‌های طوسی/آبی فرمول‌اند. اگر جای اشتباهی تایپ کردید، Ctrl+Z کافی است.", "ok")

    # ===== ۴ مدل A =====
    H(doc, "۴", "مدل A — اجرای تک‌فایل روی شبکه")
    STEP(doc, 1, "انتقال به شبکه", "فایل مادر را در پوشهٔ اشتراکی پروژه کپی کنید، مثلاً: \\\\BORNASRV\\Projects\\PRJ-1404-01\\Project_Management_Borna.xlsx")
    STEP(doc, 2, "فعال‌سازی صفحهٔ ورود (ماکرو)", "فایل را باز کنید → File → Save As → فرمت «Excel Macro-Enabled Workbook (*.xlsm)» → کلید Alt+F11 → منوی File → Import File → فایل vba/modBorna.bas را انتخاب کنید → همین کار را برای vba/modBornaSplit.bas هم بکنید → محتوای vba/ThisWorkbook.txt را در پنجرهٔ ThisWorkbook پیست کنید → ذخیره (Ctrl+S).")
    STEP(doc, 3, "رفع نوار زرد ماکرو", "File → Options → Trust Center → Trust Center Settings → Trusted Locations → Add new location → مسیر پوشهٔ شبکه را با تیک «Subfolders» ثبت کنید. (یا روی نوار زرد: Enable Content — برای هر بار بازکردن)")
    STEP(doc, 4, "ورود کاربران را ست کنید", "شیت «تنظیمات» جدول کاربران نمونهٔ زیر را دارد. برای هر کاربر واقعی یک ردیف بگذارید و ردیف‌های دمو را پاک کنید.")
    TABLE(doc, ["نام کاربری", "رمز ورود", "ردیف / نقش"],
          [["safa", "Borna-Intake1", "واحد دریافت/بازاریابی — INTAKE"], ["pm1", "Borna-PM1", "مدیر پروژه — PM (نمایش کل، ویرایش بخش خود)"],
           ["eng1", "Borna-Eng1", "مهندسی — ENG"], ["trd1 / trd2", "Borna-Trd1 / Trd2", "بازرگانی داخلی / خارجی"],
           ["plan1", "Borna-Plan1", "برنامه‌ریزی و انبار"], ["fin1", "Borna-Fin1", "مالی"],
           ["sales1", "Borna-Sales1", "فروش"], ["ceo", "Borna-Ceo1", "مدیرعامل — فقط «تصویب»"],
           ["admin", "Borna-Admin1", "واحد IT"]],
          widths=[3.2, 4.0, 9.0], fsize=9.5)
    STEP(doc, 5, "رمزهای قفل را عوض کنید (اختیاری ولی توصیه)", "پیش‌فرض‌ها: قفل شیت Borna@1405 و قفل ساختار Borna-Admin-1405 — در Review → Unprotect Sheet / Protect Workbook قابل تغییر است؛ یا فایل را با دستور §۷ و رمزهای اختصاصی بازسازی کنید. در حالت چندفایلی، رمز فایل واحدها در out/split/_ACL_HINTS.txt است (فقط IT/PMO).")
    STEP(doc, 6, "پاکسازی دمو برای شروع واقعی", "ردیف‌های نمونه در شیت‌های شناسنامه، مهندسی_BOM (ردیف ۸ تا ۲۳)، بازرگانی، برنامه‌ریزی، مالی، گانت و … را پاک کنید؛ ساختار و فرمول‌ها را دست نزنید. برای اطمینان یک کپی از فایلِ خامِ دمو نگه دارید.")
    NOTE(doc, "مدل A «بازدارنده» است نه دیوار امنیتی: کاربری که بخواهد می‌تواند با تغییر مسیر ماکرو یا کپی سلول‌ها از قفل‌ها عبور کند. برای دسترسی واقعی، مدل B را اجرا کنید.", "warn")

    # ===== ۵ مدل B =====
    H(doc, "۵", "مدل B — دسترسی واقعی روی شبکه (چندفایلی + ACL)")
    P(doc, "در این مدل هر واحد فقط فایل خودش را باز می‌کند؛ فایل او فقط ورودی خودش را دارد و بقیهٔ واحدها برایش وجود خارجی ندارند. مرز واقعی را ویندوز (NTFS) نگهبانی می‌کند نه اکسل.", size=10.5)
    H(doc, "", "۵-۱  کارهای IT (یک‌بار)", 2)
    STEP(doc, 1, "ساختار پوشه‌ها", "روی Share یک پروژه بسازید: \\master برای فایل مادر و \\units زیر پوشهٔ 01_INTAKE … 07_EXEC و در هر پوشه، فایل Form_<واحد>.xlsx مربوطه را از out/split کپی کنید.")
    STEP(doc, 2, "گروه‌ها و سطح دسترسی", "فایل tools/deploy_acl.ps1 را باز کنید، پارامترهای بالا (نام Share و گروه‌ها) را ست کنید، سپس در PowerShell با Run as Administrator اجرا کنید؛ اسکریپت گروه‌های AD را می‌سازد و NTFS را ست می‌کند: هر گروه Modify پوشهٔ خودش، بدون دسترسی به بقیه؛ Borna-PMO Modify همه‌جا (شامل master). دستور نمونه در docs/IT-DEPLOYMENT.md §۴ آمده است.")
    STEP(doc, 3, "قفل ساختار فایل واحدها", "در فایل هر واحد، رمز ساختار «توکن همان واحد» است (جدول توکن‌ها: out/split/_ACL_HINTS.txt) — بنابراین کاربر واحد حتی نمی‌تواند شیت‌های مخفی اسنپ‌شات را آشکار کند. این فایل برای IT/PMO است؛ دسترس‌پذیری آن را محدود نگه دارید.")
    H(doc, "", "۵-۲  کارهای مدیر پروژه (یک‌بار + روزانه)", 2)
    STEP(doc, 1, "ماکروها را روی فایل مادر Import کنید", "مطابق §۴ گام ۲ (modBorna.bas و modBornaSplit.bas را وارد و فایل را xlsm کنید).")
    STEP(doc, 2, "راه‌اندازی اتصال", "ماکروی Borna_SplitSetup را اجرا کنید (تب View → Macros). مسیر Share را که می‌پرسد وارد می‌کند؛ جدول «اتصالات شبکه‌ای» در شیت تنظیمات با مسیر و توکن فایل واحدها پر می‌شود. فایل را ذخیره کنید.")
    STEP(doc, 3, "چرخهٔ کاری هر هفته/روز", "وقتی واحدها کارشان را ذخیره کردند، ماکروی Borna_Sync را اجرا کنید (= Borna_Collect گرفتن ورودی‌ها + بازخوانی محاسبات مادر + Borna_Publish تازه‌سازی اسنپ‌شات‌ها). در پروژه‌های حساس، بین دو مرحله یک بار فایل مادر را ذخیره کنید.")
    NOTE(doc, "تا زمانی که Publish نشده، واحدها اعداد بالادستیِ آخرین انتشار را می‌بینند (این در راهنمای همان فایل هم نوشته شده). یک ریتم ثابت — مثلاً روزانه ساعت ۱۴ — مشکل را حذف می‌کند.", "info")

    # ===== ۶ گردش کار =====
    H(doc, "۶", "گردش کار عملیاتی — هر واحد چه کند")
    TABLE(doc, ["ترتیب", "واحد / شیت", "ورودی (سلول‌های زرد)", "خروجی خودکار"],
          [["۱", "ثبت درخواست — «شناسنامه» (INTAKE)", "مشخصات سفارش، کارفرما، بودجهٔ اولیه، تاریخ‌های شمسی", "ستون وضعیت + پیوند به داشبورد"],
           ["۲", "امکان‌سنجی (PM)", "امتیازات فنی/مالی/زمانی و مبنای هر امتیاز", "نتیجه: مجاز / مشروط / رد + بازخورد"],
           ["۳", "مهندسی — BOM و مدارک (ENG)", "اقلام فنی، ضرایب ضایعات، نفرساعت، وضعیت نقشه‌ها", "مقدار نهایی، پوشش قیمت‌گذاری، هشدار مدارک تأییدنشده"],
           ["۴", "بازرگانی (TRD)", "انتخاب کد فنی و قیمت واحد/ارز", "مبلغ هر ردیف و جمع مواد به مالی"],
           ["۵", "برنامه‌ریزی (PLAN)", "موجودی، زمان تحویل، پارامترهای سفارش", "نیاز خالص، تاریخ سفارش‌گذاری، وضعیت تأمین"],
           ["۶", "مالی (FIN)", "سربار، R&D، سایر بهاهای واقعی", "بهای تمام‌شده + انحراف هزینه‌واقعی"],
           ["۷", "فروش (SALES)", "حاشیهٔ هدف، تخفیف، شرایط پرداخت", "قیمت فروش + کف قیمت خودکار"],
           ["۸", "جمع‌بندی + گانت (PM)", "صورت‌جلسه، وظایف، تحقق روزانه", "KPIها، S-Curve، انحراف زمان"],
           ["۹", "تصویب (CEO)", "مصوب / مشروط / رد + یادداشت", "بازگشت نتیجه به شناسنامه و داشبورد"]],
          widths=[1.4, 4.6, 5.6, 4.6], fsize=9)
    NOTE(doc, "قانون طلایی: هر واحد فقط شیت خودش را ویرایش می‌کند. تاریخ‌ها را متنی شمسی وارد کنید (مثلاً 1405/04/31) — فرمت‌های 1405-04-31 و 31/04/1405 هم پذیرفته می‌شود؛ محاسبات از تقویم ۱۴۰۴ تا ۱۴۰۸ انجام می‌شود.", "info")

    # ===== ۷ سفارشی‌سازی =====
    H(doc, "۷", "بازسازی فایل‌ها با رمز/سال اختصاصی (اختیاری)")
    P(doc, "روی یک سیستم با Python 3.10+ (این مرحله فقط برای تولید فایل جدید است، نه اجرا):", size=10.5)
    CODE(doc, "pip install -r requirements.txt\n"
              "python build/build_workbook.py --out out/Project_Management_Borna.xlsx --pwd-user <رمز-قفل-شیت‌ها> --pwd-structure <رمز-قفل-ساختار>\n"
              "python build/split_deploy.py --master out/Project_Management_Borna.xlsx --out out/split")
    P(doc, "دستور دوم، فایل ۷ واحد را با همان رمزهای تازه و اسنپ‌شاتِ به‌روز تولید می‌کند. برای تقویم سال جدید: --year 1409", size=10.5)

    # ===== ۸ عیب‌یابی =====
    H(doc, "۸", "سؤالات و مشکلات متداول")
    TABLE(doc, ["مشکل", "راه‌حل"],
          [["نوار زرد «Enable Content» می‌آید", "پوشه را Trusted Location کنید (§۴ گام ۳)."],
           ["«ورود» کار نمی‌کند / پیغام VBA", "فایل باید xlsm باشد و modBorna.bas Import شده باشد؛ تنظیمات → Trust Center → دسترسی Trust access to the VBA Project را غیرفعال بگذارید."],
           ["تاریخ واردشده محاسبه نمی‌شود", "باید در بازهٔ تقویم باشد (۱۴۰۴ تا ۱۴۰۸) و جداکننده / یا - یکسان باشد؛ اگر خارج از بازه است فایل را با --year 1409 بازسازی کنید."],
           ["سلول طوسی قفل است", "طراحی همین است؛ برای ویرایش قانونی، رمز قفل شیت را از ادمین بگیرید (پیش‌فرض: Borna@1405)."],
           ["واحد می‌گوید اعداد بالادستی قدیمی است", "در مدل B یعنی Publish انجام نشده — PM باید Borna_Sync اجرا کند."],
           ["فایل واحدی شیت مخفی نشان می‌دهد", "Review → Unprotect Sheet با توکن همان واحد؛ فقط برای IT/PMO."],
           ["اکسل می‌گوید فایل Read-only است", "ACL پوشه را چک کنید: گروه همان واحد باید Modify روی فایل خودش داشته باشد نه فقط Read."] ],
          widths=[6.5, 9.7], fsize=9.5)

    # ===== ۹ چک‌لیست =====
    H(doc, "۹", "چک‌لیست پیش از افتتاح رسمی")
    for item in ["رمزهای دمو (Borna@1405 و Borna-Admin-1405) با رمز اختصاصی تعویض شده‌اند",
                 "جدول کاربران، افراد واقعی با رمز قوی‌تر هستند و دموها پاک شده‌اند",
                 "داده‌های نمونه از شیت‌های ورودی پاک شده و فقط ساختار مانده است",
                 "در مدل B: پوشه‌ها ساخته شده، ACL اعمال شده و Borna_SplitSetup اجرا شده است",
                 "یک چرخهٔ کامل آزمایشی اجرا شده: ثبت ← … ← تصویب ← بازگشت نتیجه",
                 "نسخهٔ پشتیبان از فایل خامِ ویرایش‌نشده در پوشه‌ای جدا نگهداری می‌شود"]:
        p = doc.add_paragraph(); _rtl_par(p); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(3); p.paragraph_format.left_indent = Cm(0.5)
        r = p.add_run("☐  " + item); _fix_run(r, 10.5)
    P(doc, "", size=6)
    NOTE(doc, "پشتیبانی: مستندات کامل در داخل بسته — docs/INSTALL.md (مراحل دقیق نصب) و docs/IT-DEPLOYMENT.md (ماتریس دسترسی شبکه). برای بازسازی یا افزودن شیت جدید: build/build_workbook.py.", "info")

    doc.save(out_path)
    print("saved:", out_path, round(os.path.getsize(out_path) / 1024, 1), "KB")

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "docs/Borna_Run_Guide_v1.4.docx")
