# 01 — نقشه فرآیند، Workflow و نقاط ابهام
**سامانه: مدیریت فرآیند «دریافت سفارش تا تصویب قیمت فروش» (Order-to-Quotation)**
نسخه: 1.0 — تاریخ: 2026-09-22

---

## 1.1 نقشه فرآیند (Process Map)

زبانه‌ها (Swimlanes): بازاریابی/فروش، مدیریت پروژه، مهندسی، برنامه‌ریزی، بازرگانی، مالی، فروش/قیمت‌گذاری، مدیریت ارشد.

```
[فروش]   ثبت سفارش  ≡  کنترل اعتبار مشتری/محصول/مشخصات  ──ناقص──► Return (REVISED)  + Issue
              │ تکمیل
              ▼
[مدیر پروژه]  ایجاد پروژه (Project_ID) + تیم + ضرب‌الاجل   ──►  Approved (P1)
              │
              ▼
[مهندسی]  ساخت/به‌روزرسانی Product Revision + Drawing + BOM (چندسطحی) ⇒ BOM_Revision
              │
              ▼                 (بدون BOM_Revision معتبر مسدود است)
[مهندسی]  تأیید BOM (Approve)  ──►  BOM Release
              │
              ▼
[برنامه‌ریزی]  برآورد نیاز (BOM × مقدار سفارش × (1+Scrap))، موجودی/کسری/مقدار خرید
              │          + زمان‌بندی (Schedule) با وابستگی‌ها
              ▼
[بازرگانی]  استعلام چند تأمین‌کننده (Supplier_Quotes)  ⇒  مقایسه + انتخاب توسط مسئول مجاز
              │          + ثبت قیمت خرید (منبع + تاریخ + ارز + تأمین‌کننده الزامی)
              ▼
[مالی]      Price Build-up → Landed Cost  ⇒  Costing (نسخه‌بندی‌شده)  ⇒  Budget vs Actual
              │
              ▼
[فروش]      ساخت پیشنهاد فروش: Cost + Margin + Commercial + Risk (سناریو A/B/C)
              │          + شرایط پرداخت + هزینه مالی اعتبار
              ▼
[مدیریت ارشد]  بررسی یکجا (Dashboard/Executive)  ⇒  APPROVED / WITH CONDITION / REVISION / REJECTED
              │
              └─تغییر هر ورودی تأییدشده → Revision جدید (بدون Overwrite) → چرخه مجدد از مرحله مربوط
```

هر فلش یک «انتقال وضعیت» است که در `Workflow_History` (چه کسی، چه زمانی، از چه وضعیتی به چه وضعیتی) ثبت می‌شود.

### اصول حاکم
1. **گِیتینگ (Gating):** هیچ مرحله‌ای بدون تکمیل پیش‌نیازها قابل عبور نیست. هر گیت یک ستون «قفل» محاسباتی است (نه فقط توصیه).
2. **نسخه‌بندی:** همه چیز (Product, BOM, Price, Costing, Quotation, Budget, Approval) Revision دارد؛ ثبت تأییدشده هرگز Overwrite نمی‌شود، فقط Revision جدید ساخته می‌شود.
3. **ردیابی (Traceability):** `Selling Price ← Costing ← Cost_Line ← Procurement ← Purchase_Price ← Supplier_Quote ← Material ← BOM_Line ← Product/Order`. ملاحظه: مسیر کامل دوطرفه در شیت `_Trace` پیاده‌سازی شده است.
4. **Audit:** هر انتقال، هر قیمت، هر تصمیم دارای «تاریخ + کاربر + منبع» است. جدول `Workflow_History` ستون فقرات Audit است.
5. **هیچ عدد سازمانی اختراع نمی‌شود.** همه نرخ‌ها (Overhead Rate، نرخ ارز، هزینه مالی، استهلاک قالب و …) از شیت `Settings` خوانده می‌شوند و مقادیر پیش‌فرض آن‌ها به‌صراحت `ASSUMPTION` هستند تا کاربر سازمانی جایگزین کند.

---

## 1.2 دنباله وضعیت‌ها (کارگردانی Workflow)

| Order | مرحله | مالک (نقش) | وضعیت‌های مجاز | پیش‌نیاز برای عبور |
|---|---|---|---|---|
| 1 | NEW — ثبت سفارش | Sales | Draft → In Review | Customer_ID، Product_ID، Qty>0، تاریخ، فروشنده |
| 2 | SALES REVIEW — کنترل اولیه | Sales | In Review → Approved / Revised | کنترل‌های CO-01..CO-07 سبز |
| 3 | PROJECT CREATED | PMO/PM | Pending → Approved | Order تأییدشده؛ Project_ID یکتا |
| 4 | ENGINEERING | Engineering | In Progress → BOM Submitted | Product Revision + Drawing Rev |
| 5 | BOM APPROVAL | Engineering Lead | Waiting Approval → Approved / Rejected | BOM_Revision معتبر، حداقل یک خط BOM، سطح‌بندی درست |
| 6 | PLANNING | Planning | Pending → Completed | BOM Released + Qty سفارش |
| 7 | PROCUREMENT (استعلام/انتخاب) | Procurement | In Progress → Completed | نیاز (Planning) معلوم؛ قیمت با منبع+تاریخ+ارز+تأمین‌کننده |
| 8 | COSTING | Finance | In Progress → Approved | Purchase_Prices فعال (غیرمنقضی) برای همه اقلام Make=Buy |
| 9 | FINANCE REVIEW | Finance | Waiting Approval → Approved | Costing Approved؛ Budget ثبت‌شده |
| 10 | SALES QUOTATION | Sales | Draft → Submitted | Costing تأییدشده؛ سناریو با مبنا |
| 11 | MANAGEMENT APPROVAL | Management | Waiting Approval → APPROVED / WITH CONDITION / REVISION / REJECTED | پیشنهاد با نسخه جاری + شرایط پرداخت + ریسک‌ها |

وضعیت نهایی سه‌گانهٔ مدیریت ارشد: **APPROVED / APPROVED WITH CONDITION / RETURN FOR REVISION / REJECTED** (چهار گزینه طبق الزام).

---

## 1.3 نقاط ضعف/ابهام فرآیند (صریحاً اعلام می‌شود)

| # | ابهام / ضعف | ریسک | تصمیم ما در این نسخه | چه اطلاعاتی برای تکمیل لازم است |
|---|---|---|---|---|
| 1 | **چندارزی** و نرخ تبدیل در جمع‌ها | جمع قیمت‌ها با ارزهای متفاوت نادرست می‌شود | ارز گزارشگری واحد (IRR در دمو)؛ هر قیمت دارای Currency و Fx؛ تبدیل شفاف با `IF(Currency="USD", Fx, 1)` | نرخ ارز رسمی سازمانی و تقویم نرخ‌ها |
| 2 | **سربار تولید / نرخ DM/Labor** | عدد اختراع = جعل منطق | همه نرخ‌ها ASSUMPTION در `Settings` (قابل ویرایش)، هیچ نرخی Hard-code نیست | نرخ واقعی سربار، نرخ دستمزد، روش جذب |
| 3 | **هزینه مالی اعتبار فروش** | نبود نرخ | نرخ هزینه مالی سالانه در Settings؛ محاسبه باز و قابل Audit | نرخ مالی سازمانی (WACC یا تسهیلات) |
| 4 | **استهلاک قالب/ابزار** | نحوه تسهیم | «تسهیم خطی بر تعداد قطعات پروژه» با نرخ/عمر از Settings | روش استهلاک واقعی و قالب‌های فعال |
| 5 | **ظرفیت تولید و Heijunka** | MRP واقعی در Excel محدود است | فقط «زمان ساخت بر اساس سرعت/ساعت فرضی» + تاریخ‌های برنامه‌ریزی؛ نه تخصیص ظرفیت | مراکز کاری، تقویم ظرفیت، حجم سفارشات موازی |
| 6 | **Lead Time تأمین چندمنبعی** | هر تأمین‌کننده متفاوت | Lead Time در خود ردیف Quote/Price است و به Schedule منتقل می‌شود | — |
| 7 | **بازگشت به‌عقب (Rework پس از Approval)** | Overwrite اطلاعات | Revision جدید ساخته می‌شود؛ نسخه قبلی دست‌نخورده؛ `Change_Log` ثبت می‌کند | سیاست سازمان درباره خرید/مصرف نسخه قبلی |
| 8 | **واحدهای اندازه‌گیری متنوع (kg/m/pcs)** | تبدیل واحد نیازمند جدول | فقط pcs/kg معرفی؛ تبدیل واحد خودکار پیاده‌سازی **نشده** (ASSUMPTION: یک واحد پایه) | جدول تبدیل واحد سازمانی |
| 9 | **مجوز/سقف تأیید (Threshold Approval)** | سقف امضای مدیران | سقف مبلغ تأیید در `Users`/`Approval_Matrix`؛ کنترل فقط هشدار (Excel نمی‌تواند امضا اعمال کند) | سقف‌های واقعی هیئت مدیره |
| 10 | **امنیت دسترسی واقعی** | Excel امنیت ندارد | کنترل‌های Role و قفل با VBA (اختیاری، خاموش)؛ مستندسازی محدودیت؛ معماری آینده = Permission در لایه App | سیستم SSO/AD سازمان |

> **استنتاج مستقیم از اصل ۴ (ضد جعل منطق):** در هیچ‌جای Workbook هیچ نرخ سازمانی به‌صورت ثابت داخل فرمول نوشته نشده است. تمام نرخ‌ها ستون/سلول قابل ویرایش دارند و پیش‌فرض دمو با برچسب `ASSUMPTION` علامت‌گذاری شده است.

---

## 1.4 ورودی/خروجی هر مرحله (I/O)

| مرحله | ورودی (از) | خروجی (به) | رکورد/مستند تولیدشده |
|---|---|---|---|
| Sales / Order | مشتری، مشخصات مشتری | مدیریت پروژه | Orders + Order_Lines |
| Sales Review | Orders | Engineering | وضعیت Order=Approved + پاسخ کنترل‌ها |
| Project | Orders تأیید‌شده | همه واحدها | Projects (Project_ID) |
| Engineering | Product/مشخصات فنی | Planning | Product_Revision، BOM_Header/Detail، Engineering |
| Planning | BOM Released + Qty | Procurement/Schedule | Planning (نیاز/کسری/خرید)، Schedule |
| Procurement | Planning | Costing | Supplier_Quotes، Supplier_Quote_Lines، Procurement، Purchase_Prices |
| Costing/Finance | Purchase_Prices + Procurement | Sales | Costing + Cost_Lines + Budget(Actual) |
| Sales Quotation | Costing + Budget | Management | Sales_Quotation (+نسخه‌ها) + Payment_Terms |
| Management | Quotation + ریسک/شرایط | — | Management_Approval (+ تصمیم + نسخه) |

---

## 1.5 داشبورد (خلاصه الزامات — پیاده‌سازی در شیت Dashboard)

- **KPI:** پروژه فعال، سفارش جدید، در انتظار بررسی فروش، پروژه تأخیری، BOM در انتظار تأیید، خرید در انتظار، قیمت منقضی، Costing ناقص، پیشنهاد در انتظار مدیریت / تصویب‌شده / برگشتی / ردشده، Total Sales Value، Total Cost، Gross Margin، میانگین زمان تصویب، میانگین زمان پیشنهاد.
- **Funnel:** Leads ← Orders ← Engineering ← Costing ← Quotation ← Approval ← Approved (خروجی چیزی شبیه `=COUNTIFS(Projects!H:H,"<>Cancelled")` و …؛ الزام کنترل کامل در جدولها).
- **Project Timeline:** برای هر پروژه ردیف «وضعیت هر واحد: Completed / In Progress / Pending» از شیت `Project_Statuses` — به گوی دایره‌ای با Conditional Formatting نشان داده می‌شود.
- همه KPIها از فرمول‌های `COUNTIFS/SUMIFS` روی ستون‌های کلیدی جدول‌ها محاسبه می‌شوند (نه اعداد ثابت).
