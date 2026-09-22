# 09 — User Guide (راهنمای کاربری)
نسخه: 1.0

## 9.1 شروع سریع
1. فایل `output/ExcelForBorna.xlsx` را باز کنید.
2. با `README` شروع کنید (معماری و محدودیت‌ها).
3. جریان استاندارد همان ترتیب شیت‌هاست: Orders → Projects → Engineering/BOM → Planning → Procurement/Purchase_Prices → Costing → Sales_Quotation → Management_Approval.
4. `Dashboard` وضعیت لحظه‌ای را نشان می‌دهد؛ `_Trace` مسیر قیمت را تا منبع باز می‌کند.

## 9.2 نقش‌ها و دسترسی
نقش‌ها: Sales، Engineering، Planning، Procurement، Finance، Management، Admin (شیت `Users`/`Roles`).
در اکسل کنترل دسترسی واقعی وجود ندارد؛ با VBA اختیاری و محدودیت‌های مستند در §10 مهار می‌شود. هر کاربر فقط خانه‌های مربوط به نقشش را ویرایش کند.

## 9.3 چرخه یک سفارش (نقشه عملی)
1. **Sales:** در `Order_Lines` ردیف محصول اضافه کن؛ در `Orders` کنترل‌های CO_01..04 و Ready_Flag به‌صورت خودکار پر می‌شوند. تا Ready_Flag=READY نباشد، گیت G1 (شیت Error_Checks) قرمز است.
2. **PMO:** در `Projects` ردیف پروژه با Project_ID و تیم ثبت کن؛ `Project_Statuses` خط زمانی هر واحد را نشان می‌دهد.
3. **Engineering:** در `BOM_Header` اگر تغییری لازم بود **Revision جدید** (نه ویرایش ردیف قبلی) بساز؛ `BOM_Detail` چندسطحی با Parent/Level/Sequence؛ `Status_ID=Approved` برای Release.
4. **Planning:** `Planning` خودکار BOM را برای مقدار سفارش مقداردهی می‌کند (Required/Buy_Qty)؛ `Schedule` فعالیت‌ها را با تأخیر محاسبه‌شده نگه می‌دارد.
5. **Procurement:** در `Supplier_Quotes`/`Supplier_Quote_Lines` چند تأمین‌کننده ثبت کن؛ انتخاب با مسئول مجاز (ستون Selected). `Purchase_Prices` با Revision و Landed (تبدیل ارز شفاف). قیمت بدون منبع/تاریخ/ارز/تأمین‌کننده در Error_Checks قرمز می‌شود.
6. **Finance:** `Cost_Lines` ریز اجزا؛ `Costing` جمع می‌کند؛ `Budget` بودجه در برابر عملکرد.
7. **Sales:** `Sales_Quotation` سناریو A/B/C (حتماً Basis_Memo پر شود)؛ `Payment_Terms`.
8. **Management:** `Management_Approval` با یکی از چهار تصمیم.

## 9.4 فرمول‌های کلیدی (کجا می‌بیند)
- Planning.Required = `BOM.Qty × (1+Scrap) × Order_Qty` (زنده از Order_Lines).
- Purchase_Prices.Landed = `Σ(اجزا) × Fx`.
- Costing.Total = `SUMIFS(Cost_Lines…)` بر اساس Category.
- Sales_Quotation.Proposed = `Cost + Margin + Commercial + Risk + Financial`.
- Budget: `Variance = Actual − Budget`.
- Dashboard: همه KPI با COUNTIFS/SUMIFS زنده.

## 9.5 کنترل نسخه
هر موجودیت تأییدشده با `Revision` پیش می‌رود؛ ردیف قبلی دست‌نخورده می‌ماند (مثال BOM-000215-R04 و R05). تغییرات در `Change_Log` و انتقال وضعیت در `Workflow_History` ثبت می‌شود.

## 9.6 چاپ/PDF
شیت `Report_Quotation` گزارش اجرایی یک‌صفحه‌ای است (قابل چاپ/PDF) با XLOOKUP زنده به داده‌ها.

## 9.7 افزایش سطرها
- زیر آخرین ردیف هر جدول، ردیف جدید بده (جدول Excel خودش گسترش می‌یابد و Structured References دنبال می‌شوند).
- برای شمارنده‌های ID: از ستون شمارنده در `Settings` (دمو) استفاده یا System_ID را دستی وارد کن؛ فرمول‌های Duplicate Check محدوده A3:A1000 دارند (برای سطر بیشتر، بازه‌ها را گسترش بده).
