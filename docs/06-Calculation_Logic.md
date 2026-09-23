# 06 — منطق محاسباتی (Calculation Logic) — شفاف و قابل Audit
نسخه: 1.0

## 6.0 اصول
- همه فرمول‌ها ستونی و مبتنی بر Structured Reference هستند؛ هیچ عددی داخل فرمول Hard-code نیست (نرخ‌ها از `Settings`/`Lists` خوانده می‌شوند).
- تمام نرخ‌های سازمانی که در دمو عدد دارند، در `Settings` با برچسب `ASSUMPTION` آمده‌اند.
- هر فیلد محاسباتی از ستون «مرجع» (Ref) خود اعلام می‌کند از کجا می‌آید.

## 6.1 برنامه‌ریزی (Planning)
- `Required_Qty (ماده) = BOM_Line.Gross_Qty × Total_Order_Qty(مرتبط با پروژه) × (1 + Scrap%)`
  و برای BOM چندسطحی:
  `Gross_Qty_line = Qty_Per_line × Gate_Qty_Level_Below × (1 + Scrap_line)`
  به‌گونه‌ای که «مقدار مصرف در سطح بعد × مقدار والد موردنیاز» (توضیح: مقدار هر برگ، ضرب زنجیره‌ای مقادیر از ریشه تا آن برگ است؛ دمو فقط یک سطح Assembly + Material دارد بنابراین زنجیره کوتاه است).
- `Deficit = max(0, Required_Qty - (Stock_On_Hand - Reserved))`
- `Buy_Qty = max(0, Deficit) گرد‌شده به مضرب MOQ در صورت لزوم` (در Procurement اعمال دقیق می‌شود؛ Planning مقدار خام را می‌دهد).
- `Supply_Date = امروز + Lead_Time_Days` (کاری).

## 6.2 زمان‌بندی (Schedule)
- `Duration = max(0, End_Date - Start_Date) + 1` (روز کاری)
- `Delay_Days = max(0, امروز - Planned_End) در صورت عدم Done` (یا اختلاف واقعی/برنامه)
- وابستگی‌ها: ستون `Dependency` متن Activity قبلی است؛ تاریخ شروع پیشنهادی مرحله بعد = WORKDAY(پایان قبلی) در صورت ثبت دستی.

## 6.3 قیمت خرید (Price Build-up → Landed)
برای هر واحد (یا کل ردیف):
```
Unit_Price
+ Freight_Per_Unit
+ Insurance_Per_Unit
+ Customs_Per_Unit
+ Handling_Per_Unit
+ Other_Per_Unit
──────────────
= Landed_Cost_Unit
```
قانون الزامی اعتبار ورود به محاسبات (طبق الزام):
`قیمت معتبر = AND(Price_Source≠"", Price_Date موجود, Supplier_ID موجود, Currency مشخص)`
دو وضعیت بد کنترل می‌شود:
- `Expired_Flag = IF(TODAY()>Valid_To, 1, 0)` — قیمت منقضی از Costing حذف می‌شود.
- `Stale_Flag = IF(Price_Date < امروز-90, 1, 0)` — قیمت قدیمی هشدار (استفاده از قیمت خرید قدیمی، الزام بخش 20).

## 6.4 قیمت تمام‌شده (Costing)
اجزاء:
```
Direct_Material     = Σ Landed_Cost اقلام Raw/Purchased طبق BOM برای مقدار سفارش
Purchased_Parts     = Σ Landed_Cost قطعات خریدنی (Make/Buy = Buy)
Direct_Labor        = Labor_Hours × Labor_Rate          (Rate از Settings — ASSUMPTION)
Manufacturing_Overhead = Direct_Labor × Overhead_Rate   (Rate از Settings — ASSUMPTION)
Scrap               = Direct_Material × (Scrap_Factor)  (Factor از Settings)
Packaging           = unit × qty (از Engineering یا Settings)
Logistics           = حمل/بیمه (از Procurement یا Settings)
Tooling             = Tooling_Cost / Order_Qty (استهلاک خطی) (ASSUMPTION روش)
Other_Allocated     = از Settings یا مستقیم
Financial_Cost      = هزینه مالی اعتبار فروش (نرخ سالانه × مدت/360 × مبلغ) (نرخ از Settings)
SGA                 = (Direct_Material+Labor+O/H) × SGA_Rate در صورت «سیاست سازمان = بله»
Total_Cost          = SUM(اجزای فوق)
Cost_Per_Unit       = Total_Cost / Order_Qty
```
هر جزء یک ردیف `Cost_Line` دارد تا Drill-down ممکن باشد: `Costing.Total_Cost = SUM(Cost_Lines.Amount)`.

## 6.5 بودجه
```
Total_Budget = Budget_Material + Budget_Labor + Budget_Overhead + Budget_Procurement
             + Budget_Logistics + Budget_Tooling + Budget_Other
Actual = قرائت از Costing موردنظر (همان اجزاء)
Variance = Actual - Budget (هر جزء)
Budget_vs_Actual = Actual / Budget - 1
```

## 6.6 پیشنهاد قیمت فروش
```
Total_Cost
+ Desired_Margin% × Total_Cost            (Margin مفهوم % روی Cost)
+ Commercial_Adjustment
+ Financial_Adjustment
+ Risk_Adjustment
═══════════════
= Proposed_Selling_Price
```
- `Margin_Amount = Total_Cost × Desired_Margin%`
- `Margin_% گزارش = (Selling_Price - Total_Cost) / Selling_Price`
- سناریو A/B/C هرکدام ردیف جدا دارند؛ **قانون: هیچ سناریویی بدون `Basis_Memo` (مبنای محاسبه) وارد وضعیت Submitted نمی‌شود.**

## 6.7 هزینه مالی شرایط پرداخت
```
Fin_Cost_Amount = (Credit_Portion × مبلغ) × (Fin_Cost_Rate_Annual × Credit_Days/360)
```
نرخ سالانه از Settings است؛ اگر Credit_Days=0 → 0.

## 6.8 Trace
به‌ازای هر Quotation انتخابی، شیت `_Trace` زنجیره «Selling Price ← Costing ← Cost_Line ← Material ← BOM ← Supplier Quote ← Supplier» را با فرمول XLOOKUP می‌سازد.

## 6.9 کنترل گرد کردن
گرد کردن فقط برای نمایش (فرمت عددی) و در مرحله نهایی «قیمت پیشنهادی» با ROUND انجام می‌شود؛ جمع‌های میانی گرد نمی‌شوند تا Audit دقیق بماند.

## 6.10 فرمول‌های نسخه 2.0

**نفرساعت (Time_Study):**
```
Order_Qty        = SUMIFS(Order_Lines.Qty; Order; XLOOKUP(Project→Order); Product; Product)
Man_Min_Total    = Setup_Min + Std_Min_Per_Unit × Order_Qty × (1 + Scrap_Allowance)
Man_Hours_Total  = ROUND(Man_Min_Total / 60 × Operators, 2)
```

**دستمزد در قیمت تمام‌شده (زنده از زمان‌سنجی):**
```
Cost_Lines[Direct Labor].Qty   = SUMIFS(Time_Study.Man_Hours_Total; Project)
Cost_Lines[Direct Labor].Amount= ROUND(Qty × Settings.Labor_Rate_Per_Hour)
```

**تجهیزات وارداتی:**
```
Fx_Rate          = IF(Currency="IRR", 1, Settings.Fx_USD_TO_IRR)
Unit_Price_IRR   = ROUND(Unit_Price_Quote × Fx_Rate)
Total_Cost_IRR   = Unit_Price_IRR × Qty + Transport + Installation
```

**امکان‌سنجی:**
```
Weighted = Σ وزن_بُعد × امتیاز_بُعد   (وزن‌ها و کف قبولی از Settings)
Result   = Feasible اگر Weighted ≥ کف و ریسک بالا نباشد؛ وگرنه Conditional / Not Feasible
```

**گیت‌های تجمعی (نمونه):**
```
G_Eng = S1 × S2            (ورود مهندسی)
G_Com = G_Eng × S3         (ورود بازرگانی)
… و به همین ترتیب تا G_Senior = G_PMO × S8
```
این گیت‌ها مستقیماً در Data Validation سفارشی سلول‌های ورودی استفاده می‌شوند
(مسیر: `INDEX(Gate_Status!ستون_گیت; MATCH(پروژه))=1`)، بنابراین ورود اطلاعات
مرحله بعد تا تکمیل مرحله قبل توسط خود اکسل رد می‌شود.
