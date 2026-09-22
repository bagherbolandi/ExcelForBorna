# 10 — محدودیت‌های Excel و معماری آینده (Web / Database)
نسخه: 1.0

## 10.1 محدودیت‌های صریح این نسخه (نه پنهان‌کاری)
| # | محدودیت | چرا در Excel همین است | راه‌حل آینده |
|---|---|---|---|
| 1 | **امنیت دسترسی واقعی وجود ندارد** | اکسل قفل واقعی برنامه‌ای ندارد | Role/Permission در لایه App + SSO |
| 2 | **یک‌کاربره عملی** | هم‌زمانی (Concurrency) ایمن نیست | DB تراکنشی + Optimistic Lock (Version Number) |
| 3 | **Audit تراکنشی** | Workflow_History ورودی دستی/نیمه‌خودکار است | Trigger/Event-Sourcing خودکار در هر UPDATE |
| 4 | **آرایه‌های سنگین** (XLOOKUP دوشرطی در Timeline) | با موتور سبک ما تست‌پذیر نیست؛ در اکسل واقعی درست کار می‌کند | SQL JOIN |
| 5 | **تبدیل واحد** (kg↔pcs) پیاده نشده | نیازمند جدول تبدیل سازمانی (ASSUMPTION) | جدول Unit_Conversion + محاسبه سمت سرور |
| 6 | **MrP/ظرفیت واقعی** | فقط زمان‌بندی تاریخ‌محور؛ بدون تخصیص مرکز کاری | Work Center + تقویم شیفت + الگوریتم زمان‌بندی |
| 7 | **تقویم تعطیلات/شیفت** | WORKDAY ساده بدون تعطیلات رسمی | جدول Holiday + Working Calendar |
| 8 | **مالیات/کسور قانونی** | بدون منطق مالیاتی (ASSUMPTION) | ماژول مالیاتی بر اساس قوانین |
| 9 | **چندارزی زنده** | نرخ دستی در Settings؛ بدون تاریخ نرخ | جدول FxRate با تاریخ + لود روزانه |
| 10 | **سقف امضا (Threshold)** | فقط هشدار | قانون Approval با مبلغ + گردش الکترونیک |
| 11 | **سطرهای بزرگ/کارایی** | برای صدها هزار ردیف مناسب نیست | RDBMS + ایندکس |
| 12 | **نسخه‌بندی اتوماتیک** | Revision با انضباط کاربری/VBA | Version Table + فقط Insert مجاز |
| 13 | **روابط FK** | روابط با Data Validation ‘نمایشی’ است نه سخت | FOREIGN KEY در DB |

## 10.2 چرا این نُقض «جعل منطق» نیست
هر عدد/نرخ/فرمول سازمانی در Settings ثبت شده و برچسب ASSUMPTION خورده؛ هیچ منطق سازمانی حدس زده نشده است (§34).

## 10.3 معماری آینده (تبدیل به Web/Database)
**پیشنهاد:** ماهیت داده‌ای این Workbook مستقیماً قابل انتقال است — جدول‌های همین فایل همان جداول پایگاه‌داده‌اند.

```
Frontend (React/Vue)  →  REST/GraphQL API  →  Service Layer (استفاده‌موردی)
                                              ├─ Order Service
                                              ├─ Engineering/BOM Service
                                              ├─ Procurement/Quote Service
                                              ├─ Costing Service
                                              ├─ Quotation Service
                                              └─ Workflow/Approval Engine (BPM: Camunda/Temporal)
                                              →  RDBMS (PostgreSQL): جداول = tbl* همین سامانه
                                              →  Object Storage برای پیوست‌ها
```

### نگاشت مستقیم (1:1 با Workbook)
- هر شیت Master/Transaction = یک جدول SQL با همان نام/ستون.
- `Lists` = جداول Enum/Nomenclature؛ `Settings` = جدول Configuration.
- `Workflow_History` = تراکنش‌های Workflow Engine (بدون Force-Push تاریخچه).
- `Error_Checks` = قوانین Validation در لایه خدمت + قوانین BPM.
- `Change_Log` = Event-Sourcing / Audit Log دیتابیسی.
- `_Trace` = API گزارش Drill-down (Price → Cost → BOM → Supplier).
- `Dashboard` = جداول Aggregate + BI (Metabase/Power BI).

### گام‌های مهاجرت
1. Schema Migration از Data Dictionary (سند 05) با همان کلیدها.
2. سرویس CRUD برای Master Data (مشتری/محصول/ماده/تأمین‌کننده/کاربر).
3. Workflow Engine برای زنجیره وضعیت + گیت‌ها (هر گیت در سند 02 همانجا پیاده می‌شود).
4. موتور محاسبه (Planning/Costing/Quotation) به Service Layer منتقل شود — فرمول‌ها از قبل در سند 06 مشخص هستند.
5. Role/Permission (RACI و Users) جابه‌جا شود.
6. گزارشات و Dashboard در BI.

**نتیجه:** این فایل نه‌تنها «فرم» بلکه **Blue-Print قابل اجرا** برای نرم‌افزار سازمانی است.
