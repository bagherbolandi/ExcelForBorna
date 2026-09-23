# 07 — کدهای خطا، کنترل‌ها و برنامه تست
نسخه: 1.0 (منطبق بر الزامات 20 و 33)

## 7.1 Error Codes (هر خطا Error Code / Description / Owner / Required Action / Status)

| Error Code | Error Description | Owner | Required Action | نمونٔه وضعیت / Rule-Key |
|---|---|---|---|---|
| E-001 | اطلاعات ناقص سفارش (Customer/Product/Qty/تاریخ/فروشنده) | Sales | تکمیل و ارسال مجدد | / CO-01..07 |
| E-002 | BOM بدون Revision | Engineering | صدور Revision و تأیید | / G2 |
| E-003 | قیمت خرید بدون تاریخ | Procurement | ثبت Price_Date | / PX-01 |
| E-004 | قیمت بدون منبع | Procurement | ثبت Price_Source | / PX-02 |
| E-005 | تأمین‌کننده بدون تأیید | Procurement | تکمیل Approval_Status | / SP-01 |
| E-006 | Quantity صفر | Sales | اصلاح مقدار | / CO-04 |
| E-007 | قیمت منفی | Procurement/Finance | اصلاح قیمت | / PX-03 |
| E-008 | Margin غیرمنطقی (منفی یا > سقف) | Sales | بازبینی سناریو | / SQ-02 |
| E-009 | ارز نامشخص | Procurement | انتخاب ارز | / PX-04 |
| E-010 | قیمت منقضی | Procurement | استعلام/تمدید قیمت | / PX-05 |
| E-011 | تأخیر در فعالیت | PM | بازبینی Schedule | / SC-01 |
| E-012 | اختلاف BOM و سفارش | Engineering | هماهنگی BOM/Spec | / G2-extra |
| E-013 | اختلاف بودجه و قیمت تمام‌شده | Finance | بازبینی Costing/Budget | / BG-01 |
| E-014 | عدم تأیید واحد قبلی | Owner فعلی | توقف و ارجاع | / Gate |
| E-015 | تغییر اطلاعات بعد از Approval | Owner + Approver | Revision جدید | / RV-01 |
| E-016 | نسخه قدیمی BOM | Engineering | فعال‌سازی Revision جدید | / BM-01 |
| E-017 | استفاده از قیمت خرید قدیمی (Stale) | Procurement | استعلام مجدد | / PX-06 |
| E-018 | Duplicate Order_ID | Sales | بررسی و اصلاح | / DK-01 |
| E-019 | Duplicate BOM_ID+Revision | Engineering | صدور Revision جدید | / DK-02 |
| E-020 | Duplicate Supplier_Quote_ID+Version | Procurement | نسخه جدید | / DK-03 |
| E-021 | Costing ناقص (اجزای صفر/کاربردی) | Finance | تکمیل ردیف‌ها | / CS-01 |

## 7.2 جدول Error_Checks (نمونه ردیف‌ها)
هر ردیف: `Rule_Key, Check_Description, Severity, Owner, Formula_Result, Message, Required_Action, Status`.
نمونه‌ها:
- DK-01: `=IF(COUNTIF(Orders[Order_ID], Orders[@Order_ID])>1,"ERROR","OK")`
- G2: `=IF(BOM_Status(BOM_Header[@BOM_Revision_ID])="Approved","UNLOCK","LOCK")`
- PX-05: `=IF([@Valid_To]<TODAY(),"EXPIRED","VALID")`

## 7.3 Test Plan (15 مورد الزامی بخش 33)

| # | Scenario | Input | Expected |
|---|---|---|---|
| T1 | سفارش کامل | Order با همه فیلدها | Ready=TRUE، عبور تا Management |
| T2 | اطلاعات ناقص | Order بدون Qty | Ready=FALSE، Issue E-006، قفل G1 |
| T3 | BOM Revision جدید | تغییر خط BOM پس از Release | BMR جدید، نسخه قبلی دست‌نخورده، Change_Log |
| T4 | تغییر قیمت تأمین‌کننده | قیمت جدید در Quote_Line | Landed/Refreshed، Costing به‌روز، Revision |
| T5 | تغییر مقدار سفارش | Qty 100→200 | Planning/نیاز دوبرابر، Costing، Quotation، Revision |
| T6 | تغییر شرایط پرداخت | Credit 30→60 | Fin_Cost_Amount تغییر، Quotation به‌روز |
| T7 | عودت مدیریت | Decision=RETURN FOR REVISION | Quotation Returned، Issue ثبت، پروژه به گام |
| T8 | رد پیشنهاد | Decision=REJECTED | وضعیت Rejected، Closed |
| T9 | تصویب پیشنهاد | Decision=APPROVED | APPROVED + Workflow_History |
| T10 | تغییر بعد از Approval | تغییر Qty پس از APPROVED | Revision جدید، Approval قبلی حفظ، RV-01 |
| T11 | قیمت منقضی | Valid_To در گذشته | Expired_Flag=1، از Costing کنار گذاشته می‌شود |
| T12 | تأخیر تأمین‌کننده | Arrival بعد از برنامه | Schedule.Delay>0، Issue SC-01 |
| T13 | اختلاف بودجه/هزینه | Actual≠Budget | Variance نشان داده، Issue BG-01 |
| T14 | Duplicate Order | Order_ID تکراری | Error DK-01، مجاز نیست |
| T15 | تغییر BOM بعد از Costing | BOM Rev جدید | Costing مجدد لازم، نسخه قبلی حفظ |

## 7.4 سناریوی دمو End-to-End (الزام 32)
یک پروژه کامل `PRJ-2026-00125` «پمپ گریز از مرکز سرویس سنگین» شامل: Customer CUS-000001، Order ORD-2026-00321 (۲ خط سفارش)، پروژه، مهندسی + BOM چندسطحی (Assembly + Raw/Purchased/Packaging)، Planning، Schedule، Quotes از 3 تأمین‌کننده برای مواد کلیدی، انتخاب یک مورد، Purchase_Prices با Price Build-up، Costing، Budget، Sales_Quotation سناریو A/B/C، Payment_Terms، Management_Approval.

## 7.5 نسخه 2.0 — کدهای خطای گیتینگ و تست‌های جدید

| کد | شرح | شدت | مالک |
|---|---|---|---|
| E-022 | امکان‌سنجی ناقص (قفل مهندسی) | Block | PMO |
| E-023 | تجهیزات قیمت‌دهی‌نشده | Warning | بازرگانی |
| E-024 | رسید انبار بدون ردیف برنامه‌ریزی معتبر | Error | برنامه‌ریزی |
| E-025 | جمع‌بندی گزارش‌نشده به ارشد | Warning | PMO |
| E-026 | زمان‌سنجی ثبت‌نشده | Block | مهندسی |

تست‌های افزوده: **T16..T20** (گیت امکان‌سنجی، گیت بازرگانی، زنجیره نفرساعت،
نرخ ارز اقلام وارداتی، گیت جمع‌بندی/تصمیم) — جزئیات و نتایج در سند 08.
