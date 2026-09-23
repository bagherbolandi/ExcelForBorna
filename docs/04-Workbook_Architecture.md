# 04 — معماری Workbook (ساختار فایل Excel)
نسخه: 1.0

## 4.1 تصمیم‌های کلیدی معماری
1. **دو ناحیه در هر شیت داده‌ای:** بالای هر جدول، یک «نوار کنترل» کوتاه (زمینه خاکستری) شامل Rule-Key، اعتبارسنجی‌ها و (در شیت‌های خروجی) پارامتر نما قرار می‌گیرد. خود داده از سطر ۳ به پایین **فقط جدول** است.
2. **Table + Structured References** برای همه داده‌ها: فرمول‌ها می‌توانند EFS (`=Sales_Quotation[Netting_Price]`) باشند و دامنه فرمول به‌صورت ساخت یافته است. در نوشتن فایل، از `Table1[Column]` استفاده می‌شود.
3. **چند سناریو/نسخه با جدول‌های «Header/Line» عادی** — نه جدول‌های سه‌بعدی، نه سلول‌های ادغام‌شده.
4. **نرخ‌ها و ضرایب در Settings** (بدون Hard-code).
5. **ستون‌های قفل برگشت:** در سلول‌های وضعیت، فرمول از «ورودی سمت چپ» عبور نمی‌کند مگر پیش‌نیاز سبز باشد: `=IF(AND( ...checks... ), "UNLOCK", "LOCK")`.

## 4.2 جدول کامل شیت‌ها (نام شیت + نام جدول + نوع + نقش)

### Reference / Master
| # | Sheet | Table | نوع | محتوا/نقش |
|---|---|---|---|---|
| 1 | README | — | سند | راهنما + نسخه + محدودیت‌ها |
| 2 | Workflow | (قالب) | سند | نقشه گردش کار + گام‌ها + گیت‌ها |
| 3 | Lists | Lists | Lookup | همه لیست‌های موردنیاز (واحد، ارز، Make/Buy، انواع ماده، …) |
| 4 | Settings | tblSettings | Config | نرخ‌ها/ضرایب/پیش‌فرض‌ها (ASSUMPTION) + شمارنده‌های Seq |
| 5 | Statuses | tblStatuses | Master | Status_ID/Name/Group/Sequence |
| 6 | Departments | tblDepartments | Master | Department_ID/Name |
| 7 | Users | tblUsers | Master | User_ID/Name/Dept/Role/Email/Approval_Level/Active |
| 8 | Roles | tblRoles | Master | Role_ID/Role_Name |
| 9 | Approval_Matrix | tblApprovalMatrix | Config | Entity + Gate + Approver Role + Min Level |
| 10 | RACI | tblRACI | Config | Entity/Action + R/A/C/I |
| 11 | Customers | tblCustomers | Master | Customer_ID + مشخصات |
| 12 | Products | tblProducts | Master | Product_ID/Product_Code/Name/Type/UoM |
| 13 | Product_Revisions | tblProductRevisions | Master-Ver | PRV + محصول + Drawing/Revision |
| 14 | Materials | tblMaterials | Master | Material_ID/Code/Description/Type/Spec |
| 15 | Suppliers | tblSuppliers | Master | Supplier_ID/Code/Name/Status |

### Transaction
| # | Sheet | Table | نقش |
|---|---|---|---|
| 16 | Orders | tblOrders | سفارش + کنترل‌های اولیه (CO-01..07, Req-01..07) + Status |
| 17 | Order_Lines | tblOrderLines | OLN + Order_ID + Product_ID + Qty + واحد |
| 18 | Projects | tblProjects | Project_ID + Order + PM + تاریخ‌ها + وضعیت + اولویت + تیم + % پیشرفت |
| 19 | Project_Statuses | tblProjectStatuses | وضعیت هر واحد برای هر پروژه (خط Timeline) |
| 20 | Engineering | tblEngineering | اطلاعات فنی (Drawing/Rev، BOM/Rev، یادداشت) |
| 21 | BOM_Header | tblBOMHeader | نسخه‌های BOM (BOM_ID + Revision + Product_Revision) |
| 22 | BOM_Detail | tblBOMDetail | خطوط چندسطحی BOM (Parent, Level, Qty, Scrap, Make/Buy) |
| 23 | Supplier_Quotes | tblSupplierQuotes | هدر استعلام به تفکیک ماده |
| 24 | Supplier_Quote_Lines | tblSupplierQuoteLines | یک ردیف به ازای هر تأمین‌کننده (قیمت/شرایط/امتیازها) |
| 25 | Procurement | tblProcurement | تصمیم خرید (از Quote انتخاب‌شده یا مستقیم) + لجستیک |
| 26 | Purchase_Prices | tblPurchasePrices | قیمت خرید تأییدشده + Price Build-up + Landed |
| 27 | Planning | tblPlanning | مقداردهی BOM بر حسب نیاز (Required/Gross/Stock/Deficit/Buy) |
| 28 | Schedule | tblSchedule | زمان‌بندی فعالیت‌ها (وابستگی/تأخیر) |
| 29 | Budget | tblBudget | بودجه ابتدایی + Budget vs Actual |
| 30 | Costing | tblCosting | اجرای Costing (جمع اجزا) |
| 31 | Cost_Lines | tblCostLines | ریز اجزای هزینه (Drill-down) |
| 32 | Sales_Quotation | tblSalesQuotation | سناریو A/B/C + 준بنا |
| 33 | Quotation_Versions | tblQuotationVersions | نسخه‌های پیشنهاد |
| 34 | Payment_Terms | tblPaymentTerms | شرایط پرداخت + هزینه مالی |
| 35 | Management_Approval | tblManagementApproval | تصمیم‌ها |

### Audit / Control
| # | Sheet | Table | نقش |
|---|---|---|---|
| 36 | Workflow_History | tblWorkflowHistory | لاگ انتقال وضعیت |
| 37 | Change_Log | tblChangeLog | لاگ تغییرات |
| 38 | Issues | tblIssues | Register ریسک/مشکل |
| 39 | Error_Checks | tblErrorChecks | خطاهای فعال سیستم |
| 40 | Test_Results | tblTestResults | خروجی Test Cases (الزام 33) |

### Presentation
| # | Sheet | نقش |
|---|---|---|
| 41 | Dashboard | KPI + Funnel + Timeline |
| 42 | _Trace | Drill-down قیمت تا منبع |
| 43 | Report_Quotation | گزارش اجرایی قابل چاپ/PDF |
| 44 | _NamedRanges | فهرست Named Ranges (مستندسازی) |

## 4.3 طرح مشترک شیت‌های داده (Layout Contract)
- سطر 1: عنوان شیت.
- سطر 2: نوار کنترل (فقط برای شیت‌های دارای فرمول‌هایش به‌صورت نتایج پس از سطر 3).
- سطر 3: سرستون‌های جدول، سطر 4 به بعد: داده.
- استفاده از اعداد اعشاری با ۲ رقم؛ درصدها با `0.0%`؛ تاریخ‌ها با `yyyy-mm-dd`.

## 4.4 قرارداد شماره‌ستون برای Accessors
در برنامه نویسنده، مختصات ردیف 3 (هدر) به‌عنوان «نقشه ستون» استفاده می‌شود؛ فرمول‌ها با `Table[Col]` نوشته می‌شوند، اما مقادیر در مرحله ارزیابی موتور با A1 (ردیف دقیق) مطابقت داده می‌شوند. این قرارداد در `src/builder/build.py` مستند است.

## 4.5 جداول «موقت» برای Demo و Test
داد‌ه‌های دمو فقط در جداول معمولی قرار دارند (بدون شیت پنهان اضافه). هر شیت تست با یک کادر شکل/Rich در صورت نیاز به `README` اشاره می‌کند. استخراج CSV در تست‌ها بدون محاسبه مجدد (مقدارهای نوشته‌شده) انجام می‌شود.

## 4.6 معماری نسخه 2.0 (53 شیت)

- **ترتیب زبانه‌ها** بر اساس گردش کار چیده شده است: داشبورد/وردک‌فلو/گیت‌ها در ابتدا، سپس شیت‌های مراحل ۱ تا ۹ و در انتها شیت‌های مرجع/حاکمیتی.
- **لایه رندر جدید (`renderer_v2.py`)** چهار مسئولیت دارد: (1) راست‌به‌چپ و رنگ زبانه‌ها، (2) گیت‌های Data Validation سفارشی با ارجاع به `Gate_Status`، (3) محافظت همه شیت‌ها با گذرواژه راهبر + تزریق `<protectedRanges>` رمزدار پس از ذخیره، (4) فرمت شرطی گانت و نمودارهای داشبورد.
- **چرخه‌های سلولی مجاز بین شیتی** (مثل `Gate_Status ↔ PMO_Summary`) فقط در سطح سلول غیردوری‌اند؛ موتور ارزیابی (`verify.py`) برای بی‌نیازی از ترتیب شیت‌ها به الگوریتم **نقطه‌ثابت (ژاکوبی)** ارتقا یافت.
- ظرفیت پیش‌فرض محدوده‌ها برای ورود داده‌های آینده: تا ردیف ۲۰/۳۰/۶۰ بسته به شیت (فراتر از آن، محدوده‌ها را در `renderer_v2.EDIT_RANGES` و بازه فرمول‌ها گسترش دهید).
