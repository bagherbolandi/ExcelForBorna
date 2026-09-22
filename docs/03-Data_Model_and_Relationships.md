# 03 — مدل داده، شناسه‌ها و روابط جداول (Data Dictionary)
نسخه: 1.0

## 3.0 اصول
- هر جدول یک **Excel Table** است (Structured References). ستون کلید واقع در سمت چپ.
- هر جدول User-Owned (جداول ثبت عملیات) شامل ۸ ستون حسابرسی است:
  `System_ID, Business_Code, Revision, Status_ID, Create_Date, Created_By, Modify_Date, Modified_By`
- شناسه‌های System ID در ستون جدا از Business Code و با قالب `<PREFIX>-<seq6>` (خودکار با فرمول) تولید می‌شوند.
- مقادیر کلید با **Data Validation از لیست رسمی** فهرست وارسی می‌شود.
- هیچ داده‌ای در فرمول Hard-code نیست؛ نرخ‌ها و لیست‌ها در Settings/Lists.

## 3.1 شناسه‌ها (ID Standards)

| Entity | Prefix | نمونه نمایشی | تولید |
|---|---|---|---|
| Customer | CUS | CUS-000001 | Settings Seq + فرمول |
| Project | PRJ | PRJ-2026-00125 | سال + Seq |
| Order | ORD | ORD-2026-00321 | سال + Seq |
| Order Line | OLN | OLN-000456 | Seq سراسری |
| Product | PRD | PRD-000125 | Seq |
| Product Revision | PRV | PRV-000125-R03 | PRD + R+seq |
| BOM | BOM | BOM-000215 | Seq |
| BOM Revision | BMR | BOM-000215-R05 | BOM + R+seq |
| BOM Line | BLI | BLI-000987 | Seq |
| Material | MAT | MAT-001542 | Seq |
| Supplier | SUP | SUP-000089 | Seq |
| Supplier Quote | SQT | SQT-2026-00545 | سال + Seq |
| Supplier Quote Line | SQL | SQL-002354 | Seq |
| Purchase Price | PPR | PPR-000554 | Seq |
| Planning | PLN | PLN-000325 | Seq |
| Schedule | SCH | SCH-000256 | Seq |
| Costing | CST | CST-000321 | Seq |
| Cost Line | CLI | CLI-005654 | Seq |
| Budget | BDG | BDG-000245 | Seq |
| Quotation | QUO | QUO-2026-00125 | سال + Seq |
| Quotation Version | QUV | QUO-00125-V03 | QUO + V+seq |
| Approval | APR | APR-000547 | Seq |
| Change Log | CHG | CHG-001245 | Seq |
| Issue | ISS | ISS-000215 | Seq |
| Workflow History | WFH | WFH-000001 | Seq |
| User | USR | USR-01 | Seq |
| Status | ST | ST-01 | دستی/مقداردهی |

**قانون Business_Code:** برای موجودیت‌های دارای سال، `Business_Code` کد تجاری (مثل ORD-2026-00321) و `System_ID` همان کد است. برای نسخه‌ها، `System_ID` کد + `-R##` / `-V##` است.

## 3.2 روابط (ER)

```
Customer 1──n Project
Project 1──n Order
Order 1──n Order_Line
Order_Line n──1 Product
Product 1──n Product_Revision (هر Revision یک رکورد جدید، WHOLE نسخه)
Product_Revision 1──n BOM_Header (BOM_Revision_ID)
BOM_Header 1──n BOM_Detail (با Parent_BOM_Line_ID برای چندسطحی)
BOM_Detail n──1 Material
Material 1──n Purchase_Price
Material 1──n Supplier_Quote_Line
Supplier 1──n Supplier_Quote
Supplier_Quote 1──n Supplier_Quote_Line
Project 1──n Planning
Planning 1──n Schedule (یا Project 1──n Schedule مستقل)
Project 1──n Costing
Costing 1──n Cost_Line
Project 1──n Budget
Project 1──n Sales_Quotation
Quotation 1──n Quotation_Version
Quotation_Version 1──n Management_Approval
Project 1──n Change_Log
Project 1──n Issues
(هر Entity_Type در Workflow_History توسط Workflow_History_ID ردیابی می‌شود)
```

آدرس‌دهی کلید خارجی (FK) در ستون‌های جدول‌ها با Data Validation به فهرست شناسه‌های جدول مادر وارسی می‌شود.

## 3.3 مسیر ردیابی قیمت (Trace)

```
Quotation_Version.Netting_Price
 └─ Sales_Quotation (سناریوی انتخابی)
     └─ Costing.Total_Cost  (Costing_Revision_ID)
         └─ Cost_Line (Direct Material / Purchased Parts / Labour / O/H / Scrap / Pack / Freight / Tooling / Financial / SG&A)
             ├─ Material/BOM_Line × Purchase_Price.Landed_Cost
             ├─ Purchase_Price ← Supplier_Quote_Line.Selected
             └─ Supplier_Quote ← Supplier
```

در شیت `_Trace`، برای یک Quotation مشخص هر سطح Price→Cost→BOM→Material→Supplier Quote به‌صورت سلسله‌مراتبی باز می‌شود (Drill-down دوطرفه).

## 3.4 ساختار BOM چندسطحی

هر ردیف `BOM_Line_ID` با `Parent_BOM_Line_ID` (خالی برای ریشه) + `Level_No` + `Sequence_No` درخت می‌سازد:
- سطح 0: کالای نهایی (Header) — در Detail با Parent خالی
- سطح 1: Assembly
- سطح 2: Sub Assembly
- سطح 3: Material / Purchased Part / Packaging

قاعده عمومی کمیت: برای عمق > ۰، `Qty_Per = مقدار مصرف در سطح والد × مقدار والد در بالاتر` — ما برای دمو از منطق «BOM Quantity در هر سطح به ازای ۱ واحد والد» استفاده می‌کنیم و مقدار ناخالص در Planning به‌صورت **ضرب زنجیره‌ای از ریشه به برگ** محاسبه می‌شود (توضیح فرمول در سند 06).

## 3.5 جدول Universal Status

| Status_ID | Status_Name | Status_Group | Sequence |
|---|---|---|---|
| ST-01 | Draft | Draft | 1 |
| ST-02 | In Review | Review | 2 |
| ST-03 | Waiting Approval | Approval | 3 |
| ST-04 | Approved | Approved | 4 |
| ST-05 | Rejected | Rejected | 5 |
| ST-06 | Returned | Rejected | 6 |
| ST-07 | Closed | Closed | 7 |
| ST-08 | Cancelled | Closed | 8 |

(برای Order وضعیت مشتق `Revised`، برای Engineering `In Progress/Submitted`، برای Approval نتایج `APPROVED/WITH CONDITION/REVISION/REJECTED` از طریق Status_ID نگاشت می‌شود؛ Status_Group متناظر در `Status_Map` تعریف شده است.)

## 3.6 جدول Universal Users

| User_ID | Name | Department_ID | Role_ID | Email | Approval_Level | Is_Active |
|---|---|---|---|---|---|---|
| USR-01 | بابک (مدیر فروش) | Sales | Sales | … | 2 | 1 |
| … | (و برای هر نقش) | | | | | |

## 3.7 Workflow History

| Workflow_History_ID | Project_ID | Entity_Type | Entity_ID | From_Status | To_Status | Action_Date | User_ID | Comment |

این جدول پایه Audit است؛ هر انتقال وضعیت / هر اقدام مهم در آن ثبت می‌شود. منابع آن در نسخه Excel «دستی/با VBA اختیاری» و در نسخه App خودکار است — اما ساختار دقیقاً همین است.

## 3.8 جدول‌های Workbook (فهرست نهایی)

داده‌های پایه (Reference/Master):
`README, Lists, Settings, Statuses, Departments, Users, Roles, Approval_Matrix, RACI, Customers, Products, Product_Revisions, Materials, Suppliers, Currencies(داخل Lists)`

عملیات (Transaction):
`Orders, Order_Lines, Projects, Project_Statuses, Engineering, BOM_Header, BOM_Detail, Supplier_Quotes, Supplier_Quote_Lines, Procurement, Purchase_Prices, Planning, Schedule, Budget, Costing, Cost_Lines, Sales_Quotation, Quotation_Versions, Payment_Terms, Management_Approval`

اکسسوار Audit/کنترل:
`Workflow_History, Change_Log, Issues, Error_Checks, Test_Results`

ارائه:
`Dashboard, Workflow, _Trace, Report_Quotation, _NamedRanges(مستند)`

جمعاً ~۴۰ شیت؛ ترکیب دقیق در سند 04 (Workbook Architecture) با ستون‌های هر جدول در سند Data Dictionary (05).
