# 02 — Workflow، ماتریس تأیید و RACI
نسخه: 1.0

## 2.1 زنجیره وضعیت سراسری
```
NEW ─► SALES REVIEW ─► PROJECT CREATED ─► ENGINEERING ─► BOM APPROVAL
    ─► PLANNING ─► PROCUREMENT ─► COSTING ─► FINANCE REVIEW
    ─► SALES QUOTATION ─► MANAGEMENT APPROVAL ─► APPROVED / REVISION / REJECTED
```
مسیرهای بازگشتی:
- هر `RETURN FOR REVISION` در Approval، وضعیت پیشنهاد را «Returned» و پروژه را به گام مشخص‌شده برمی‌گرداند.
- تغییر هر ورودی تأییدشده (BOM، قیمت، مقدار، شرایط پرداخت، زمان، قیمت فروش) = **Revision جدید** در موجودیت مربوط + رکورد `Change_Log` + رکورد `Workflow_History`.

## 2.2 گیت‌های اجباری (Gating Rules)
هر گیت یک Rule-Key دارد و در `Error_Checks` نظارت می‌شود:

| Gate | Rule | فرمول کنترلی (نمونه منطق) |
|---|---|---|
| G1 ورود به مهندسی | Order کامل | `=IF(AND(CO_01..CO_07=سبز),"UNLOCK","LOCK")` |
| G2 ورود به Planning | BOM Released | `=IF(BOM Header Status="Approved","UNLOCK","LOCK")` |
| G3 ورود به Costing | همه اقلام قیمت فعال دارند | `=IF(Count(MAT با قیمت فعال)=Count(MAT موردنیاز),"UNLOCK","LOCK")` |
| G4 ورود به Quotation | Costing Approved | `=IF(Costing.Status="Approved","UNLOCK","LOCK")` |
| G5 ورود به Management | Quotation مناسب با مبنا | `=IF(AND(Quotation Submitted, Basis_Memo<>""),"UNLOCK","LOCK")` |

هر «UNLOCK» صرفاً یک نمایش است؛ در کنار هر سلول تصمیم، فرمول `Allowed` پیش‌نیاز را وارسی می‌کند تا مرحله بعد بدون تکمیل پیش‌نیاز سبز نشود.

## 2.3 ماتریس تأیید (Approval Matrix)
| Entity | Gate/تصمیم | Approver Role (نقش) | Min Approval_Level | R | A | C | I |
|---|---|---|---|---|---|---|---|
| Order | Sales Review | Sales Manager | 2 | Sales | Sales Mgr | — | PM |
| Project | Project Created | PMO / PM | 3 | PM | PMO | Sales | All |
| BOM | BOM Approval | Engineering Lead | 3 | Eng | Eng Lead | QA/Prod | PM |
| Supplier Selection | Compare+Select | Procurement Mgr | 3 | Procurement | Proc Mgr | Eng | Finance |
| Purchase Price | Price Confirm | Procurement Mgr | 3 | Procurement | Proc Mgr | Finance | — |
| Costing | Finance Review | Finance Mgr | 4 | Finance | Fin Mgr | — | Sales |
| Quotation | Sales Quotation | Sales Mgr | 3 | Sales | Sales Mgr | Finance | Mgmt |
| Quotation | Management Approval | Senior Mgmt | 5 | Mgmt | Senior Mgmt | All | All |
| تغییر پس از تصویب (هر موجودیت) | Revision | Owner + Approver مرتبط | ≥ سطح گیت | Owner | Approver | — | Mgmt |

## 2.4 RACI ریز (عرض فرآیند)
| فعالیت | Sales | Eng | Planning | Procurement | Finance | PM | Mgmt |
|---|---|---|---|---|---|---|---|
| ثبت سفارش | R/A | I | — | — | — | — | — |
| کنترل اولیه | R | C | — | — | — | I | — |
| ایجاد پروژه | I | I | I | I | I | R/A | I |
| مهندسی + ساخت BOM | C | R/A | C | C | — | I | — |
| تأیید BOM | I | A | C | C | — | I | I |
| برنامه‌ریزی نیاز | I | C | R/A | I | — | I | — |
| زمان‌بندی | I | C | R/A | C | — | I | I |
| استعلام تأمین‌کننده | — | C | I | R | I | I | — |
| انتخاب تأمین‌کننده | — | C | I | R/A | C | I | I |
| ثبت قیمت خرید | — | C | I | R | I | I | — |
| محاسبه بهای تمام‌شده | C | C | C | C | R/A | I | I |
| بودجه | I | C | C | C | R/A | I | I |
| پیشنهاد قیمت فروش | R/A | C | I | C | C | I | I |
| شرایط پرداخت | R | — | — | — | C | — | A |
| تصمیم مدیریت ارشد | I | I | I | I | I | I | R/A |
| ثبت تغییر/نسخه | R (مالک) | R (مالک) | R (مالک) | R (مالک) | R (مالک) | A | I |

## 2.5 نکات اجرایی در فایل
- جدول RACI در شیت `RACI` (Entity, Action, Sales, Eng, Planning, Procurement, Finance, PMO, Mgmt) با ۱ خالی و Conditional Formatting.
- جدول Approval_Matrix با ستون‌های Entity, Gate, Approver_Role, Min_Approval_Level, Responsible, Reviewer, Approver, Informed.
