# ExcelForBorna

سامانه جامع Excel برای مدیریت فرآیند **«دریافت سفارش تا تصویب قیمت فروش»**
(Order → Engineering → BOM → Procurement → Costing → Quotation → Approval)

یک Workflow سازمانی قابل ردیابی — نه یک فرم ساده — با معماری Database-like،
نسخه‌بندی، کنترل تغییرات، داشبورد و ردیابی کامل قیمت.

---

## تحویلی‌ها (Deliverables)

### A. فایل Excel نهایی
**`output/ExcelForBorna.xlsx`** — ۴۴ شیت، Excel Table + Structured References،
Data Validation، Conditional Formatting، فرمول‌های شفاف و قابل Audit.

### B..M. مستندات معماری (`docs/`)
| فایل | محتوا |
|---|---|
| `01-Process_Map.md` | نقشه فرآیند + Workflow + نقاط ضعف/ابهام صریح + I/O هر مرحله |
| `02-Workflow_Approval_Matrix_RACI.md` | کارگردانی گردش کار، گیت‌ها، ماتریس تأیید، RACI |
| `03-Data_Model_and_Relationships.md` | مدل داده، استاندارد شناسه‌ها، ER، مسیر ردیابی، BOM چندسطحی |
| `04-Workbook_Architecture.md` | معماری Workbook، قرارداد Layout، تصمیم‌های کلیدی |
| `05-Data_Dictionary.md` | فرهنگ داده کامل (مطابق اسکیمای ساخته‌شده) |
| `06-Calculation_Logic.md` | همه فرمول‌های اصلی (شفاف و قابل Audit) |
| `07-Error_Codes_and_Test_Plan.md` | کدهای خطا (E-001..E-021) + برنامه تست ۱۵ مورد |
| `08-Test_Report.md` | گزارش تست — 15/15 PASS + نتیجه زنجیره مالی |
| `09-User_Guide.md` | راهنمای کاربری |
| `10-Excel_Limitations_and_Future_Architecture.md` | محدودیت‌های Excel + معماری آینده Web/DB |
| `Test_Results.csv` | خروجی ماشینی ۱۵ تست |

### سورس (قابل توسعه / قابل تکرار)
| فایل | نقش |
|---|---|
| `src/seed.py` | داده فرضی (Single Source of Truth؛ نرخ‌ها با برچسب ASSUMPTION) |
| `src/model.py` | مدل درون‌حافظه‌ای Workbook |
| `src/builder.py` | ساخت ۴۴ شیت + فرمول‌های A1 |
| `src/formula_engine.py` | موتور راستی‌آزمایی فرمول (دترمینیستیک) |
| `src/verify.py` | راستی‌آزمایی زنجیره مالی (Total Cost/Margin/Proposed) |
| `src/tests.py` | اجرای ۱۵ تست اجباری |
| `src/renderer.py` | رندر نهایی به .xlsx (جدول/اعتبارسنجی/قالب/CF) |
| `src/build_all.py` | اجرای کامل پایپ‌لاین |

## بازتولید فایل
```bash
cd ExcelForBorna
.venv/bin/python src/build_all.py     # ساخت + راستی‌آزمایی + ۱۵ تست + رندر
```

## اعداد تأییدشده دمو (IRR)
| شاخص | مقدار |
|---|---|
| Total Cost | 9,062,112,475 |
| Cost / Unit | 4,531,056,238 |
| Proposed (Scenario B) | 10,156,441,748 |

## اصل صداقت (anti-fabrication)
هیچ عدد/نرخ سازمانی داخل فرمول Hard-code نیست؛ همه نرخ‌ها در `Settings` با
برچسب `ASSUMPTION` هستند و عناصر فاقد اطلاعات سازمانی صراحتاً اعلام شده‌اند.
