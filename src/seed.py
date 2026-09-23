"""
seed.py — demo dataset (single source of truth)
================================================
One realistic end-to-end scenario. Every number chosen here is labelled
ASSUMPTION in Settings/docs; the workbook never hard-codes rates in formulas.

Money currency for the demo: IRR (reporting currency). A few USD quotes exist
to prove multi-currency handling; Fx rates live in Settings.
"""
from __future__ import annotations

import datetime as dt

TODAY = dt.date(2026, 9, 22)

# --------------------------------------------------------------------------- #
# Reference data
# --------------------------------------------------------------------------- #
STATUSES = [
    # ID, name, group, sequence
    ("ST-01", "Draft", "Draft", 1),
    ("ST-02", "In Review", "Review", 2),
    ("ST-03", "Waiting Approval", "Approval", 3),
    ("ST-04", "Approved", "Approved", 4),
    ("ST-05", "Rejected", "Rejected", 5),
    ("ST-06", "Returned", "Rejected", 6),
    ("ST-07", "Closed", "Closed", 7),
    ("ST-08", "Cancelled", "Closed", 8),
]

DEPARTMENTS = [
    ("DEP-01", "Sales"),
    ("DEP-02", "Engineering"),
    ("DEP-03", "Planning"),
    ("DEP-04", "Procurement"),
    ("DEP-05", "Finance"),
    ("DEP-06", "PMO"),
    ("DEP-07", "Management"),
]

ROLES = [
    ("ROL-01", "Sales Manager"),
    ("ROL-02", "Sales Rep"),
    ("ROL-03", "Project Manager"),
    ("ROL-04", "Engineering Lead"),
    ("ROL-05", "Planning Manager"),
    ("ROL-06", "Procurement Manager"),
    ("ROL-07", "Finance Manager"),
    ("ROL-08", "Senior Management"),
    ("ROL-09", "Admin"),
    ("ROL-10", "Warehouse Keeper"),
]

USERS = [
    # User_ID, Name, Dept_ID, Role_ID, Email, Approval_Level, Is_Active
    ("USR-01", "Babak A", "DEP-01", "ROL-01", "babak@example.com", 2, 1),
    ("USR-02", "Sara S", "DEP-01", "ROL-02", "sara@example.com", 1, 1),
    ("USR-03", "Reza M", "DEP-06", "ROL-03", "reza@example.com", 3, 1),
    ("USR-04", "Neda E", "DEP-02", "ROL-04", "neda@example.com", 3, 1),
    ("USR-05", "Kaveh P", "DEP-03", "ROL-05", "kaveh@example.com", 2, 1),
    ("USR-06", "Mina T", "DEP-04", "ROL-06", "mina@example.com", 3, 1),
    ("USR-07", "Ali F", "DEP-05", "ROL-07", "ali@example.com", 4, 1),
    ("USR-08", "Borna G", "DEP-07", "ROL-08", "borna@example.com", 5, 1),
    ("USR-10", "Vahid K", "DEP-03", "ROL-10", "vahid@example.com", 1, 1),
]

APPROVAL_MATRIX = [
    # Entity, Gate, Approver_Role, Min_Approval_Level, Responsible, Reviewer, Approver, Informed
    ("Order", "Sales Review", "ROL-01", 2, "Sales Rep", "Sales Manager", "Sales Manager", "PMO"),
    ("Project", "Project Created", "ROL-03", 3, "Project Manager", "PMO", "PMO", "All"),
    ("BOM", "BOM Approval", "ROL-04", 3, "Engineering", "Engineering Lead", "Engineering Lead", "PMO"),
    ("Supplier", "Supplier Selection", "ROL-06", 3, "Procurement", "Procurement Manager", "Procurement Manager", "Finance"),
    ("Purchase Price", "Price Confirm", "ROL-06", 3, "Procurement", "Procurement Manager", "Procurement Manager", "Finance"),
    ("Costing", "Finance Review", "ROL-07", 4, "Finance", "Finance Manager", "Finance Manager", "Sales"),
    ("Quotation", "Sales Quotation", "ROL-01", 3, "Sales Manager", "Sales Manager", "Sales Manager", "Management"),
    ("Quotation", "Management Approval", "ROL-08", 5, "Sales Manager", "Senior Management", "Senior Management", "All"),
    ("*", "Post-Approval Change", "ROL-08", 5, "Owner", "Owner Approver", "Owner Approver", "Management"),
]

RACI = [
    # Activity, Sales, Engineering, Planning, Procurement, Finance, PMO, Management
    ("Register Order", "R/A", "I", "", "", "", "", ""),
    ("Initial Check", "R", "C", "", "", "", "I", ""),
    ("Create Project", "I", "I", "I", "I", "I", "R/A", "I"),
    ("Engineering + BOM", "C", "R/A", "C", "C", "", "I", ""),
    ("BOM Approval", "I", "A", "C", "C", "", "I", "I"),
    ("Planning", "I", "C", "R/A", "I", "", "I", ""),
    ("Scheduling", "I", "C", "R/A", "C", "", "I", "I"),
    ("Supplier Enquiry", "", "C", "I", "R", "I", "I", ""),
    ("Supplier Selection", "", "C", "I", "R/A", "C", "I", "I"),
    ("Purchase Price Register", "", "C", "I", "R", "I", "I", ""),
    ("Costing", "C", "C", "C", "C", "R/A", "I", "I"),
    ("Budget", "I", "C", "C", "C", "R/A", "I", "I"),
    ("Sales Quotation", "R/A", "C", "I", "C", "C", "I", "I"),
    ("Payment Terms", "R", "", "", "", "C", "", "A"),
    ("Management Decision", "I", "I", "I", "I", "I", "I", "R/A"),
]

CUSTOMERS = [
    # CID, code, name, legal, type, country, city, delivery, payment_cond, credit_limit, currency, tax, active, owner, notes
    ("CUS-000001", "CUS-1", "صنایع پتروشیمی صبا", "شرکت صنایع پتروشیمی صبا (سهامی خاص)",
     "Corporate", "Iran", "عسلویه", "سایت عسلویه، فاز ۲",
     "30% پیش‌پرداخت + 70% اعتبار 60 روزه", 25000000000, "IRR", "VAT", 1, "USR-02",
     "مشتری کلیدی؛ سابقه پرداخت خوب"),
    ("CUS-000002", "CUS-2", "پالایش نفت ستاره", "شرکت پالایش نفت ستاره",
     "Corporate", "Iran", "بندرعباس", "انبار مرکزی بندرعباس",
     "LC دیداری", 20000000000, "IRR", "VAT", 1, "USR-02", "سفارش‌های فصلی"),
]

PRODUCTS = [
    # PID, code, name, type, uom, category, active
    ("PRD-000125", "PMP-250", "پمپ گریز از مرکز سرویس سنگین PMP-250", "Finished Good", "pcs", "Pumps", 1),
    ("PRD-000126", "PMP-150", "پمپ گریز از مرکز PMP-150", "Finished Good", "pcs", "Pumps", 1),
]

PRODUCT_REVISIONS = [
    # PRV id, product, rev no, drawing, drawing rev, bom policy, approved by, date, reason
    ("PRV-000125-R03", "PRD-000125", "R03", "DRW-250-A", "Rev C", "Each Rev = New BOM Revision", "USR-04", "2026-09-10",
     "به‌روزرسانی متریال پروانه طبق استاندارد جدید"),
    ("PRV-000126-R02", "PRD-000126", "R02", "DRW-150-B", "Rev B", "Each Rev = New BOM Revision", "USR-04", "2026-06-01",
     "اصلاح تلورانس"),
]

MATERIALS = [
    # MID, code, description, type, spec, uom, make_buy, preferred supplier, active
    ("MAT-001542", "MAT-01", "پوسته چدنی (Cast Iron Casing)", "Raw Material",
     "GG-25 مطابق DIN 1691", "kg", "Make", "", 1),
    ("MAT-001543", "MAT-02", "پروانه فولاد ضدزنگ (Impeller SS316)", "Raw Material",
     "AISI 316L", "kg", "Make", "", 1),
    ("MAT-001544", "MAT-03", "بلبرینگ کف‌گرد (Thrust Bearing)", "Purchased Part",
     "SKF 29420 E", "pcs", "Buy", "SUP-000089", 1),
    ("MAT-001545", "MAT-04", "مکانیکال سیل (Mechanical Seal)", "Purchased Part",
     "EagleBurgmann H74", "pcs", "Buy", "SUP-000088", 1),
    ("MAT-001546", "MAT-05", "الکتروموتور 45kW", "Purchased Part",
     "IEC 400V 50Hz IM B3", "pcs", "Buy", "SUP-000087", 1),
    ("MAT-001547", "MAT-06", "شفت فولادی (Shaft CK45)", "Raw Material",
     "CK45 مطابق DIN 17200", "kg", "Make", "", 1),
    ("MAT-001548", "MAT-07", "بسته‌بندی صادراتی چوبی", "Packaging",
     "ISPM-15", "pcs", "Buy", "SUP-000086", 1),
]

SUPPLIERS = [
    # SID, code, name, legal, country, currency, payment default, incoterm, rating, approval, contact, active
    ("SUP-000086", "SUP-86", "بسته‌بندی پارس", "بسته‌بندی پارس (چوب)", "Iran", "IRR",
     "30 روزه", "EXW", 4, "Approved", "تلفن 021-1111", 1),
    ("SUP-000087", "SUP-87", "الکتروموتور توان", "موتور توان (نماینده زیمنس)", "Iran", "IRR",
     "45 روزه", "CPT", 4, "Approved", "تلفن 021-2222", 1),
    ("SUP-000088", "SUP-88", "سیل آریا", "سیل‌های مکانیکی آریا", "Iran", "IRR",
     "30 روزه", "CPT", 5, "Approved", "تلفن 021-3333", 1),
    ("SUP-000089", "SUP-89", "بلبرینگ ایران", "بلبرینگ‌های صنعتی ایران", "Iran", "IRR",
     "30 روزه", "CPT", 4, "Approved", "تلفن 021-4444", 1),
    ("SUP-000090", "SUP-90", "فولاد جنوب", "فولاد جنوب (ریخته‌گری)", "Iran", "IRR",
     "45 روزه", "EXW", 3, "Approved", "تلفن 071-5555", 1),
]

# --------------------------------------------------------------------------- #
# Orders / Lines
# --------------------------------------------------------------------------- #
ORDERS = [
    # OID, customer, project(ref later), order date, requested delivery, type,
    # delivery terms, payment terms, validity days, confidentiality, sales, doc, status
    ("ORD-2026-00321", "CUS-000001", "PRJ-2026-00125", "2026-09-20", "2026-12-20",
     "New", "CPT سایت عسلویه", "30% پیش‌پرداخت + 70% اعتبار 60 روزه", 30,
     "Confidential", "USR-02", "RFQ-SABA-2026-09-001", "Approved"),
    ("ORD-2026-00322", "CUS-000002", "PRJ-2026-00126", "2026-09-21", "2027-01-15",
     "New", "CPT بندرعباس", "LC", 45,
     "Internal", "USR-02", "RFQ-STAREH-2026-09-002", "Draft"),
    # incomplete order on purpose → triggers CO checks / Issue (Test T2 demo)
    ("ORD-2026-00323", "CUS-000001", "", "2026-09-21", "",
     "New", "", "", 30,
     "Internal", "USR-02", "RFQ-SABA-2026-09-009", "Draft"),
]

ORDER_LINES = [
    # OLN, order, line, product, qty, uom, latest
    ("OLN-000456", "ORD-2026-00321", 1, "PRD-000125", 2, "pcs", 1),
    ("OLN-000457", "ORD-2026-00321", 2, "PRD-000126", 1, "pcs", 1),
    ("OLN-000458", "ORD-2026-00322", 1, "PRD-000125", 1, "pcs", 1),
]

PROJECTS = [
    # PID, code, order, customer, product(rep), name, PM, start, due, priority, status, progress, sponsor, team
    ("PRJ-2026-00125", "PRJ-001", "ORD-2026-00321", "CUS-000001", "PRD-000125",
     "پمپ PMP-250 صبا", "USR-03", "2026-09-21", "2026-12-20", "High", "Approved", 45,
     "USR-08", "Neda E, Kaveh P, Mina T, Ali F"),
    ("PRJ-2026-00126", "PRJ-002", "ORD-2026-00322", "CUS-000002", "PRD-000125",
     "پمپ PMP-250 ستاره", "USR-03", "2026-09-22", "2027-01-15", "Medium", "Draft", 5,
     "USR-08", "TBD"),
]

PROJECT_STATUSES = [
    # PSID, project, stage, stage status, start, end, remark
    ("PS-0001", "PRJ-2026-00125", "Sales", "Completed", "2026-09-20", "2026-09-21", ""),
    ("PS-0002", "PRJ-2026-00125", "Engineering", "Completed", "2026-09-21", "2026-09-30", "BOM R5 تأیید شد"),
    ("PS-0003", "PRJ-2026-00125", "Planning", "In Progress", "2026-10-01", None, "برآورد نیاز در حال انجام"),
    ("PS-0004", "PRJ-2026-00125", "Procurement", "Pending", None, None, ""),
    ("PS-0005", "PRJ-2026-00125", "Finance", "Pending", None, None, ""),
    ("PS-0006", "PRJ-2026-00125", "Management", "Pending", None, None, ""),
    ("PS-0007", "PRJ-2026-00126", "Sales", "In Review", "2026-09-21", None, ""),
    ("PS-0008", "PRJ-2026-00126", "Engineering", "Pending", None, None, ""),
    ("PS-0009", "PRJ-2026-00126", "Planning", "Pending", None, None, ""),
    ("PS-0010", "PRJ-2026-00126", "Procurement", "Pending", None, None, ""),
    ("PS-0011", "PRJ-2026-00126", "Finance", "Pending", None, None, ""),
    ("PS-0012", "PRJ-2026-00126", "Management", "Pending", None, None, ""),
]

ENGINEERING = [
    # EID, project, product, product rev, drawing, drawing rev, bom revision(ref), status
    ("ENG-000125", "PRJ-2026-00125", "PRD-000125", "PRV-000125-R03", "DRW-250-A", "Rev C", "BOM-000215-R05", "Approved"),
    ("ENG-000126", "PRJ-2026-00125", "PRD-000126", "PRV-000126-R02", "DRW-150-B", "Rev B", "BOM-000216-R02", "Approved"),
]

# --------------------------------------------------------------------------- #
# BOM — multi-level
# --------------------------------------------------------------------------- #
BOM_HEADERS = [
    # BMR id, BOM id, product revision, rev no, rev date, reason, changed by, approved by, status, active
    ("BOM-000215-R05", "BOM-000215", "PRV-000125-R03", "R05", "2026-09-28", "تغییر تأمین‌کننده مکانیکال سیل و افزودن شفت", "USR-04", "USR-04", "Approved", 1),
    ("BOM-000215-R04", "BOM-000215", "PRV-000125-R03", "R04", "2026-08-01", "نسخه قبلی (نگهداری برای Audit — بدون Overwrite)", "USR-04", "USR-04", "Superseded", 0),
    ("BOM-000216-R02", "BOM-000216", "PRV-000126-R02", "R02", "2026-06-05", "نسخه اولیه", "USR-04", "USR-04", "Approved", 1),
]

# (BLI, BMR, parent, level, seq, material, qty_per, uom, scrap%, make_buy, approved_supplier)
BOM_DETAILS = [
    # Product PMP-250, revision R05 (ACTIVE)
    ("BLI-000987", "BOM-000215-R05", "", 0, 10, "PRD-000125", 1.0, "pcs", 0.0, "Make", ""),
    ("BLI-000988", "BOM-000215-R05", "BLI-000987", 1, 20, "MAT-001542", 180.0, "kg", 0.05, "Make", ""),
    ("BLI-000989", "BOM-000215-R05", "BLI-000987", 1, 30, "MAT-001543", 55.0, "kg", 0.05, "Make", ""),
    ("BLI-000990", "BOM-000215-R05", "BLI-000987", 1, 40, "MAT-001547", 22.0, "kg", 0.05, "Make", ""),
    ("BLI-000991", "BOM-000215-R05", "BLI-000987", 1, 50, "MAT-001544", 1.0, "pcs", 0.0, "Buy", "SUP-000089"),
    ("BLI-000992", "BOM-000215-R05", "BLI-000987", 1, 60, "MAT-001545", 1.0, "pcs", 0.0, "Buy", "SUP-000088"),
    ("BLI-000993", "BOM-000215-R05", "BLI-000987", 1, 70, "MAT-001546", 1.0, "pcs", 0.0, "Buy", "SUP-000087"),
    ("BLI-000994", "BOM-000215-R05", "BLI-000987", 1, 80, "MAT-001548", 1.0, "pcs", 0.0, "Buy", "SUP-000086"),
    # Product PMP-250 revision R04 (previous, kept for non-overwrite proof)
    ("BLI-000950", "BOM-000215-R04", "", 0, 10, "PRD-000125", 1.0, "pcs", 0.0, "Make", ""),
    ("BLI-000951", "BOM-000215-R04", "BLI-000950", 1, 20, "MAT-001542", 180.0, "kg", 0.05, "Make", ""),
    ("BLI-000952", "BOM-000215-R04", "BLI-000950", 1, 30, "MAT-001543", 55.0, "kg", 0.05, "Make", ""),
    ("BLI-000953", "BOM-000215-R04", "BLI-000950", 1, 40, "MAT-001547", 22.0, "kg", 0.05, "Make", ""),
    ("BLI-000954", "BOM-000215-R04", "BLI-000950", 1, 50, "MAT-001544", 1.0, "pcs", 0.0, "Buy", "SUP-000089"),
    ("BLI-000955", "BOM-000215-R04", "BLI-000950", 1, 60, "MAT-001545", 1.0, "pcs", 0.0, "Buy", "SUP-000088"),
    ("BLI-000956", "BOM-000215-R04", "BLI-000950", 1, 70, "MAT-001546", 1.0, "pcs", 0.0, "Buy", "SUP-000087"),
    ("BLI-000957", "BOM-000215-R04", "BLI-000950", 1, 80, "MAT-001548", 1.0, "pcs", 0.0, "Buy", "SUP-000086"),
    # Product PMP-150
    ("BLI-000900", "BOM-000216-R02", "", 0, 10, "PRD-000126", 1.0, "pcs", 0.0, "Make", ""),
    ("BLI-000901", "BOM-000216-R02", "BLI-000900", 1, 20, "MAT-001542", 95.0, "kg", 0.05, "Make", ""),
    ("BLI-000902", "BOM-000216-R02", "BLI-000900", 1, 30, "MAT-001547", 14.0, "kg", 0.05, "Make", ""),
    ("BLI-000903", "BOM-000216-R02", "BLI-000900", 1, 40, "MAT-001544", 1.0, "pcs", 0.0, "Buy", "SUP-000089"),
    ("BLI-000904", "BOM-000216-R02", "BLI-000900", 1, 50, "MAT-001545", 1.0, "pcs", 0.0, "Buy", "SUP-000088"),
    ("BLI-000905", "BOM-000216-R02", "BLI-000900", 1, 60, "MAT-001546", 1.0, "pcs", 0.0, "Buy", "SUP-000087"),
    ("BLI-000906", "BOM-000216-R02", "BLI-000900", 1, 70, "MAT-001548", 1.0, "pcs", 0.0, "Buy", "SUP-000086"),
]

# --------------------------------------------------------------------------- #
# Supplier quotes (multi-supplier comparison)
# --------------------------------------------------------------------------- #
# (SQL, SQT, supplier, unit_price, currency, moq, lead, pay, incoterm, freight, valid_to, source, confirm,
#   quality, ontime, capacity, risk, longevity, selected)
SUPPLIER_QUOTE_LINES = [
    # Mechanical Seal (MAT-001545) — 3 suppliers
    ("SQL-002354", "SQT-2026-00545", "SUP-000088", 185_000_000, "IRR", 1, 30, "30 days", "CPT",
     2_000_000, "2026-11-30", "Email Quote 12-Sep", "Confirmed", 5, 5, 4, 2, 3, 1),
    ("SQL-002355", "SQT-2026-00545", "SUP-000089", 172_000_000, "IRR", 2, 45, "60 days", "CPT",
     2_500_000, "2026-11-15", "Portal RFQ", "Confirmed", 4, 4, 3, 3, 3, 0),
    ("SQL-002356", "SQT-2026-00545", "SUP-000090", 190_000_000, "IRR", 1, 25, "advance", "EXW",
     3_000_000, "2026-12-10", "Phone + Email", "Confirmed", 3, 3, 2, 4, 2, 0),
    # Thrust Bearing (MAT-001544) — 2 suppliers
    ("SQL-002360", "SQT-2026-00546", "SUP-000089", 95_000_000, "IRR", 1, 25, "30 days", "CPT",
     1_200_000, "2026-11-20", "Email Quote 14-Sep", "Confirmed", 5, 4, 4, 2, 4, 1),
    ("SQL-002361", "SQT-2026-00546", "SUP-000090", 99_000_000, "IRR", 1, 30, "45 days", "EXW",
     1_500_000, "2026-11-10", "Portal RFQ", "Confirmed", 4, 3, 3, 3, 3, 0),
    # Motor 45kW (MAT-001546) — 2 suppliers (multi-currency: USD vs IRR)
    ("SQL-002364", "SQT-2026-00547", "SUP-000087", 2_500, "USD", 1, 45, "45 days", "CPT",
     30, "2026-12-01", "Signed Quote", "Confirmed", 5, 4, 5, 2, 5, 1),
    ("SQL-002365", "SQT-2026-00547", "SUP-000090", 1_220_000_000, "IRR", 2, 60, "60 days", "EXW",
     18_000_000, "2026-11-25", "Email Quote", "Confirmed", 4, 3, 4, 3, 4, 0),
    # Packaging (MAT-001548)
    ("SQL-002368", "SQT-2026-00548", "SUP-000086", 12_000_000, "IRR", 1, 10, "cash", "EXW",
     0, "2026-12-20", "Price list", "Confirmed", 4, 5, 5, 1, 5, 1),
]

SUPPLIER_QUOTES = [
    # SQT id, project, material, qty required, quote date, valid until, status, selected line
    ("SQT-2026-00545", "PRJ-2026-00125", "MAT-001545", 2, "2026-09-15", "2026-12-10", "Closed", "SQL-002354"),
    ("SQT-2026-00546", "PRJ-2026-00125", "MAT-001544", 2, "2026-09-15", "2026-11-20", "Closed", "SQL-002360"),
    ("SQT-2026-00547", "PRJ-2026-00125", "MAT-001546", 2, "2026-09-15", "2026-12-01", "Closed", "SQL-002364"),
    ("SQT-2026-00548", "PRJ-2026-00125", "MAT-001548", 2, "2026-09-15", "2026-12-20", "Closed", "SQL-002368"),
]

# --------------------------------------------------------------------------- #
# Procurement / Purchase prices (Price build-up → Landed)
# --------------------------------------------------------------------------- #
# (PPR, material, supplier, revision, price date, source, currency, fx, unit price,
#  freight, insurance, customs, handling, other, valid from, valid to, status, approved by)
PURCHASE_PRICES = [
    ("PPR-000554", "MAT-001542", "SUP-000090", "R01", "2026-09-15", "Signed Quote", "USD",
     500000, 1.15, 0.05, 0, 0, 0, 0, "2026-09-15", "2026-11-30", "Approved", "USR-06"),
    ("PPR-000555", "MAT-001543", "SUP-000090", "R01", "2026-09-15", "Signed Quote", "USD",
     500000, 3.2, 0.05, 0, 0, 0, 0, "2026-09-15", "2026-11-30", "Approved", "USR-06"),
    ("PPR-000556", "MAT-001544", "SUP-000089", "R01", "2026-09-15", "Email Quote", "IRR",
     1, 95_000_000, 1_200_000, 0, 0, 0, 0, "2026-09-15", "2026-11-20", "Approved", "USR-06"),
    ("PPR-000557", "MAT-001545", "SUP-000088", "R01", "2026-09-15", "Email Quote", "IRR",
     1, 185_000_000, 2_000_000, 0, 0, 0, 0, "2026-09-15", "2026-11-30", "Approved", "USR-06"),
    ("PPR-000558", "MAT-001546", "SUP-000087", "R01", "2026-09-15", "Signed Quote", "USD",
     500000, 2_500, 30, 0, 0, 0, 0, "2026-09-15", "2026-12-01", "Approved", "USR-06"),
    ("PPR-000559", "MAT-001547", "SUP-000090", "R01", "2026-09-15", "Signed Quote", "IRR",
     1, 2_200_000, 0, 0, 0, 0, 0, "2026-09-15", "2026-09-10", "Expired", "USR-06"),
    ("PPR-000560", "MAT-001547", "SUP-000090", "R02", "2026-09-30", "Revised Quote", "IRR",
     1, 2_450_000, 0, 0, 0, 0, 0, "2026-10-01", "2026-12-01", "Approved", "USR-06"),
    ("PPR-000561", "MAT-001548", "SUP-000086", "R01", "2026-09-15", "Price list", "IRR",
     1, 12_000_000, 0, 0, 0, 0, 0, "2026-09-15", "2026-12-20", "Approved", "USR-06"),
]

# (Stock_ID, Material_ID, Stock_On_Hand, Reserved)
INVENTORY = [
    ("INV-0001", "MAT-001542", 500.0, 0.0),
    ("INV-0002", "MAT-001543", 120.0, 0.0),
    ("INV-0003", "MAT-001544", 6.0, 2.0),
    ("INV-0004", "MAT-001545", 4.0, 1.0),
    ("INV-0005", "MAT-001546", 2.0, 0.0),
    ("INV-0006", "MAT-001547", 80.0, 0.0),
    ("INV-0007", "MAT-001548", 3.0, 0.0),
]

# --------------------------------------------------------------------------- #
# Planning (per BOM line for the active revision, incl. extra/expired PRJ rows)
# --------------------------------------------------------------------------- #
PLANNING = []

# --------------------------------------------------------------------------- #
# Schedule
# --------------------------------------------------------------------------- #
SCHEDULE = [
    # SCH, project, activity, start, end, responsible, dependency, status
    ("SCH-000256", "PRJ-2026-00125", "Engineering", "2026-09-21", "2026-09-30", "USR-04", "", "Done"),
    ("SCH-000257", "PRJ-2026-00125", "BOM Release", "2026-09-28", "2026-09-30", "USR-04", "Engineering", "Done"),
    ("SCH-000258", "PRJ-2026-00125", "Procurement", "2026-10-01", "2026-10-20", "USR-06", "BOM Release", "In Progress"),
    ("SCH-000259", "PRJ-2026-00125", "Supplier Confirmation", "2026-10-05", "2026-10-20", "USR-06", "Procurement", "In Progress"),
    ("SCH-000260", "PRJ-2026-00125", "Material Arrival", "2026-11-01", "2026-11-15", "USR-06", "Supplier Confirmation", "Not Started"),
    ("SCH-000261", "PRJ-2026-00125", "Production Planning", "2026-11-10", "2026-11-20", "USR-05", "Material Arrival", "Not Started"),
    ("SCH-000262", "PRJ-2026-00125", "Production", "2026-11-20", "2026-12-05", "USR-05", "Production Planning", "Not Started"),
    ("SCH-000263", "PRJ-2026-00125", "QC", "2026-12-06", "2026-12-10", "USR-04", "Production", "Not Started"),
    ("SCH-000264", "PRJ-2026-00125", "Packing", "2026-12-11", "2026-12-14", "USR-05", "QC", "Not Started"),
    ("SCH-000265", "PRJ-2026-00125", "Delivery", "2026-12-15", "2026-12-18", "USR-02", "Packing", "Not Started"),
]

# --------------------------------------------------------------------------- #
# Budget
# --------------------------------------------------------------------------- #
# (BDG, project, revision, period, b_mat, b_labor, b_oh, b_proc, b_log, b_tool, b_other)
BUDGET = [
    ("BDG-000245", "PRJ-2026-00125", "R01", "2026-Q4",
     4_000_000_000, 350_000_000, 700_000_000, 200_000_000, 150_000_000, 100_000_000, 50_000_000),
]

# --------------------------------------------------------------------------- #
# Costing
# --------------------------------------------------------------------------- #
# (CST, project, revision, status, cost basis date, qty basis, approver)
COSTING = [
    ("CST-000321", "PRJ-2026-00125", "R01", "Approved", "2026-10-05", 2, "USR-07"),
]

# --------------------------------------------------------------------------- #
# Sales quotations (scenarios A/B/C)
# --------------------------------------------------------------------------- #
# (QUO, project, code, scenario, costing ref, margin%, comm, risk, fin, currency, basis memo, status)
SALES_QUOTATION = [
    ("QUO-2026-00125", "PRJ-2026-00125", "QUO-001", "Scenario A", "CST-000321",
     0.12, 20_000_000, 0, 0, "IRR",
     "Margin 12% روی کل هزینه به علاوه پوشش هزینه بازرگانی؛ بدون تعدیل ریسک", "Submitted"),
    ("QUO-2026-00125", "PRJ-2026-00125", "QUO-001", "Scenario B", "CST-000321",
     0.10, 10_000_000, 15_000_000, 0, "IRR",
     "Margin 10% + پوشش ریسک ارزی/زنجیره 15M؛ گزینه باثبات‌تر", "Submitted"),
    ("QUO-2026-00125", "PRJ-2026-00125", "QUO-001", "Scenario C", "CST-000321",
     0.15, 0, 30_000_000, 0, "IRR",
     "Margin تهاجمی 15% + تعدیل ریسک بالاتر برای تأخیر تأمین", "Draft"),
]

# --------------------------------------------------------------------------- #
# Payment terms
# --------------------------------------------------------------------------- #
PAYMENT_TERMS = [
    # PTID, quotation ref (quotation code), advance%, on delivery%, credit days, installment, bg, currency, risk, active
    ("PT-0001", "QUO-001", 0.30, 0.10, 60, "", "SBLC 10%", "IRR", "Medium", 1),
]

# --------------------------------------------------------------------------- #
# Quotation versions
# --------------------------------------------------------------------------- #
QUOTATION_VERSIONS = [
    # QUV, quotation code, version, costing rev, total cost, margin, selling price, change summary, approved step
    ("QUO-00125-V01", "QUO-001", "V01", "CST-000321", None, None, None, "نسخه اولیه سناریو B برای بررسی مدیریت", "Management"),
]

# --------------------------------------------------------------------------- #
# Management approval
# --------------------------------------------------------------------------- #
MANAGEMENT_APPROVAL = [
    # APR, quotation version, project, decision, maker, date, comment, version ref, status
    ("APR-000547", "QUO-00125-V01", "PRJ-2026-00125", "APPROVED WITH CONDITION", "USR-08", "2026-10-07",
     "مشروط به تأمین مکانیکال سیل قبل از ۱۵ آبان و تعهد کتبی بلبرینگ", "QUO-00125-V01", "Approved"),
]

# --------------------------------------------------------------------------- #
# Change log
# --------------------------------------------------------------------------- #
CHANGE_LOG = [
    # CHG, project, entity type, entity id, field, old, new, reason, by, date, rev after
    ("CHG-001245", "PRJ-2026-00125", "BOM", "BOM-000216", "Material", "MAT-001545", "MAT-001545",
     "تایید تامین‌کننده جدید مکانیکال سیل", "USR-04", "2026-09-28", "R05"),
]

# --------------------------------------------------------------------------- #
# Issues
# --------------------------------------------------------------------------- #
ISSUES = [
    # ISS, project, error code, description, owner, action, status, severity, date
    ("ISS-000215", "PRJ-2026-00125", "IG-02", "قیمت شفت (MAT-001547) در تاریخ ۱۵ اکتبر منقضی می‌شود؛ نیاز به تمدید",
     "USR-06", "استعلام مجدد قبل از ۱۰ اکتبر", "Open", "Medium", "2026-09-22"),
    ("ISS-000216", "", "E-001", "سفارش ORD-2026-00323 ناقص است (تاریخ تحویل و شرایط پرداخت ندارد)",
     "USR-02", "تکمیل اطلاعات سفارش قبل از ورود به مهندسی", "Open", "High", "2026-09-22"),
]

# --------------------------------------------------------------------------- #
# Workflow history
# --------------------------------------------------------------------------- #
WORKFLOW_HISTORY = [
    # WFH, project, entity type, entity id, from, to, date, user, comment
    ("WFH-000001", "PRJ-2026-00125", "Order", "ORD-2026-00321", "Draft", "In Review", "2026-09-20", "USR-02", "ثبت سفارش"),
    ("WFH-000002", "PRJ-2026-00125", "Order", "ORD-2026-00321", "In Review", "Approved", "2026-09-21", "USR-01", "کنترل اولیه تأیید شد"),
    ("WFH-000003", "PRJ-2026-00125", "Project", "PRJ-2026-00125", "Draft", "Approved", "2026-09-21", "USR-03", "تأسیس پروژه"),
    ("WFH-000004", "PRJ-2026-00125", "BOM", "BOM-000215-R05", "Waiting Approval", "Approved", "2026-09-30", "USR-04", "تأیید BOM R05"),
    ("WFH-000005", "PRJ-2026-00125", "Quotation", "QUO-2026-00125", "Draft", "Submitted", "2026-10-05", "USR-01", "ارسال به مدیریت"),
    ("WFH-000006", "PRJ-2026-00125", "Approval", "APR-000547", "Waiting Approval", "Approved", "2026-10-07", "USR-08", "تصویب مشروط"),
]

# --------------------------------------------------------------------------- #
# Settings (ASSUMPTION-labelled rates & sequences)
# --------------------------------------------------------------------------- #
SETTINGS = [
    # key, value, datatype, assumption flag, description
    ("Labor_Rate_Per_Hour", 2_500_000, "Number", 1, "ASSUMPTION: نرخ دستمزد ساعتی (IRR)"),
    ("Manufacturing_Overhead_Rate", 2.0, "Number", 1, "ASSUMPTION: نرخ سربار تولید (× دستمزد مستقیم)"),
    ("Scrap_Factor", 0.02, "Number", 1, "ASSUMPTION: درصد ضایعات فرآیند"),
    ("Financial_Cost_Rate_Annual", 0.18, "Number", 1, "ASSUMPTION: نرخ هزینه مالی سالانه اعتبار"),
    ("SG&A_Rate", 0.05, "Number", 1, "ASSUMPTION: نرخ اداری/فروش (در صورت اعمال سیاست)"),
    ("Apply_SGA", 1, "Number", 1, "1=اعمال SG&A در Costing (سیاست سازمان)"),
    ("Tooling_Cost_Per_Project", 100_000_000, "Number", 1, "ASSUMPTION: هزینه قالب پروژه (تسهیم خطی)"),
    ("Logistics_Base_Cost", 150_000_000, "Number", 1, "ASSUMPTION: هزینه لجستیک پایه پروژه"),
    ("Packaging_Unit_Cost", 12_000_000, "Number", 0, "از قیمت تأمین‌کننده (MAT-001548)"),
    ("Reporting_Currency", "IRR", "Text", 0, "ارز گزارشگری"),
    ("Fx_USD_TO_IRR", 500000, "Number", 1, "ASSUMPTION: نرخ تبدیل USD→IRR"),
    ("Stale_Price_Days", 90, "Number", 0, "آستانه قیمت قدیمی (روز)"),
    ("Max_Margin_Percent", 0.35, "Number", 0, "سقف منطقی مارجین برای E-008"),
    ("Default_Validity_Days", 30, "Number", 0, "اعتبار پیش‌فرض پیشنهاد"),
    # v2.0 — feasibility weights & gating policy
    ("Feasibility_Weight_Market", 0.25, "Number", 1, "ASSUMPTION: وزن امتیاز بازار در امکان‌سنجی"),
    ("Feasibility_Weight_Technical", 0.30, "Number", 1, "ASSUMPTION: وزن امتیاز فنی در امکان‌سنجی"),
    ("Feasibility_Weight_Economic", 0.30, "Number", 1, "ASSUMPTION: وزن امتیاز اقتصادی در امکان‌سنجی"),
    ("Feasibility_Weight_Schedule", 0.15, "Number", 1, "ASSUMPTION: وزن امتیاز زمان‌بندی در امکان‌سنجی"),
    ("Feasibility_Min_Score", 70, "Number", 1, "ASSUMPTION: حداقل امتیاز وزنی برای نتیجه «امکان‌پذیر»"),
    ("Gate_Enforcement", 1, "Number", 0, "سیاست گیتینگ: 1=اجرای سخت (تا پیش‌نیاز تکمیل نشود، ورود داده در مرحله بعد مسدود است)"),
    # Sequence counters
    ("Seq_Customer", 3, "Number", 0, "شمارنده مشتری"),
    ("Seq_Project", 127, "Number", 0, "شمارنده پروژه"),
    ("Seq_Order", 323, "Number", 0, "شمارنده سفارش"),
    ("Seq_Product", 127, "Number", 0, "شمارنده محصول"),
    ("Seq_Material", 1549, "Number", 0, "شمارنده ماده"),
    ("Seq_Supplier", 91, "Number", 0, "شمارنده تأمین‌کننده"),
    ("Seq_Quote", 550, "Number", 0, "شمارنده استعلام"),
    ("Seq_Quotation", 126, "Number", 0, "شمارنده پیشنهاد"),
]

LISTS = [
    # list name, values (comma-separated) or single entries below
    ("UoM", "pcs,kg,m,m2,liter,set,ton"),
    ("Currency", "IRR,USD,EUR"),
    ("Material_Type", "Raw Material,Purchased Part,Assembly,Sub Assembly,Packaging,Consumable"),
    ("Make_Buy", "Make,Buy"),
    ("Order_Type", "New,Repeat,Amendment"),
    ("Confidentiality", "Public,Internal,Confidential,Strictly Confidential"),
    ("Priority", "Low,Medium,High,Critical"),
    ("Decision", "APPROVED,APPROVED WITH CONDITION,RETURN FOR REVISION,REJECTED"),
    ("Stage", "Sales,Engineering,Planning,Procurement,Finance,Management"),
    ("Stage_Status", "Completed,In Progress,Pending,Not Started,Blocked"),
    ("Schedule_Status", "Not Started,In Progress,Done,Delayed"),
    ("Incoterm", "EXW,FCA,CPT,CIP,DAP,DDP"),
    ("Payment_Risk", "Low,Medium,High"),
    ("YesNo", "Yes,No"),
    ("Issue_Status", "Open,In Progress,Closed"),
    ("Quote_Confirm_Status", "Quoted,Confirmed,Revoked"),
    ("Price_Status", "Approved,Repriced,Expired"),
]

# =========================================================================== #
# v2.0 — Project lifecycle extensions
#   امکان‌سنجی، مرجع تجهیزات، تجهیزات پروژه، زمان‌سنجی/نفرساعت،
#   رسیدهای انبار، جمع‌بندی مدیریت پروژه، جدول گیت‌ها و کنترل دسترسی
# =========================================================================== #

FEASIBILITY = [
    # FEA, project, version, market, technical, economic, schedule, risk(0-100; بالاتر=پرریسک‌تر),
    # recommendation, status, date, by
    ("FEA-000125", "PRJ-2026-00125", "R01", 85, 90, 78, 80, 30,
     "اجرا شود؛ ریسک تأمین مکانیکال سیل با پیش‌خرید و تعهد کتبی تأمین‌کننده مدیریت شود",
     "Approved", "2026-09-22", "USR-03"),
    # پروژه دوم: امکان‌سنجی عمداً ناقص است تا قفل گیت‌ها (عدم اجازه ورود به مهندسی) نمایش داده شود
    ("FEA-000126", "PRJ-2026-00126", "R01", 70, None, None, 65, None,
     "", "Draft", "2026-09-22", "USR-03"),
]

EQUIPMENT = [
    # EQP, code, name, spec, uom, origin(Domestic/Imported), lead days, estimated price, currency, active
    ("EQP-000001", "EQP-01", "جرثقیل سقفی 5 تن", "DIN 15018 — دهانه 12 متر", "set", "Domestic", 45,
     8_500_000_000, "IRR", 1),
    ("EQP-000002", "EQP-02", "دستگاه جوشکاری صنعتی", "Fronius TPS 400i", "pcs", "Imported", 60,
     4_200, "USD", 1),
    ("EQP-000003", "EQP-03", "پرس هیدرولیک 100 تن", "قاب چهارستونه — 1000×1000", "pcs", "Domestic", 90,
     3_000_000_000, "IRR", 1),
    ("EQP-000004", "EQP-04", "تست‌بنچ پمپ (ادوات دقیق)", "فلومتر تا 300 m3/h + درایو VFD", "set", "Imported", 75,
     12_000, "USD", 1),
]

PROJECT_EQUIPMENT = [
    # PEQ, project, equipment, qty, need date, eng notes,
    # quote price (None = هنوز قیمت‌دهی نشده), currency, transport, installation, status, date, by
    ("PEQ-000001", "PRJ-2026-00125", "EQP-000004", 1, "2026-11-25",
     "تست عملکردی پمپ‌ها قبل از تحویل — شامل نصب در سایت تست",
     12_500, "USD", 150_000_000, 80_000_000, "Priced", "2026-10-02", "USR-06"),
    ("PEQ-000002", "PRJ-2026-00125", "EQP-000001", 1, "2026-11-20",
     "جابجایی پوسته‌ها در سالن مونتاژ",
     8_900_000_000, "IRR", 200_000_000, 120_000_000, "Priced", "2026-10-02", "USR-06"),
    # ردیف پروژه دوم: تا تکمیل زنجیره گیت‌ها (مهندسی) بازرگانی مجاز به قیمت‌دهی نیست
    ("PEQ-000003", "PRJ-2026-00126", "EQP-000002", 2, "2027-01-10",
     "جوشکاری بدنه — در انتظار تکمیل مهندسی",
     None, "", None, None, "Pending", "2026-09-23", "USR-04"),
]

TIME_STUDY = [
    # TST, project, product, operation, work center, setup min, std min/unit, operators, scrap allowance
    ("TST-000001", "PRJ-2026-00125", "PRD-000125", "برش و آماده‌سازی مواد اولیه", "WC-CUT", 240, 900, 2, 0.02),
    ("TST-000002", "PRJ-2026-00125", "PRD-000125", "ماشین‌کاری پوسته (CNC)", "WC-CNC", 480, 2400, 1, 0.02),
    ("TST-000003", "PRJ-2026-00125", "PRD-000125", "ماشین‌کاری پروانه و بالانس استاتیک", "WC-CNC", 240, 1200, 1, 0.02),
    ("TST-000004", "PRJ-2026-00125", "PRD-000125", "مونتاژ عمومی و بالانس دینامیک", "WC-ASM", 360, 2700, 2, 0.02),
    ("TST-000005", "PRJ-2026-00125", "PRD-000125", "تست عملکردی نهایی و کنترل کیفیت", "WC-QC", 60, 60, 1, 0.02),
]

RECEIPTS = [
    # RCV, planning line, receipt date, qty received, received by, QC status, date, by
    ("RCV-000001", "PLN-000001", "2026-10-18", 400, "USR-10", "Accepted", "2026-10-18", "USR-10"),
    ("RCV-000002", "PLN-000004", "2026-10-25", 2, "USR-10", "Accepted", "2026-10-25", "USR-10"),
]

PMO_SUMMARY = [
    # PMO, project, performance summary, gantt updated, feasibility updated, report date, date, by
    ("PMO-000125", "PRJ-2026-00125",
     "عملکرد واحدها مطابق برنامه: مهندسی و بازرگانی به‌موقع؛ برنامه‌ریزی تولید زمان‌بندی نیاز اقلام را صادر کرد؛ "
     "مالی قیمت تمام‌شده و بودجه راه‌اندازی را تدوین نمود. ریسک باقیمانده: تأمین مکانیکال سیل (پایش هفتگی).",
     "Yes", "Yes", "2026-10-06", "2026-10-06", "USR-03"),
    ("PMO-000126", "PRJ-2026-00126", "", "", "", None, "2026-09-23", "USR-03"),
]

# --------------------------------------------------------------------------- #
# Access control — نقش‌ها، مالکیت شیت‌ها و گذرواژه «محدوده‌های مجاز ویرایش»
# (در فایل نهایی به‌صورت Protected Range با گذرواژه اعمال می‌شود)
# --------------------------------------------------------------------------- #
ADMIN_PASSWORD = "راهبر1404"

ACCESS_ROLES = [
    # Role_ID, Role_Name, Department, Password, Scope (owned inputs), Notes
    ("ROL-02", "Sales (فروش)", "DEP-01", "فروش1404",
     "Orders (ثبت و بررسی سفارش) | Sales_Quotation (قیمت نهایی/حاشیه) | Payment_Terms (نحوه پرداخت)",
     "فقط ورودی‌های واحد فروش؛ مشاهده همه شیت‌ها آزاد است"),
    ("ROL-03", "Project Manager (مدیر پروژه)", "DEP-06", "مدیر1404",
     "Projects | Feasibility | Schedule | Gantt | PMO_Summary",
     "مشاهده همه بخش‌ها آزاد؛ ویرایش فقط در بخش‌های خود (تعریف پروژه، امکان‌سنجی، گانت، جمع‌بندی)"),
    ("ROL-04", "Engineering (مهندسی)", "DEP-02", "مهندسی1404",
     "Engineering | BOM_Header | BOM_Detail | Materials | Equipment | Project_Equipment (بخش مهندسی) | Time_Study",
     "تهیه لیست اقلام/تجهیزات، زمان‌سنجی و نفرساعت؛ نگهداری لیست‌های مرجع"),
    ("ROL-06", "Commercial (بازرگانی)", "DEP-04", "بازرگانی1404",
     "Supplier_Quotes | Supplier_Quote_Lines | Procurement | Purchase_Prices | Project_Equipment (بخش قیمت‌دهی)",
     "قیمت‌دهی اقلامِ فهرست مهندسی + نرخ ارز برای اقلام وارداتی"),
    ("ROL-05", "Planning (برنامه‌ریزی تولید و انبارها)", "DEP-03", "برنامه1404",
     "Planning (تاریخ نیاز/وضعیت تأمین) | Inventory",
     "مقدار مورد نیاز از مهندسی (BOM) محاسبه می‌شود؛ برنامه‌ریزی زمان‌بندی نیاز را ثبت می‌کند"),
    ("ROL-10", "Warehouse (انبار)", "DEP-03", "انبار1404",
     "Receipts (رسیدهای انبار) | Inventory (موجودی)",
     "ثبت رسید اقلام و نگهداری موجودی"),
    ("ROL-07", "Finance (مالی)", "DEP-05", "مالی1404",
     "Cost_Lines | Costing | Budget",
     "قیمت تمام‌شده محصول + تدوین بودجه راه‌اندازی"),
    ("ROL-08", "Senior Management (مدیریت ارشد)", "DEP-07", "ارشد1404",
     "Management_Approval (تصمیم نهایی: تأیید / رد / برگشت برای اصلاح)",
     "فقط ثبت تصمیم؛ سایر بخش‌ها صرفاً قابل مشاهده"),
    ("ROL-09", "Admin (راهبر)", "—", ADMIN_PASSWORD,
     "باز/بستن محافظت شیت‌ها و تغییر گذرواژه‌ها",
     "گذرواژه محافظت شیت‌ها؛ در اختیار مدیر سیستم"),
]
