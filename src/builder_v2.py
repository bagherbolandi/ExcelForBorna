"""
builder_v2.py — v2.0 lifecycle extensions
==========================================
Adds the full project lifecycle requested by the business owner:

  ۱. فروش: دریافت و بررسی سفارش                       (Orders — موجود)
  ۲. مدیر پروژه: تعریف پروژه + گانت + امکان‌سنجی      (Projects + Feasibility + Gantt)
  ۳. مهندسی: اقلام/تجهیزات/زمان‌سنجی/نفرساعت           (BOM + Equipment + Time_Study)
  ۴. بازرگانی: قیمت‌دهی اقلام فهرست مهندسی (+نرخ ارز)  (Purchase_Prices + Project_Equipment)
  ۵. برنامه‌ریزی تولید و انبارها: مقدار نیاز + زمان‌بندی (Planning + Receipts)
  ۶. مالی: قیمت تمام‌شده + بودجه راه‌اندازی            (Cost_Lines/Costing/Budget — موجود)
  ۷. فروش: قیمت نهایی + نحوه پرداخت                   (Sales_Quotation + Payment_Terms — موجود)
  ۸. مدیریت پروژه: جمع‌بندی + به‌روزرسانی گانت/امکان‌سنجی + اعلام به ارشد (PMO_Summary)
  ۹. مدیریت ارشد: تأیید / رد / برگشت برای اصلاح        (Management_Approval — موجود)

New sheets:
  Feasibility, Equipment, Project_Equipment, Time_Study, Receipts,
  PMO_Summary, Gate_Status, Gantt, Access_Control

Gate_Status is the heart of the "تقدم و تأخر ورود اطلاعات" (sequencing)
requirement: every stage flag is a live formula, cumulative gates are
products of the previous flags, and the renderer turns them into hard
Data-Validation locks on each unit's input cells.
"""
from __future__ import annotations

import datetime as dt

import seed
from builder import TB, D, F
from model import col_letter

# --------------------------------------------------------------------------- #
# Gantt horizon
# --------------------------------------------------------------------------- #
GANTT_HORIZON = dt.date(2026, 9, 21)   # اولین ستون دوره هفتگی
GANTT_WEEKS = 26                        # ۲۶ هفته ≈ ۶ ماه افق نمایش

# Gate_Status column letters (documented — renderer & DV depend on these):
#   A=System_ID B=Project_ID C=Order_ID
#   D..K = S1..S8 stage flags, L=Decision_Recorded, M=Decision_Text
#   N..U = cumulative gates G_PM..G_Senior
#   V=Stages_Completed, W=Completion_Percent, X=Current_Stage
GS_STAGE_COLS = ["D", "E", "F", "G", "H", "I", "J", "K", "L"]
GS_GATE_COLS = [None, "N", "O", "P", "Q", "R", "S", "T", "U"]  # gate required to enter stage i

STAGE_NAMES = [
    "بررسی سفارش (فروش)",
    "امکان‌سنجی (مدیر پروژه)",
    "مهندسی (اقلام/تجهیزات/زمان‌سنجی)",
    "بازرگانی (قیمت‌دهی اقلام)",
    "برنامه‌ریزی تولید و انبارها",
    "مالی (قیمت تمام‌شده و بودجه)",
    "قیمت‌گذاری فروش و نحوه پرداخت",
    "جمع‌بندی مدیریت پروژه",
    "تصمیم مدیریت ارشد",
]


# =========================================================================== #
# Feasibility — امکان‌سنجی (مالک: مدیر پروژه)
# =========================================================================== #
def _build_feasibility(wb, SET):
    h = ["System_ID", "Project_ID", "Version", "Market_Score", "Technical_Score",
         "Economic_Score", "Schedule_Score", "Risk_Score", "Weighted_Score", "Risk_Level",
         "Feasibility_Result", "Go_NoGo_Recommendation", "Completed_Flag", "Status_ID",
         "Create_Date", "Created_By"]
    t = TB(wb, "Feasibility", "tblFeasibility", h,
           note="امکان‌سنجی توسط مدیر پروژه تکمیل می‌شود؛ بدون Completed_Flag=COMPLETE گیت مهندسی باز نمی‌شود. "
                "وزن‌ها و کف امتیاز از Settings خوانده می‌شود.")
    wm, wt, we, ws_ = (SET["Feasibility_Weight_Market"], SET["Feasibility_Weight_Technical"],
                        SET["Feasibility_Weight_Economic"], SET["Feasibility_Weight_Schedule"])
    minscore = SET["Feasibility_Min_Score"]
    for f in seed.FEASIBILITY:
        r = t.row + 1
        t.add([
            f[0], f[1], f[2], f[3], f[4], f[5], f[6], f[7],
            F(f'=IF(AND(ISNUMBER($D{r}),ISNUMBER($E{r}),ISNUMBER($F{r}),ISNUMBER($G{r})),'
              f'ROUND({wm}*$D{r}+{wt}*$E{r}+{we}*$F{r}+{ws_}*$G{r},2),"")'),
            F(f'=IF(ISNUMBER($H{r}),IF($H{r}>=67,"High",IF($H{r}>=34,"Medium","Low")),"")'),
            F(f'=IF($I{r}="","",IF(AND($I{r}>={minscore},$J{r}<>"High"),"Feasible",'
              f'IF($I{r}>=50,"Conditional","Not Feasible")))'),
            f[8],
            F(f'=IF(AND(ISNUMBER($D{r}),ISNUMBER($E{r}),ISNUMBER($F{r}),ISNUMBER($G{r}),'
              f'ISNUMBER($H{r}),$L{r}<>""),"COMPLETE","INCOMPLETE")'),
            f[9], D(f[10]), f[11],
        ])
    t.finish()


# =========================================================================== #
# Equipment master — لیست مرجع تجهیزات (مالک: مهندسی)
# =========================================================================== #
def _build_equipment(wb):
    h = ["System_ID", "Equipment_Code", "Equipment_Name", "Specification", "UoM", "Origin",
         "Lead_Time_Days", "Estimated_Price", "Currency", "Is_Active"]
    t = TB(wb, "Equipment", "tblEquipment", h,
           note="لیست مرجع تجهیزات؛ مهندسی از این فهرست در Project_Equipment انتخاب می‌کند و بازرگانی همان ردیف‌ها را قیمت می‌دهد. "
                "برای تجهیزات وارداتی نرخ ارز از Settings اعمال می‌شود.")
    for e in seed.EQUIPMENT:
        t.add(list(e))
    t.finish()


# =========================================================================== #
# Project_Equipment — تجهیزات پروژه: انتخاب با مهندسی، قیمت‌دهی با بازرگانی
# =========================================================================== #
def _build_project_equipment(wb, SET):
    h = ["System_ID", "Project_ID", "Equipment_ID", "Qty", "Need_Date", "Engineering_Notes",
         "Unit_Price_Quote", "Currency", "Fx_Rate", "Unit_Price_IRR", "Transport_Cost",
         "Installation_Cost", "Total_Cost_IRR", "Priced_Flag", "Status_ID", "Create_Date", "Created_By"]
    t = TB(wb, "Project_Equipment", "tblProjectEquipment", h,
           note="ستون‌های آبی (B:F) ورودی مهندسی — انتخاب از لیست مرجع تجهیزات؛ ستون‌های قیمت (G:K) فقط ورودی بازرگانی پس از باز شدن گیت بازرگانی. "
                "قیمت وارداتی = نرخ ارزی × نرخ ارز (Settings)؛ قیمت نهایی شامل حمل و نصب.")
    fx = SET["Fx_USD_TO_IRR"]
    for p in seed.PROJECT_EQUIPMENT:
        r = t.row + 1
        t.add([
            p[0], p[1], p[2], p[3], D(p[4]), p[5],
            p[6], p[7],
            F(f'=IF($H{r}="IRR",1,{fx})'),
            F(f'=IF($G{r}="","",ROUND($G{r}*$I{r},0))'),
            p[8], p[9],
            F(f'=IF($G{r}="","",ROUND($G{r}*$I{r}*$D{r}'
              f'+IF(ISNUMBER($K{r}),$K{r},0)+IF(ISNUMBER($L{r}),$L{r},0),0))'),
            F(f'=IF($G{r}="","PENDING","PRICED")'),
            p[10], D(p[11]), p[12],
        ])
    t.finish()


# =========================================================================== #
# Time_Study — زمان‌سنجی و نفرساعت (مالک: مهندسی) → مصرف در دستمزد Costing
# =========================================================================== #
def _build_time_study(wb):
    h = ["System_ID", "Project_ID", "Product_ID", "Operation", "Work_Center", "Setup_Min",
         "Std_Min_Per_Unit", "Operators", "Scrap_Allowance_%", "Order_Qty", "Man_Min_Total",
         "Man_Hours_Total", "Man_Hours_Per_Unit", "Status_ID", "Create_Date", "Created_By"]
    t = TB(wb, "Time_Study", "tblTimeStudy", h,
           note="زمان‌سنجی عملیات توسط مهندسی؛ نفرساعت = (آماده‌سازی + زمان استاندارد×مقدار سفارش×(1+ضایعات))/60 × تعداد اپراتور. "
                "جمع نفرساعت مستقیماً وارد دستمزد مستقیم در Cost_Lines می‌شود.")
    for s in seed.TIME_STUDY:
        r = t.row + 1
        t.add([
            s[0], s[1], s[2], s[3], s[4], s[5], s[6], s[7], s[8],
            F(f'=SUMIFS(Order_Lines!$E$3:$E$20,Order_Lines!$B$3:$B$20,'
              f'XLOOKUP($B{r},Projects!$A$3:$A$20,Projects!$C$3:$C$20,""),'
              f'Order_Lines!$D$3:$D$20,$C{r})'),
            F(f'=IF($J{r}=0,0,$F{r}+$G{r}*$J{r}*(1+$I{r}))'),
            F(f'=ROUND($K{r}/60*$H{r},2)'),
            F(f'=IF($J{r}=0,0,ROUND($L{r}/$J{r},2))'),
            "Approved", D("2026-09-27"), "USR-04",
        ])
    t.finish()


# =========================================================================== #
# Receipts — رسیدهای انبار (مالک: انبار)
# =========================================================================== #
def _build_receipts(wb):
    h = ["System_ID", "Planning_ID", "Project_ID", "Material_ID", "Receipt_Date",
         "Qty_Received", "Received_By", "QC_Status", "Create_Date", "Created_By"]
    t = TB(wb, "Receipts", "tblReceipts", h,
           note="ثبت رسید اقلام توسط انبار با ارجاع به ردیف برنامه‌ریزی (PLN). پروژه/آیتم به‌صورت خودکار از Planning خوانده می‌شود.")
    for rc in seed.RECEIPTS:
        r = t.row + 1
        t.add([
            rc[0], rc[1],
            F(f'=IFERROR(XLOOKUP($B{r},Planning!$A$3:$A$30,Planning!$B$3:$B$30,""),"")'),
            F(f'=IFERROR(XLOOKUP($B{r},Planning!$A$3:$A$30,Planning!$F$3:$F$30,""),"")'),
            D(rc[2]), rc[3], rc[4], rc[5], D(rc[6]), rc[7],
        ])
    t.finish()


# =========================================================================== #
# PMO_Summary — جمع‌بندی مدیریت پروژه و اعلام به مدیریت ارشد
# =========================================================================== #
def _build_pmo_summary(wb):
    h = ["System_ID", "Project_ID", "Units_Completed", "Completion_Percent", "Performance_Summary",
         "Gantt_Updated_Flag", "Feasibility_Updated_Flag", "Report_Date", "Decision",
         "Reported_Flag", "Next_Action", "Create_Date", "Created_By"]
    t = TB(wb, "PMO_Summary", "tblPmoSummary", h,
           note="مدیر پروژه: جمع‌بندی عملکرد واحدها + به‌روزرسانی گانت/امکان‌سنجی + تاریخ اعلام به ارشد. "
                "تصمیم ستون I از Management_Approval خوانده می‌شود (فقط مشاهده).")
    for p in seed.PMO_SUMMARY:
        r = t.row + 1
        t.add([
            p[0], p[1],
            F(f'=IFERROR(XLOOKUP($B{r},Gate_Status!$B$3:$B$20,Gate_Status!$V$3:$V$20),0)'),
            F(f'=ROUND($C{r}/8,3)'),
            p[2], p[3], p[4], D(p[5]),
            F(f'=IFERROR(XLOOKUP($B{r},Management_Approval!$C$3:$C$10,Management_Approval!$D$3:$D$10,""),"")'),
            F(f'=IF($H{r}<>"","REPORTED","NOT REPORTED")'),
            F(f'=IF($I{r}="APPROVED","پروژه مصوب — شروع اجرا",'
              f'IF($I{r}="APPROVED WITH CONDITION","مصوب مشروط — پیگیری شروط",'
              f'IF($I{r}="RETURN FOR REVISION","اصلاح و ارسال مجدد",'
              f'IF($I{r}="REJECTED","بایگانی/خاتمه","در انتظار تصمیم مدیریت ارشد"))))'),
            D(p[6]), p[7],
        ])
    t.finish()


# =========================================================================== #
# Gate_Status — وضعیت گیت‌ها (قلب تقدم و تأخر ورود اطلاعات)
# =========================================================================== #
def _build_gate_status(wb):
    h = ["System_ID", "Project_ID", "Order_ID",
         "S1_Sales_Ready", "S2_PM_Feasibility", "S3_Engineering", "S4_Commercial",
         "S5_Planning_Warehouse", "S6_Finance", "S7_Sales_Pricing", "S8_PMO_Summary",
         "Decision_Recorded", "Decision_Text",
         "G_PM", "G_Eng", "G_Com", "G_Pln", "G_Fin", "G_Price", "G_PMO", "G_Senior",
         "Stages_Completed", "Completion_Percent", "Current_Stage"]
    t = TB(wb, "Gate_Status", "tblGateStatus", h,
           note="هر ستون S پرچم تکمیل یک مرحله (فرمول زنده)؛ ستون‌های G گیت تجمعی هستند: تا مرحله قبل تکمیل نشود، "
                "ورود اطلاعات در مرحله بعد توسط Data Validation مسدود است. Any_G_* ها در ردیف ۱ برای گیت شیت‌های غیرپروژه‌محور.")
    ws = t.ws
    # workbook-level aggregates (used by DV on material-level sheets)
    ws.set(1, 26, value="Any_G_Eng")
    ws.set(1, 27, formula="=MAX($O$3:$O$20)")
    ws.set(1, 28, value="Any_G_Com")
    ws.set(1, 29, formula="=MAX($P$3:$P$20)")
    ws.set(1, 30, value="Any_G_Pln")
    ws.set(1, 31, formula="=MAX($Q$3:$Q$20)")

    for i, p in enumerate(seed.PROJECTS):
        r = t.row + 1
        pr = 3 + i  # corresponding Projects row
        t.add([
            f"GST-{i+1:06d}",
            F(f"=Projects!$A${pr}"),
            F(f"=Projects!$C${pr}"),
            # S1 — سفارش آماده (فروش)
            F(f'=IF(COUNTIFS(Orders!$A$3:$A$20,$C{r},Orders!$R$3:$R$20,"READY")>0,1,0)'),
            # S2 — امکان‌سنجی کامل (مدیر پروژه)
            F(f'=IF(COUNTIFS(Feasibility!$B$3:$B$20,$B{r},Feasibility!$M$3:$M$20,"COMPLETE")>0,1,0)'),
            # S3 — مهندسی: رکورد تأییدشده + زمان‌سنجی + نسخه تأییدشده BOM برای محصول پروژه
            F(f'=IF(AND(COUNTIFS(Engineering!$B$3:$B$20,$B{r},Engineering!$H$3:$H$20,"Approved")>0,'
              f'COUNTIFS(Time_Study!$B$3:$B$20,$B{r})>0,'
              f'COUNTIFS(BOM_Header!$C$3:$C$20,'
              f'XLOOKUP(XLOOKUP($B{r},Projects!$A$3:$A$20,Projects!$E$3:$E$20,""),'
              f'Product_Revisions!$B$3:$B$20,Product_Revisions!$A$3:$A$20,""),'
              f'BOM_Header!$I$3:$I$20,"Approved")>0),1,0)'),
            # S4 — بازرگانی: همه تجهیزات پروژه قیمت‌گذاری‌شده + حداقل یک قیمت خرید معتبر تأییدشده
            F(f'=IF(AND(COUNTIFS(Project_Equipment!$B$3:$B$20,$B{r})>0,'
              f'COUNTIFS(Project_Equipment!$B$3:$B$20,$B{r},Project_Equipment!$N$3:$N$20,"PRICED")'
              f'=COUNTIFS(Project_Equipment!$B$3:$B$20,$B{r}),'
              f'COUNTIFS(Purchase_Prices!$S$3:$S$20,"Approved",Purchase_Prices!$R$3:$R$20,0)>0),1,0)'),
            # S5 — برنامه‌ریزی/انبار: همه ردیف‌های نیاز پروژه تأیید شده
            F(f'=IF(AND(COUNTIFS(Planning!$B$3:$B$30,$B{r})>0,'
              f'COUNTIFS(Planning!$B$3:$B$30,$B{r},Planning!$R$3:$R$30,"Approved")'
              f'=COUNTIFS(Planning!$B$3:$B$30,$B{r})),1,0)'),
            # S6 — مالی: Costing تأییدشده برای پروژه
            F(f'=IF(COUNTIFS(Costing!$B$3:$B$10,$B{r},Costing!$D$3:$D$10,"Approved")>0,1,0)'),
            # S7 — فروش: پیشنهاد قیمت Submitted
            F(f'=IF(COUNTIFS(Sales_Quotation!$B$3:$B$10,$B{r},Sales_Quotation!$O$3:$O$10,"Submitted")>0,1,0)'),
            # S8 — جمع‌بندی مدیریت پروژه گزارش‌شده
            F(f'=IF(COUNTIFS(PMO_Summary!$B$3:$B$10,$B{r},PMO_Summary!$J$3:$J$10,"REPORTED")>0,1,0)'),
            # تصمیم مدیریت ارشد
            F(f'=IF($M{r}<>"",1,0)'),
            F(f'=IFERROR(XLOOKUP($B{r},Management_Approval!$C$3:$C$10,Management_Approval!$D$3:$D$10,""),"")'),
            # gates (cumulative)
            F(f"=$D{r}"),
            F(f"=$D{r}*$E{r}"),
            F(f"=$O{r}*$F{r}"),
            F(f"=$P{r}*$G{r}"),
            F(f"=$Q{r}*$H{r}"),
            F(f"=$R{r}*$I{r}"),
            F(f"=$S{r}*$J{r}"),
            F(f"=$T{r}*$K{r}"),
            F(f"=SUM($D{r}:$K{r})"),
            F(f"=ROUND($V{r}/8,3)"),
            F(f'=IF($L{r}=1,"چرخه کامل — تصمیم ثبت شده",'
              f'IF($D{r}=0,"بررسی سفارش (فروش)",'
              f'IF($E{r}=0,"تعریف پروژه و امکان‌سنجی (مدیر پروژه)",'
              f'IF($F{r}=0,"مهندسی (اقلام/تجهیزات/زمان‌سنجی)",'
              f'IF($G{r}=0,"بازرگانی (قیمت‌دهی اقلام)",'
              f'IF($H{r}=0,"برنامه‌ریزی تولید و انبارها",'
              f'IF($I{r}=0,"مالی (قیمت تمام‌شده و بودجه)",'
              f'IF($J{r}=0,"فروش (قیمت نهایی و نحوه پرداخت)",'
              f'IF($K{r}=0,"جمع‌بندی و اعلام به مدیریت ارشد",'
              f'"در انتظار تصمیم مدیریت ارشد")))))))))'),
        ])
    t.finish()


# =========================================================================== #
# Gantt — نمودار گانت زمان‌بندی/عملکرد پروژه‌ها (مالک ورودی‌ها: مدیر پروژه از طریق Schedule)
# =========================================================================== #
def _build_gantt(wb):
    ws = wb.sheet("Gantt")
    ws.set(1, 1, value="Gantt — نمودار گانت برنامه/عملکرد (منبع: Schedule؛ به‌روزرسانی توسط مدیر پروژه)")
    ws.set(1, 2, value="هر ستون یک هفته؛ میله = بازه شروع تا پایان فعالیت. رنگ‌ها: سبز=انجام‌شده، آبی=در جریان، قرمز=معوق، خاکستری=برنامه‌ریزی‌شده.")
    headers = ["System_ID", "Project_ID", "Activity", "Responsible", "Status_ID",
               "Start_Date", "End_Date", "Duration", "Display_Status"]
    for c, htxt in enumerate(headers, start=1):
        ws.set(2, c, value=htxt)
    # week period headers
    first_col = len(headers) + 1
    for k in range(GANTT_WEEKS):
        ws.set(2, first_col + k, value=GANTT_HORIZON + dt.timedelta(days=7 * k))

    n = len(seed.SCHEDULE) + 1  # builder._build_schedule adds one extra delayed row
    for i in range(n):
        r = 3 + i
        sr = 3 + i  # Schedule row
        ws.set(r, 1, formula=f"=Schedule!$A${sr}")
        ws.set(r, 2, formula=f"=Schedule!$B${sr}")
        ws.set(r, 3, formula=f"=Schedule!$C${sr}")
        ws.set(r, 4, formula=f"=Schedule!$G${sr}")
        ws.set(r, 5, formula=f"=Schedule!$I${sr}")
        ws.set(r, 6, formula=f"=Schedule!$D${sr}")
        ws.set(r, 7, formula=f"=Schedule!$E${sr}")
        ws.set(r, 8, formula=f'=IF(OR($F{r}="",$G{r}=""),0,$G{r}-$F{r}+1)')
        ws.set(r, 9, formula=(f'=IF($F{r}="","",IF($E{r}="Done","Done",'
                              f'IF($G{r}<TODAY(),"Late",'
                              f'IF($E{r}="Not Started","Planned","In Progress"))))'))
        for k in range(GANTT_WEEKS):
            cl = col_letter(first_col + k)
            ws.set(r, first_col + k,
                   formula=f'=IF(AND($F{r}<>"",{cl}$2>=$F{r},{cl}$2<=$G{r}),1,0)')
    ws.col_widths[3] = 26
    ws.freeze = "F3"
    return ws


# =========================================================================== #
# Access_Control — نقش‌ها و گذرواژه محدوده‌های مجاز ویرایش
# =========================================================================== #
def _build_access_control(wb):
    h = ["Role_ID", "Role_Name", "Department_ID", "Edit_Password", "Scope", "Notes"]
    t = TB(wb, "Access_Control", "tblAccessControl", h,
           note="هر نقش فقط محدوده‌های خودش را (با گذرواژه خود) می‌تواند ویرایش کند؛ مشاهده همه شیت‌ها آزاد است. "
                "اعمال فنی: Sheet Protection + Protected Ranges. برای تغییر گذرواژه‌ها، اسکریپت ساخت را ویرایش و بازتولید کنید.")
    for a in seed.ACCESS_ROLES:
        t.add([a[0], a[1], a[2], a[3], a[4], a[5]])
    t.finish()


# =========================================================================== #
# Dashboard v2 — داشبورد مدیریتی: KPI + قیف + ماتریس گردش کار + داده نمودار
# =========================================================================== #
def build_dashboard_v2(wb):
    ws = wb.sheet("Dashboard")
    ws.set(1, 1, value="Dashboard — داشبورد مدیریت فرآیند «سفارش تا تصویب» (KPI + ماتریس گردش کار + گیت‌ها)")
    n_ord = len(seed.ORDERS)
    n_proj = len(seed.PROJECTS)
    of, ol = 3, 3 + n_ord - 1
    pf, pl = 3, 3 + n_proj - 1

    def kpi(label, formula, r):
        ws.set(r, 1, value=label)
        ws.set(r, 2, formula=formula)
        return r + 1

    r = 3
    r = kpi("پروژه‌های فعال", f'=COUNTIF(Projects!$K${pf}:$K${pl},"Approved")', r)
    r = kpi("سفارش‌های جدید", f'=COUNTA(Orders!$A${of}:$A${ol})', r)
    r = kpi("در انتظار بررسی فروش", f'=COUNTIF(Orders!$M${of}:$M${ol},"Draft")', r)
    r = kpi("پروژه‌های تأخیرخورده", f'=SUM(Projects!$N${pf}:$N${pl})', r)
    r = kpi("امکان‌سنجی‌های کامل", '=COUNTIF(Feasibility!$M$3:$M$20,"COMPLETE")', r)
    r = kpi("امکان‌سنجی‌های ناقص (قفل مهندسی)", '=COUNTIF(Feasibility!$M$3:$M$20,"INCOMPLETE")', r)
    r = kpi("تجهیزات قیمت‌دهی‌شده", '=COUNTIF(Project_Equipment!$N$3:$N$20,"PRICED")', r)
    r = kpi("تجهیزات در انتظار قیمت‌دهی", '=COUNTIF(Project_Equipment!$N$3:$N$20,"PENDING")', r)
    r = kpi("اقلام دریافت‌شده در انبار (تعداد)", '=SUM(Receipts!$F$3:$F$20)', r)
    r = kpi("BOMهای در انتظار تأیید", '=COUNTIF(BOM_Header!$I$3:$I$5,"Waiting Approval")', r)
    r = kpi("قیمت‌های منقضی", '=COUNTIF(Purchase_Prices!$R$3:$R$10,1)', r)
    r = kpi("پیشنهاد در انتظار مدیریت", '=COUNTIF(Sales_Quotation!$O$3:$O$5,"Submitted")', r)
    r = kpi("تصمیم تأیید (مطلق)", '=COUNTIF(Management_Approval!$D$3:$D$10,"APPROVED")', r)
    r = kpi("تصمیم تأیید مشروط", '=COUNTIF(Management_Approval!$D$3:$D$10,"APPROVED WITH CONDITION")', r)
    r = kpi("تصمیم برگشت برای اصلاح", '=COUNTIF(Management_Approval!$D$3:$D$10,"RETURN FOR REVISION")', r)
    r = kpi("تصمیم رد", '=COUNTIF(Management_Approval!$D$3:$D$10,"REJECTED")', r)
    r = kpi("Total Sales Value (پیشنهاد جاری)", '=Quotation_Versions!$G$3', r)
    r = kpi("Total Cost", '=SUM(Costing!$Q$3:$Q$3)', r)
    r = kpi("اختلاف بودجه و هزینه واقعی", '=Budget!$S$3', r)

    r += 1
    ws.set(r, 1, value="Funnel (تعداد)")
    funnel = [
        ("Leads / سفارش‌ها", f'=COUNTA(Orders!$A${of}:$A${ol})'),
        ("Orders → READY", f'=COUNTIF(Orders!$R${of}:$R${ol},"READY")'),
        ("امکان‌سنجی کامل", '=COUNTIF(Feasibility!$M$3:$M$20,"COMPLETE")'),
        ("مهندسی تأییدشده", '=COUNTIF(Engineering!$H$3:$H$4,"Approved")'),
        ("Costing تأییدشده", '=COUNTIF(Costing!$D$3:$D$3,"Approved")'),
        ("پیشنهاد Submitted", '=COUNTIF(Sales_Quotation!$O$3:$O$5,"Submitted")'),
        ("تصمیم مدیریت (کل)", '=COUNTA(Management_Approval!$D$3:$D$10)'),
    ]
    for fname, fform in funnel:
        r += 1
        ws.set(r, 1, value=fname)
        ws.set(r, 2, formula=fform)

    # ------------------------------------------------------------------ #
    # Workflow matrix — وضعیت هر پروژه در ۹ مرحله (از Gate_Status)
    # ------------------------------------------------------------------ #
    r += 2
    matrix_title_row = r
    ws.set(r, 1, value="ماتریس گردش کار پروژه‌ها — ✅ تکمیل / 🔵 در جریان / 🔒 قفل (گیت بسته)")
    r += 1
    header_row = r
    ws.set(r, 1, value="پروژه")
    for si, name in enumerate(STAGE_NAMES):
        ws.set(r, 2 + si, value=name)
    ws.set(r, 2 + len(STAGE_NAMES), value="مراحل تکمیل‌شده")
    ws.set(r, 3 + len(STAGE_NAMES), value="درصد پیشروی")

    def match_proj(rr):
        return f"MATCH($A{rr},Gate_Status!$B$3:$B$20,0)"

    for pi, proj in enumerate(seed.PROJECTS):
        rr = r + 1 + pi
        ws.set(rr, 1, value=proj[0])
        for si in range(9):
            fc = GS_STAGE_COLS[si]
            gc = GS_GATE_COLS[si]
            colidx = 2 + si
            if gc is None:
                ws.set(rr, colidx, formula=(
                    f'=IFERROR(IF(INDEX(Gate_Status!${fc}$3:${fc}$20,{match_proj(rr)})=1,'
                    f'"✅ تکمیل","🔵 در جریان"),"—")'))
            else:
                ws.set(rr, colidx, formula=(
                    f'=IFERROR(IF(INDEX(Gate_Status!${fc}$3:${fc}$20,{match_proj(rr)})=1,'
                    f'"✅ تکمیل",IF(INDEX(Gate_Status!${gc}$3:${gc}$20,{match_proj(rr)})=1,'
                    f'"🔵 در جریان","🔒 قفل")),"—")'))
        last = 2 + len(STAGE_NAMES)
        ws.set(rr, last, formula=(
            f'=IFERROR(INDEX(Gate_Status!$V$3:$V$20,{match_proj(rr)}),"")'))
        ws.set(rr, last + 1, formula=(
            f'=IFERROR(ROUND(INDEX(Gate_Status!$W$3:$W$20,{match_proj(rr)}),3),"")'))

    # ------------------------------------------------------------------ #
    # Chart source data — stages completed counts
    # ------------------------------------------------------------------ #
    r = r + 1 + len(seed.PROJECTS) + 1
    chart_title_row = r
    ws.set(r, 1, value="داده نمودار — تعداد پروژه‌های تکمیل‌کننده هر مرحله")
    r += 1
    chart_first = r
    stage_cols = ["D", "E", "F", "G", "H", "I", "J", "K", "L"]
    for si in range(9):
        ws.set(r + si, 1, value=STAGE_NAMES[si])
        ws.set(r + si, 2, formula=f"=SUM(Gate_Status!${stage_cols[si]}$3:${stage_cols[si]}$20)")
    chart_last = r + 8
    ws.set(1, 26, value=f"CHART1_ROWS={chart_first}:{chart_last}")  # machine marker for renderer

    ws.col_widths[1] = 38
    ws.col_widths[2] = 22
    for c in range(3, 13):
        ws.col_widths[c] = 16
    return ws, chart_first, chart_last


# =========================================================================== #
# Workflow sheet v2 — نقشه ۹ مرحله‌ای با مالک و گیت
# =========================================================================== #
def build_workflow_v2(wb):
    ws = wb.sheet("Workflow")
    ws.set(1, 1, value="Workflow — گردش کار ۹ مرحله‌ای «سفارش تا تصویب مدیریت ارشد» + گیت‌ها و مالکان")
    ws.set(2, 1, value="№")
    ws.set(2, 2, value="مرحله")
    ws.set(2, 3, value="مالک (واحد)")
    ws.set(2, 4, value="گیت ورود (پیش‌نیاز)")
    ws.set(2, 5, value="شرط عبور به مرحله بعد")
    stages = [
        ("۱", "دریافت و بررسی سفارش", "فروش", "—", "سفارش کامل و تأیید اولیه (Ready_Flag=READY)"),
        ("۲", "تعریف پروژه + تهیه گانت + تکمیل امکان‌سنجی", "مدیر پروژه", "گیت ۱ (سفارش آماده)",
         "امکان‌سنجی کامل (امتیازها + نتیجه + توصیه)"),
        ("۳", "تهیه اقلام (BOM)، تجهیزات، زمان‌سنجی و نفرساعت", "مهندسی", "گیت ۲ (امکان‌سنجی کامل)",
         "مهندسی تأییدشده + BOM تأییدشده + زمان‌سنجی ثبت‌شده"),
        ("۴", "تعیین قیمت اقلام و تجهیزات فهرست مهندسی (با نرخ ارز وارداتی)", "بازرگانی",
         "گیت ۳ (مهندسی کامل)", "همه اقلام/تجهیزات پروژه قیمت‌گذاری‌شده و معتبر"),
        ("۵", "مقدار مورد نیاز اقلام برای تولید + زمان‌بندی نیاز + ثبت رسید انبار",
         "برنامه‌ریزی تولید و انبارها", "گیت ۴ (بازرگانی کامل)",
         "تاریخ نیاز و وضعیت تأمین همه ردیف‌های نیاز ثبت‌شده"),
        ("۶", "قیمت تمام‌شده محصول + تدوین بودجه راه‌اندازی", "مالی", "گیت ۵ (برنامه‌ریزی کامل)",
         "Costing تأییدشده + بودجه ثبت‌شده"),
        ("۷", "تعیین قیمت نهایی فروش و نحوه پرداخت", "فروش", "گیت ۶ (مالی کامل)",
         "پیشنهاد قیمت Submitted با سناریو و شرایط پرداخت"),
        ("۸", "جمع‌بندی عملکرد واحدها + به‌روزرسانی گانت/امکان‌سنجی + اعلام به ارشد",
         "مدیریت پروژه (PMO)", "گیت ۷ (قیمت‌گذاری کامل)", "گزارش جمع‌بندی ارسال‌شده (Report_Date)"),
        ("۹", "تأیید / رد / برگشت برای اصلاح", "مدیریت ارشد", "گیت ۸ (گزارش کامل)",
         "ثبت تصمیم چهارگانه: APPROVED / مشروط / برگشت / رد"),
    ]
    for i, row in enumerate(stages):
        rr = 3 + i
        for c, v in enumerate(row, start=1):
            ws.set(rr, c, value=v)
        ws.set(rr, 6, value="↓" if i < len(stages) - 1 else "")
    rr = 3 + len(stages) + 1
    ws.set(rr, 1, value="برگشت برای اصلاح: تصمیم RETURN FOR REVISION پروژه را به مرحله مربوط بازمی‌گرداند؛ تغییر هر ورودیِ تصویب‌شده فقط با Revision جدید.")
    ws.set(rr + 1, 1, value="اجرای گیت‌ها: ستون‌های G در شیت Gate_Status + Data Validation روی سلول‌های ورودی + محدوده‌های ویرایش رمزدار (شیت Access_Control).")
    ws.col_widths[2] = 62
    ws.col_widths[3] = 26
    ws.col_widths[4] = 26
    ws.col_widths[5] = 52
    return ws


# =========================================================================== #
# Sheet ordering — workflow-first tab layout
# =========================================================================== #
DESIRED_ORDER = [
    "Dashboard", "Workflow", "Gate_Status", "Access_Control",
    "Orders", "Order_Lines", "Projects", "Feasibility", "Gantt", "Schedule",
    "Engineering", "BOM_Header", "BOM_Detail", "Time_Study",
    "Materials", "Equipment", "Project_Equipment",
    "Supplier_Quotes", "Supplier_Quote_Lines", "Procurement", "Purchase_Prices",
    "Planning", "Receipts", "Inventory",
    "Cost_Lines", "Costing", "Budget",
    "Payment_Terms", "Sales_Quotation", "Quotation_Versions",
    "PMO_Summary", "Management_Approval", "Project_Statuses",
    "Customers", "Products", "Product_Revisions", "Suppliers",
    "Settings", "Lists", "Statuses", "Departments", "Roles", "Users",
    "Approval_Matrix", "RACI",
    "Workflow_History", "Change_Log", "Issues", "Error_Checks", "Test_Results",
    "_Trace", "_NamedRanges", "Report_Quotation",
]


def reorder(wb):
    missing = [s for s in DESIRED_ORDER if not wb.has(s)]
    extra = [s for s in wb.order if s not in DESIRED_ORDER]
    if missing or extra:
        raise RuntimeError(f"sheet order mismatch: missing={missing} extra={extra}")
    wb.order = list(DESIRED_ORDER)


# =========================================================================== #
# Entry point
# =========================================================================== #
def extend(wb, SET):
    _build_feasibility(wb, SET)
    _build_equipment(wb)
    _build_project_equipment(wb, SET)
    _build_time_study(wb)
    _build_receipts(wb)
    _build_pmo_summary(wb)
    _build_gate_status(wb)
    _build_gantt(wb)
    _build_access_control(wb)
    # NOTE: reorder() is called from builder.build_model AFTER the v2
    # Workflow/Dashboard sheets are created.
