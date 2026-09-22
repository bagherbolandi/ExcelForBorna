# 05 — Data Dictionary (فرهنگ داده) — مطابق نسخه ساخته‌شده
نسخه: 1.1
> این سند با اسکیمای واقعی Workbook تولید شده است (44 شیت، نام جدول‌ها و ستون عینی).

## قرارداد
- هر جدول Transaction ستون‌های پایه: System_ID (کلید)، Business_Code (در هدرها)، Revision، Status_ID، Create_Date، Created_By، Modify/By (در موارد لازم).
- ستون‌های محاسباتی با «(calc)» مشخص‌اند.

## جداول Reference / Master
| جدول | ستون‌ها |
|---|---|
| Settings | Key, Value, Data_Type, ASSUMPTION, Description |
| Lists | List_Name, Value |
| Statuses | Status_ID, Status_Name, Status_Group, Sequence |
| Departments | Department_ID, Department_Name |
| Roles | Role_ID, Role_Name |
| Users | User_ID, Name, Department_ID, Role_ID, Email, Approval_Level, Is_Active |
| Approval_Matrix | Entity, Gate, Approver_Role, Min_Approval_Level, Responsible, Reviewer, Approver, Informed |
| RACI | Activity, Sales, Engineering, Planning, Procurement, Finance, PMO, Management |
| Customers | System_ID, Customer_Code, Customer_Name, Legal_Name, Customer_Type, Country, City, Delivery_Location, Payment_Condition, Credit_Limit, Currency, Tax_Status, Is_Active, Sales_Owner, Notes |
| Products | System_ID, Product_Code, Product_Name, Product_Type, Base_UoM, Category, Is_Active |
| Product_Revisions | System_ID, Product_ID, Revision_No, Drawing_No, Drawing_Revision, BOM_Policy, Approved_By, Approve_Date, Change_Reason |
| Materials | System_ID, Material_Code, Material_Description, Material_Type, Specification, UoM, Make_Buy_Default, Preferred_Supplier, Is_Active |
| Suppliers | System_ID, Supplier_Code, Supplier_Name, Legal_Name, Country, Currency, Payment_Terms_Default, Incoterm_Default, Rating, Approval_Status, Contact, Is_Active |
| Inventory | System_ID, Material_ID, Stock_On_Hand, Reserved |

## جداول Transaction
| جدول | ستون‌ها |
|---|---|
| Orders | System_ID, Customer_ID, Order_Date, Requested_Delivery_Date, Request_Type, Total_Qty(calc), Delivery_Terms, Payment_Terms, Validity_Days, Confidentiality_Level, Sales_Responsible, Doc_Ref, Status_ID, CO_01_Customer_Valid(calc), CO_02_Qty_Defined(calc), CO_03_Delivery_Defined(calc), CO_04_Payment_Defined(calc), Ready_Flag(calc), Duplicate_Flag(calc), Create_Date, Created_By |
| Order_Lines | System_ID, Order_ID, Line_No, Product_ID, Quantity, UoM, Latest_FLAG, Create_Date, Created_By |
| Projects | System_ID, Project_Code, Order_ID, Customer_ID, Product_ID, Project_Name, Project_Manager, Start_Date, Due_Date, Priority, Status_ID, Progress_%, Sponsor, Team_List, Overdue_Flag(calc), Create_Date, Created_By |
| Project_Statuses | System_ID, Project_ID, Stage, Stage_Status, Start_Date, End_Date, Remark |
| Engineering | System_ID, Project_ID, Product_ID, Product_Revision_ID, Drawing_No, Drawing_Revision, BOM_Revision_ID, Status_ID, Process, Machine, Tooling, Packaging, Engineering_Notes, Submitted_By, Approved_By, Create_Date, Created_By |
| BOM_Header | System_ID, BOM_ID, Product_Revision_ID, Revision_No, Rev_Date, Change_Reason, Changed_By, Approved_By, Status_ID, Is_Active, Create_Date, Created_By |
| BOM_Detail | System_ID, BOM_Revision_ID, Parent_BOM_Line_ID, Level_No, Sequence_No, Material_ID, Qty_Per, UoM, Scrap_%, Net_Qty(calc), Gross_Qty(calc), Make_Buy, Approved_Supplier, Create_Date, Created_By |
| Supplier_Quotes | System_ID, Project_ID, Material_ID, Qty_Required, Quoted_Date, Valid_Until, Status_ID, Selected_Line_ID, Create_Date, Created_By |
| Supplier_Quote_Lines | System_ID, Supplier_Quote_ID, Supplier_ID, Unit_Price, Currency, MOQ, Lead_Time_Days, Payment_Terms, Incoterm, Freight, Valid_To, Price_Source, Price_Confirm_Status, Quality_Score, OnTime_Score, Capacity_Score, Risk_Score, Longevity_Score, Score_Total(calc), Selected, Landed_Estimate(calc), Create_Date, Created_By |
| Procurement | System_ID, Project_ID, Material_ID, Supplier_ID, Buyer, Currency, Fx_Rate, Buy_Qty, Unit_Purchase_Price, MOQ_Applied, Lead_Time_Days, Payment_Terms, Incoterm, Freight_Per_Unit, Insurance_Per_Unit, Customs_Per_Unit, Handling_Per_Unit, Other_Per_Unit, Landed_Cost_Unit(calc), Price_Date, Price_Source, Price_Valid_To, Price_Expired_Flag(calc), Status_ID, Create_Date, Created_By |
| Purchase_Prices | System_ID, Material_ID, Supplier_ID, Revision, Price_Date, Price_Source, Currency, Fx_Rate, Unit_Price, Freight_Per_Unit, Insurance_Per_Unit, Customs_Per_Unit, Handling_Per_Unit, Other_Per_Unit, Landed_Cost_Unit(calc), Valid_From, Valid_To, Expired_Flag(calc), Status_ID, Approved_By, Create_Date, Created_By |
| Planning | System_ID, Project_ID, BOM_Revision_ID, BOM_Line_ID, Product_ID, Material_ID, Make_Buy, Order_Qty(calc زنده از Order_Lines), Bom_Qty_Per, Scrap_%, Required_Qty(calc), Stock_On_Hand(calc), Reserved(calc), Deficit(calc), Buy_Qty(calc), Lead_Time_Days, Supply_Date(calc), Status_ID, Create_Date, Created_By |
| Schedule | System_ID, Project_ID, Activity, Start_Date, End_Date, Duration(calc), Responsible, Dependency, Status_ID, Delay_Days(calc), Delay_Reason, Create_Date, Created_By |
| Cost_Lines | System_ID, Costing_ID, Project_ID, Category, Material_ID, Supplier_ID, Qty, UoM, Unit_Cost, Amount(calc), Source_Ref, Memo, Create_Date, Created_By |
| Costing | System_ID, Project_ID, Revision, Status_ID, Cost_Basis, Qty_Basis, Direct_Material(calc), Purchased_Parts(calc), Direct_Labor(calc), Mfg_Overhead(calc), Scrap(calc), Packaging(calc), Logistics(calc), Tooling(calc), SG&A(calc), Other_Allocated, Total_Cost(calc), Cost_Per_Unit(calc), Approved_By, Create_Date, Created_By |
| Budget | System_ID, Project_ID, Revision, Period, Budget_Material, Budget_Labor, Budget_Overhead, Budget_Procurement, Budget_Logistics, Budget_Tooling, Budget_Other, Total_Budget(calc), Actual_Material(calc), Actual_Labor(calc), Actual_Overhead(calc), Actual_Logistics(calc), Actual_Other(calc), Actual_Total(calc), Variance_Total(calc), Budget_vs_Actual(calc), Create_Date, Created_By |
| Payment_Terms | System_ID, Quotation_ID, Advance_Payment_%, Payment_On_Delivery_%, Credit_Days, Installment_Terms, Bank_Guarantee, Currency, Payment_Risk_Level, Fin_Cost_Rate_Annual, Fin_Cost_Amount(calc), Is_Active_FLAG, Create_Date, Created_By |
| Sales_Quotation | System_ID, Project_ID, Quotation_Code, Scenario, Costing_ID, Total_Cost(calc), Desired_Margin_%, Margin_Amount(calc), Commercial_Adjustment, Risk_Adjustment, Financial_Adjustment(calc), Proposed_Selling_Price(calc), Currency, Basis_Memo, Status_ID, Create_Date, Created_By |
| Quotation_Versions | System_ID, Quotation_ID, Version_No, Costing_Revision_ID, Total_Cost(calc), Margin_Amount(calc), Proposed_Selling_Price(calc), Change_Summary, Approved_Step, Status_ID, Create_Date, Created_By |
| Management_Approval | System_ID, Quotation_Version_ID, Project_ID, Decision, Decision_Maker, Decision_Date, Comment, Version_Ref, Status_ID, Create_Date, Created_By |

## جداول Audit/کنترل
| جدول | ستون‌ها |
|---|---|
| Workflow_History | Workflow_History_ID, Project_ID, Entity_Type, Entity_ID, From_Status, To_Status, Action_Date, User_ID, Comment |
| Change_Log | Change_Log_ID, Project_ID, Entity_Type, Entity_ID, Field_Changed, Old_Value, New_Value, Change_Reason, Changed_By, Change_Date, Revision_After, Create_Date, Created_By |
| Issues | Issue_ID, Project_ID, Error_Code, Error_Description, Owner, Required_Action, Status, Severity, Raised_Date, Create_Date, Created_By |
| Error_Checks | Rule_Key, Error_Code, Check_Description, Severity, Owner, Formula_Result(calc), Message, Required_Action, Status, Create_Date, Created_By |
| Test_Results | Test_ID, Test_Scenario, Input, Expected_Result, Actual_Result, Pass_Fail, Issue, Executed_Date, By |

## شیت‌های نمایش
Dashboard (KPI + Funnel + Timeline)، Workflow (گردش کار)، _Trace (Drill-down)، Report_Quotation (گزارش اجرایی)، _NamedRanges (مستند Named Ranges).
