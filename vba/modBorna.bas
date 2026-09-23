Attribute VB_Name = "modBorna"
' ============================================================================
'  سامانه مدیریت پروژه — برنا کامپوزیت ایرانیان
'  ماژول کنترل دسترسی نقش‌محور (Login / Role-based visibility)
'  نسخه 1.4 — مطابق ماتریس دسترسی شیت «تنظیمات»
'
'  نصب:  ۱) فایل xlsx را باز کنید  ۲) Alt+F11  ۳) File > Import File را زده
'        و modBorna.bas را ایمپورت کنید  ۴) کد این تابع را به ThisWorkbook اضافه کنید:
'            Private Sub Workbook_Open(): modBorna.Borna_Init: End Sub
'        ۵) Save As xlsm (Macro-Enabled) — راهنمای کامل: docs/INSTALL.md
'
'  ماکروها:  Borna_Login (ورود، Alt+F8)  |  Borna_Logout (خروج)  |  Borna_Admin (ورود اضطراری IT)
'
'  توجه: رمزها نمونه‌اند؛ پیش از استقرار در شبکه تغییرشان دهید (اینجا + شیت‌ها + ساختار).
'  توجه امنیتی: این لایه «بازدارنده» است. حفاظت فیزیکی داده با NTFS/سایت‌کنترل — اسناد IT.
' ============================================================================
Option Explicit

' ---- پیکربندی (هماهنگ با build_workbook.py) -----------------------------
Private Const SH_LOGIN As String = "ورود"
Private Const SH_GUIDE As String = "راهنما"
Private Const SH_DASH As String = "داشبورد"
Private Const SH_SET As String = "تنظیمات"
Private Const PWD_SHEETS As String = "Borna@1405"          ' رمز قفل شیت‌ها
Private Const PWD_STRUCTURE As String = "Borna-Admin-1405" ' رمز قفل ساختار فایل

Private Const USER_FIRST As Long = 14, USER_LAST As Long = 23   ' جدول کاربران تنظیمات
Private Const ACC_HEADER As Long = 26                            ' سرستون ماتریس
Private Const ACC_FIRST As Long = 27, ACC_LAST As Long = 45     ' ردیف‌های ماتریس

Private gRole As String          ' کد نقش فعال (خالی = مهمان)
Private gRoleName As String

' ---- ورود -----------------------------------------------------------------
Public Sub Borna_Login()
    Dim wsS As Worksheet, u As String, p As String, r As Long
    Dim found As Boolean
    Set wsS = ThisWorkbook.Worksheets(SH_SET)
    u = Trim$(CStr(ThisWorkbook.Worksheets(SH_LOGIN).Range("B5").Value))
    p = Trim$(CStr(ThisWorkbook.Worksheets(SH_LOGIN).Range("B6").Value))
    If u = "" Then
        MsgBox "نام کاربری را وارد کنید.", vbExclamation, "ورود"
        Exit Sub
    End If
    For r = USER_FIRST To USER_LAST
        If StrComp(Trim$(CStr(wsS.Cells(r, 1).Value)), u, vbTextCompare) = 0 Then
            If CStr(wsS.Cells(r, 2).Value) = p Then
                gRole = Trim$(CStr(wsS.Cells(r, 4).Value))
                gRoleName = Trim$(CStr(wsS.Cells(r, 5).Value))
                found = True
            Else
                MsgBox "رمز عبور نادرست است.", vbCritical, "ورود"
                Exit Sub
            End If
            Exit If
        End If
    Next r
    If Not found Then
        MsgBox "این نام کاربری در جدول کاربران (تنظیمات) ثبت نشده است.", vbCritical, "ورود"
        Exit Sub
    End If
    ApplyRole
    ThisWorkbook.Worksheets(SH_LOGIN).Range("B6").ClearContents
    ThisWorkbook.Worksheets(SH_LOGIN).Range("C6").Value = _
        gRoleName & "  —  ورود: " & Format(Now, "yyyy/mm/dd hh:nn")
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_DASH).Activate
    If Err.Number <> 0 Then ThisWorkbook.Worksheets(SH_GUIDE).Activate
    On Error GoTo 0
End Sub

' ---- اعمال ماتریس دسترسی ---------------------------------------------------
Private Sub ApplyRole()
    Dim wsS As Worksheet, ws As Worksheet, col As Long, r As Long
    Dim perm As String, name_ As String
    Set wsS = ThisWorkbook.Worksheets(SH_SET)
    col = Application.Match(gRole, wsS.Range(wsS.Cells(ACC_HEADER, 2), _
          wsS.Cells(ACC_HEADER, 11)), 0)
    If IsError(col) Then
        MsgBox "نقش «" & gRole & "» در سرستون ماتریس دسترسی یافت نشد.", vbCritical
        Exit Sub
    End If
    With ThisWorkbook
        .Unprotect Password:=PWD_STRUCTURE
        ' 1) ابتدا همه قفل/مخفی
        For Each ws In .Worksheets
            If ws.Name <> SH_LOGIN And ws.Name <> SH_GUIDE Then
                ws.Protect Password:=PWD_SHEETS, UserInterfaceOnly:=False, _
                          AllowFormattingCells:=True, AllowInsertingRows:=True, _
                          AllowDeletingRows:=True, AllowSorting:=True, _
                          AllowFiltering:=True
            End If
        Next ws
        ' 2) سپس طبق ماتریس باز/نمایش
        For r = ACC_FIRST To ACC_LAST
            name_ = Trim$(CStr(wsS.Cells(r, 1).Value))
            If name_ <> "" Then
                On Error Resume Next
                Set ws = .Worksheets(name_)
                On Error GoTo 0
                If Not ws Is Nothing Then
                    perm = UCase$(Trim$(CStr(wsS.Cells(r, col).Value)))
                    Select Case perm
                        Case "W"
                            ws.Visible = xlSheetVisible
                            ws.Unprotect Password:=PWD_SHEETS
                            If gRole <> "ADMIN" Then ws.EnableSelection = xlUnlockedCells
                            ws.Protect Password:=PWD_SHEETS, UserInterfaceOnly:=True, _
                                      AllowFormattingCells:=True, AllowInsertingRows:=True, _
                                      AllowDeletingRows:=True, AllowSorting:=True, _
                                      AllowFiltering:=True
                        Case "R"
                            ws.Visible = xlSheetVisible
                            ' قفل با رمز: بدون تغییر محتوا (فرمول‌ها و متن‌ها محافظت می‌شوند)
                        Case Else
                            If ws.Name <> SH_LOGIN And ws.Name <> SH_GUIDE Then _
                                ws.Visible = xlSheetVeryHidden
                    End Select
                End If
            End If
            Set ws = Nothing
        Next r
        If gRole <> "ADMIN" Then .Protect Password:=PWD_STRUCTURE, Structure:=True
    End With
    MsgBox "ورود موفق — نقش: " & gRoleName, vbInformation, "سامانه پروژه برنا"
End Sub

' ---- خروج -------------------------------------------------------------------
Public Sub Borna_Logout()
    Dim ws As Worksheet
    With ThisWorkbook
        .Unprotect Password:=PWD_STRUCTURE
        For Each ws In .Worksheets
            If ws.Name <> SH_LOGIN And ws.Name <> SH_GUIDE Then
                ws.Protect Password:=PWD_SHEETS
                ws.Visible = xlSheetVeryHidden
            End If
        Next ws
        .Protect Password:=PWD_STRUCTURE, Structure:=True
        .Worksheets(SH_LOGIN).Range("C6").Value = "مهمان — بدون نشست فعال"
        .Worksheets(SH_LOGIN).Range("B6").ClearContents
        .Worksheets(SH_LOGIN).Activate
    End With
    gRole = "": gRoleName = ""
End Sub

' ---- وضعیت شروع فایل (از Workbook_Open صدا زده شود) --------------------------
Public Sub Borna_Init()
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_LOGIN).Range("C6").Value = "مهمان — بدون ورود"
    If gRole = "" Then Borna_Logout
End Sub

' ---- ورود اضطراری مدیر سیستم --------------------------------------------------
Public Sub Borna_Admin()
    Dim pw As String, ws As Worksheet
    pw = InputBox("رمز مدیر سیستم:", "Borna_Admin")
    If pw <> PWD_STRUCTURE Then
        MsgBox "رمز اشتباه است.", vbCritical: Exit Sub
    End If
    With ThisWorkbook
        .Unprotect Password:=PWD_STRUCTURE
        For Each ws In .Worksheets
            ws.Visible = xlSheetVisible
            ws.Unprotect Password:=PWD_SHEETS
        Next ws
    End With
    MsgBox "حالت مدیر سیستم فعال شد — پس از اتمام کار، فایل را ببندید و دوباره باز کنید.", vbInformation
End Sub
