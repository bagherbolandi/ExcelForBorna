Attribute VB_Name = "modBornaSplit"
' ============================================================================
'  Borna PM — حالت شبکه‌ای چندفایلی (دسترسی واقعی با NTFS + Publish/Collect)
'  این ماژول فقط روی فایل مادر (MASTER) نصب می‌شود؛ فایل واحدها ماکرو ندارد.
'
'  نصب در master:  File > Import File → modBorna.bas و modBornaSplit.bas
'
'  ماکروها:
'    Borna_SplitSetup    یک‌بار پس از کپی ساختار out/split کنار فایل مادر
'    Borna_Collect       دریافت داده‌های واحدها → فایل مادر (فقط سلول‌های unlocked هر برگه)
'    Borna_Publish       انتشار اسنپ‌شات تازه (مقادیر فایل مادر) در فایل واحدها
'    Borna_Sync          Collect + Publish (چرخه روزانهٔ مدیر پروژه)
'
'  ثاب‌ها باید با تنظیماتِ build_workbook.py هماهنگ بمانند:
'    جدول ماتریس دسترسی: سرستون ۲۶، ردیف‌های ۲۷..۴۵ (ستون A=شیت، B..K=نقش‌ها)
'    جدول اتصال واحدها: سرستون ۱۰۴، ردیف‌های ۱۰۵..۱۱۱
'    جدول کاربران (باید در فایل واحدها دیده نشود): ردیف‌های ۱۴..۲۳
'    مسیر فایل واحدها = ThisWorkbook.Path & "\" & <پوشه> & "\" & <فایل>
' ============================================================================
Option Explicit

Private Const SH_SET As String = "تنظیمات"
Private Const ACC_HDR As Long = 26
Private Const ACC_FIRST As Long = 27, ACC_LAST As Long = 45
Private Const CONN_HDR As Long = 104
Private Const CONN_FIRST As Long = 105, CONN_LAST As Long = 111
Private Const USR_FIRST As Long = 14, USR_LAST As Long = 23
Private Const PWD_M As String = "Borna@1405"   ' = PWD_SHEETS در modBorna و build_workbook.py

' ---------------- ابزار داخلی ----------------------------------------------
Private Function DeptInputSheets(wsS As Worksheet, colsList As String) As Collection
    ' شیت‌هایی که دست‌کم یکی از نقش‌های واحد در ماتریس «W» دارد
    Dim res As New Collection, cl As Variant, r As Long, mcol As Variant
    For Each cl In Split(colsList, ",")
        mcol = Application.Match(Trim$(CStr(cl)), _
               wsS.Range(wsS.Cells(ACC_HDR, 2), wsS.Cells(ACC_HDR, 11)), 0)
        If Not IsError(mcol) Then
            For r = ACC_FIRST To ACC_LAST
                If UCase$(Trim$(CStr(wsS.Cells(r, 1 + mcol).Value))) = "W" Then
                    On Error Resume Next
                    res.Add CStr(wsS.Cells(r, 1).Value), CStr(wsS.Cells(r, 1).Value)
                    On Error GoTo 0
                End If
            Next r
        End If
    Next cl
    Set DeptInputSheets = res
End Function

Private Sub CopyUnlocked(wsSrc As Worksheet, wsDst As Worksheet, ByRef n As Long)
    ' کپی مقدارِ سلول‌های unlocked (ورودی‌های واحد) به فایل مادر
    Dim u As Range, c As Range
    On Error Resume Next
    Set u = wsSrc.UsedRange
    On Error GoTo 0
    If u Is Nothing Then Exit Sub
    For Each c In u
        If Not c.Protection.Locked Then
            If Not c.Protection.Locked Then
                If Len(CStr(c.Value)) > 0 Or Not IsEmpty(c.Value) Then
                    wsDst.Cells(c.Row, c.Column).Value = c.Value
                    n = n + 1
                End If
            End If
        End If
    Next c
End Sub

' ---------------- دریافت (Collect) ------------------------------------------
Public Sub Borna_Collect()
    Dim wsS As Worksheet, row As Long, n As Long, miss As String, lockedList As String
    Set wsS = ThisWorkbook.Worksheets(SH_SET)
    Dim root As String: root = ThisWorkbook.Path & Application.PathSeparator
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    For row = CONN_FIRST To CONN_LAST
        Dim folder As String, fname As String, cols As String, full As String
        folder = Trim$(CStr(wsS.Cells(row, 1).Value))
        fname = Trim$(CStr(wsS.Cells(row, 2).Value))
        cols = Trim$(CStr(wsS.Cells(row, 6).Value))
        If Len(folder) > 0 And Len(fname) > 0 Then
            full = root & folder & Application.PathSeparator & fname
            If Dir(full) = "" Then
                miss = miss & folder & "؛ "
            Else
                Dim wbd As Workbook
                On Error Resume Next
                Set wbd = Workbooks.Open(full, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True)
                If Not wbd Is Nothing Then
                    Dim inputs As Collection, sh As Variant
                    Set inputs = DeptInputSheets(wsS, cols)
                    For Each sh In inputs
                        Dim wsMx As Worksheet, wsDx As Worksheet
                        On Error Resume Next
                        Set wsMx = ThisWorkbook.Worksheets(CStr(sh))
                        Set wsDx = wbd.Worksheets(CStr(sh))
                        On Error GoTo 0
                        If Not (wsMx Is Nothing) And Not (wsDx Is Nothing) Then
                            wsMx.Unprotect Password:=PWD_M      ' نوشتن برنامه‌ای روی شیت قفل
                            CopyUnlocked wsDx, wsMx, n
                            wsMx.Protect Password:=PWD_M, AllowFormattingCells:=True, _
                                          AllowInsertingRows:=True, AllowDeletingRows:=True, _
                                          AllowSorting:=True, AllowFiltering:=True
                        Else
                            lockedList = lockedList & folder & "/" & CStr(sh) & "؛ "
                        End If
                        Set wsMx = Nothing: Set wsDx = Nothing
                    Next sh
                    wbd.Close savechanges:=False
                    Set wbd = Nothing
                    wsS.Cells(row, 5).Value = Now
                Else
                    lockedList = lockedList & folder & " (باز است)؛ "
                    On Error GoTo 0
                End If
            End If
        End If
    Next row
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(SH_SET).Parent.Save
    MsgBox "دریافت شد — " & n & " ورودی تازه." & vbCrLf & _
           IIf(miss = "", "", "فایل‌های نبود: " & miss & vbCrLf) & _
           IIf(lockedList = "", "", "جاافتاده: " & lockedList), _
           IIf(miss = "" And lockedList = "", vbInformation, vbExclamation), "Borna_Collect"
End Sub

' ---------------- انتشار (Publish) --------------------------------------------
Public Sub Borna_Publish()
    Dim wsS As Worksheet, row As Long, updated As Long, miss As String
    Set wsS = ThisWorkbook.Worksheets(SH_SET)
    Dim root As String: root = ThisWorkbook.Path & Application.PathSeparator
    Application.ScreenUpdating = False
    For row = CONN_FIRST To CONN_LAST
        Dim folder As String, fname As String, tok As String, cols As String, full As String
        folder = Trim$(CStr(wsS.Cells(row, 1).Value))
        fname = Trim$(CStr(wsS.Cells(row, 2).Value))
        tok = Trim$(CStr(wsS.Cells(row, 3).Value))
        cols = Trim$(CStr(wsS.Cells(row, 6).Value))
        If Len(folder) > 0 And Len(fname) > 0 Then
            full = root & folder & Application.PathSeparator & fname
            If Dir(full) = "" Then
                miss = miss & folder & "؛ "
            Else
                Dim wbd As Workbook
                On Error Resume Next
                Set wbd = Workbooks.Open(full, UpdateLinks:=0)
                On Error GoTo 0
                If wbd Is Nothing Then
                    miss = miss & folder & " (باز/قفل)؛ "
                Else
                    Dim inputs As Collection: Set inputs = DeptInputSheets(wsS, cols)
                    Dim wsD As Worksheet, wsM As Worksheet, skipIt As Boolean, i As Long
                    For Each wsD In wbd.Worksheets
                        skipIt = False
                        For i = 1 To inputs.Count
                            If inputs(i) = wsD.Name Then skipIt = True: Exit For
                        Next i
                        If Not skipIt Then
                            On Error Resume Next
                            Set wsM = ThisWorkbook.Worksheets(wsD.Name)
                            On Error GoTo 0
                            If Not wsM Is Nothing Then
                                On Error Resume Next
                                wsD.Unprotect Password:=tok
                                ' کپی مقدار (نه فرمول) کل محدودهٔ اشغال‌شده
                                wsD.Range(wsM.UsedRange.Address).Value = wsM.UsedRange.Value
                                On Error GoTo 0
                                If wsD.Name = SH_SET Then
                                    ' هرگز رمز کاربران/توکن‌ها را به فایل واحد نده:
                                    wsD.Range("A" & USR_FIRST & ":F" & USR_LAST).ClearContents
                                    wsD.Range("A" & CONN_HDR - 1 & ":F" & CONN_LAST).ClearContents
                                End If
                                wsD.Cells.Locked = True
                                wsD.Protect Password:=tok
                            End If
                        End If
                    Next wsD
                    wbd.Save
                    wbd.Close savechanges:=False
                    Set wbd = Nothing
                    wsS.Cells(row, 4).Value = Now
                    updated = updated + 1
                End If
            End If
        End If
    Next row
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    MsgBox "انتشار انجام شد — " & updated & " فایل واحد تازه شد." & _
           IIf(miss = "", "", vbCrLf & "نبود/قفل: " & miss), vbInformation, "Borna_Publish"
End Sub

' ---------------- چرخه کامل ----------------------------------------------------
Public Sub Borna_Sync()
    Borna_Collect
    Borna_Publish
End Sub

' ---------------- آماده‌سازی اولیهٔ پوشه‌ها (یک‌بار) ------------------------------
Public Sub Borna_SplitSetup()
    Dim wsS As Worksheet, row As Long, made As Long
    Set wsS = ThisWorkbook.Worksheets(SH_SET)
    Dim root As String, tpl As String, dst As String
    root = ThisWorkbook.Path & Application.PathSeparator
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    For row = CONN_FIRST To CONN_LAST
        Dim folder As String, fname As String
        folder = Trim$(CStr(wsS.Cells(row, 1).Value))
        fname = Trim$(CStr(wsS.Cells(row, 2).Value))
        If Len(folder) > 0 Then
            If Not fso.FolderExists(root & folder) Then fso.CreateFolder (root & folder)
            dst = root & folder & Application.PathSeparator & fname
            tpl = root & "split" & Application.PathSeparator & folder & Application.PathSeparator & fname
            If Not fso.FileExists(dst) Then
                If fso.FileExists(tpl) Then
                    fso.CopyFile tpl, dst
                    made = made + 1
                End If
            End If
        End If
    Next row
    MsgBox "پوشه‌ها آماده شد؛ " & made & " فایل واحد از قالب split کپی شد (فایل‌های موجود دست نخورد)." & _
           vbCrLf & "اکنون icacls/tools/deploy_acl.ps1 را اجرا کنید تا هر واحد فقط پوشهٔ خودش را ببیند.", vbInformation, "Borna_SplitSetup"
End Sub
