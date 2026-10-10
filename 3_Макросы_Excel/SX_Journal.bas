Attribute VB_Name = "SX_Journal"
Option Explicit
'==========================================================================
' SX_Journal - export of the cable journal of an Odnolineyka v4 workbook
'              into a light values-only file for AutoCAD Data Link.
'              Lives in PERSONAL.XLSB, works on ActiveWorkbook.
'
' Macro:
'   SX_ExportJournal - sheet "Kabelnyi_zhurnal" A2:M... (+ summary O4:Q...)
'                      -> <workbook name>_<journal>.xlsx next to the workbook
'                      (same file every time: the Data Link in AutoCAD stays valid).
' The workbook itself is not changed. Russian text is encoded (^XXXX), see U().
'==========================================================================

Private Function U(ByVal s As String) As String
    Dim i As Long, n As Long, ch As String, r As String
    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "^" Then
            r = r & ChrW(CLng("&H" & Mid$(s, i + 1, 4)))
            i = i + 5
        Else
            r = r & ch
            i = i + 1
        End If
    Loop
    U = r
End Function

Public Sub SX_ExportJournal()
    Dim wb As Workbook, ws As Worksheet, nb As Workbook, t As Worksheet
    Dim src As Range, sm As Range, path As String, base As String, nrow As Long
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets(U("^041A^0430^0431^0435^043B^044C^043D^044B^0439_^0436^0443^0440^043D^0430^043B"))
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox U("^0412 ^0430^043A^0442^0438^0432^043D^043E^0439 ^043A^043D^0438^0433^0435 ^043D^0435^0442 ^043B^0438^0441^0442^0430 ^00AB^041A^0430^0431^0435^043B^044C^043D^044B^0439_^0436^0443^0440^043D^0430^043B^00BB. ^041E^0442^043A^0440^043E^0439^0442^0435 ^043A^043D^0438^0433^0443 ^043E^0434^043D^043E^043B^0438^043D^0435^0439^043A^0438 v4."), vbExclamation
        Exit Sub
    End If
    If wb.Path = "" Then
        MsgBox U("^0421^043D^0430^0447^0430^043B^0430 ^0441^043E^0445^0440^0430^043D^0438^0442^0435 ^043A^043D^0438^0433^0443: ^0444^0430^0439^043B ^0436^0443^0440^043D^0430^043B^0430 ^043A^043B^0430^0434^0451^0442^0441^044F ^0440^044F^0434^043E^043C ^0441 ^043D^0435^0439."), vbExclamation
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo fail
    wb.Application.Calculate
    Set src = ws.Range("A5").CurrentRegion
    Set sm = ws.Range("O4").CurrentRegion
    nrow = src.Rows.Count - 1
    Set nb = Workbooks.Add(xlWBATWorksheet)
    Set t = nb.Worksheets(1)
    t.Name = U("^0416^0443^0440^043D^0430^043B")
    t.Range("A1").Resize(src.Rows.Count, src.Columns.Count).Value = src.Value
    t.Rows(1).Font.Bold = True
    t.Columns.AutoFit
    Set t = nb.Worksheets.Add(After:=nb.Worksheets(1))
    t.Name = U("^0421^0432^043E^0434^043A^0430")
    t.Range("A1").Resize(sm.Rows.Count, sm.Columns.Count).Value = sm.Value
    t.Rows(1).Font.Bold = True
    t.Columns.AutoFit
    nb.Worksheets(1).Activate
    base = wb.Name
    If InStrRev(base, ".") > 0 Then base = Left$(base, InStrRev(base, ".") - 1)
    path = wb.Path & Application.PathSeparator & base & U("_^0436^0443^0440^043D^0430^043B.xlsx")
    nb.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
    nb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox U("^0416^0443^0440^043D^0430^043B ^0432^044B^0433^0440^0443^0436^0435^043D (^0442^043E^043B^044C^043A^043E ^0437^043D^0430^0447^0435^043D^0438^044F). ^0412 AutoCAD ^0441^0432^044F^0436^0438^0442^0435 ^0442^0430^0431^043B^0438^0446^0443 (Data Link) ^0441 ^044D^0442^0438^043C ^0444^0430^0439^043B^043E^043C:") & vbCrLf & path & vbCrLf & U("^0421^0442^0440^043E^043A: ") & nrow, vbInformation
    Exit Sub
fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not nb Is Nothing Then nb.Close SaveChanges:=False
    MsgBox U("^0412^044B^0433^0440^0443^0437^043A^0430 ^043D^0435 ^0443^0434^0430^043B^0430^0441^044C: ") & Err.Description, vbCritical
End Sub
