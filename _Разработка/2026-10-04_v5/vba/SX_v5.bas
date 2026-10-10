Attribute VB_Name = "SX_v5"
Option Explicit
'==========================================================================
' SX v5 - service macros of the Odnolineyka v5 workbook
'   SX_RepairFormulas - re-apply formulas of the service and result columns
'                       of table "Lines" (after rows are inserted / deleted / sorted)
'   SX_FillAuto       - emptied "auto" (gray) cells get their formula back
'   SX_AutoFillOff / SX_AutoFillRestore - while this workbook is active,
'       Excel option "fill formulas in tables" is off: a formula typed into
'       one cell does not spread to the whole column.
' Russian text is encoded (^XXXX) and decoded by U().
'==========================================================================
Private PrevAutoFill As Variant

Public Function U(ByVal s As String) As String
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

Public Function SxTable() As ListObject
    On Error Resume Next
    Set SxTable = ThisWorkbook.Worksheets(U("^041B^0438^043D^0438^0438")).ListObjects(U("^041B^0438^043D^0438^0438"))
End Function

Public Function SxTpl() As Worksheet
    On Error Resume Next
    Set SxTpl = ThisWorkbook.Worksheets(U("_^0444^043E^0440^043C^0443^043B^044B"))
End Function

Public Sub SX_AutoFillOff()
    On Error Resume Next
    If IsEmpty(PrevAutoFill) Then PrevAutoFill = Application.AutoCorrect.AutoFillFormulasInLists
    Application.AutoCorrect.AutoFillFormulasInLists = False
End Sub

Public Sub SX_AutoFillRestore()
    On Error Resume Next
    If Not IsEmpty(PrevAutoFill) Then
        Application.AutoCorrect.AutoFillFormulasInLists = PrevAutoFill
        PrevAutoFill = Empty
    End If
End Sub

Public Sub SX_FillAuto(ByVal rng As Range)
    Dim lo As ListObject, ws As Worksheet, c As Range, r As Variant, h As String, k As String
    Set lo = SxTable()
    Set ws = SxTpl()
    If lo Is Nothing Or ws Is Nothing Then Exit Sub
    For Each c In rng.Cells
        If Not c.HasFormula Then
            If Len(CStr(c.Formula)) = 0 Then
                h = CStr(lo.HeaderRowRange.Cells(1, c.Column - lo.Range.Column + 1).Value)
                r = Application.Match(h, ws.Columns(1), 0)
                If Not IsError(r) Then
                    k = CStr(ws.Cells(r, 3).Value)
                    If k = "a" Or k = "xa" Then c.Formula2R1C1 = CStr(ws.Cells(r, 2).Value)
                End If
            End If
        End If
    Next
End Sub

Public Sub SX_RepairFormulas()
    Dim lo As ListObject, ws As Worksheet, r As Long, k As String, f As String, calc As Long, bad As Long
    Set lo = SxTable()
    Set ws = SxTpl()
    If lo Is Nothing Or ws Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = -4135
    On Error Resume Next
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        k = CStr(ws.Cells(r, 3).Value)
        f = CStr(ws.Cells(r, 2).Value)
        If f <> "" Then
            If k = "c" Or k = "h" Then
                lo.ListColumns(CStr(ws.Cells(r, 1).Value)).DataBodyRange.Formula2R1C1 = f
                If Err.Number <> 0 Then
                    bad = bad + 1
                    Err.Clear
                End If
            ElseIf k = "a" Or k = "xa" Then
                SX_FillAuto lo.ListColumns(CStr(ws.Cells(r, 1).Value)).DataBodyRange
                Err.Clear
            End If
        End If
    Next
    On Error GoTo 0
    Application.Calculation = calc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If bad > 0 Then
        MsgBox U("^041D^0435 ^0443^0434^0430^043B^043E^0441^044C ^0437^0430^043F^0438^0441^0430^0442^044C ^0444^043E^0440^043C^0443^043B^044B ^0441^0442^043E^043B^0431^0446^043E^0432: ") & bad, vbExclamation
    Else
        Application.StatusBar = U("^0424^043E^0440^043C^0443^043B^044B ^0442^0430^0431^043B^0438^0446^044B ^00AB^041B^0438^043D^0438^0438^00BB ^0432^043E^0441^0441^0442^0430^043D^043E^0432^043B^0435^043D^044B")
    End If
End Sub

Private Function SxSheetName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", "?", "*", "[", "]", ":")
        s = Replace(s, ch, "_")
    Next
    If Len(s) > 31 Then s = Left$(s, 31)
    SxSheetName = s
End Function

Private Function SxSpill(ByVal c As Range) As Range
    If c.HasSpill Then
        Set SxSpill = c.SpillingToRange
    Else
        Set SxSpill = c
    End If
End Function

Private Sub SxCopyBlock(ByVal src As Range, ByVal ex As Workbook, ByVal nm As String)
    Dim ws As Worksheet
    Set ws = ex.Worksheets.Add(After:=ex.Worksheets(ex.Worksheets.Count))
    ws.Name = SxSheetName(nm)
    src.Copy
    ws.Range("A1").PasteSpecial -4163
    ws.Range("A1").PasteSpecial -4122
    ws.Range("A1").PasteSpecial 8
    Application.CutCopyMode = False
End Sub

' light file for AutoCAD Data Link: per board - table under the scheme, totals per input, cable journal, specification
Public Sub SX_ExportAutoCAD()
    Dim wb As Workbook, ex As Workbook, lo As ListObject, c As Range, sp As Range, b As String, p As String, msg As String
    Dim wsS As Worksheet, wsJ As Worksheet, wsP As Worksheet, n As Long
    Dim oS As Variant, oJ As Variant, oP As Variant
    Set wb = ThisWorkbook
    On Error GoTo fail
    Set wsS = wb.Worksheets(U("^0429^0438^0442"))
    Set wsJ = wb.Worksheets(U("^041A^0430^0431^0435^043B^044C^043D^044B^0439_^0436^0443^0440^043D^0430^043B"))
    Set wsP = wb.Worksheets(U("^0421^043F^0435^0446^0438^0444^0438^043A^0430^0446^0438^044F"))
    Set lo = wb.Worksheets(U("^0429^0438^0442^044B")).ListObjects(U("^0429^0438^0442^044B"))
    p = wb.Path & Application.PathSeparator & Left$(wb.Name, InStrRev(wb.Name, ".") - 1) & U("_AutoCAD.xlsx")
    oS = wsS.Range("C3").Value
    oJ = wsJ.Range("E1").Value
    oP = wsP.Range("D1").Value
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set ex = Workbooks.Add(-4167)
    For Each c In lo.ListColumns(U("^0429^0438^0442")).DataBodyRange.Cells
        b = Trim$(CStr(c.Value))
        If Len(b) > 0 Then
            wsS.Range("C3").Value = b
            wsJ.Range("E1").Value = b
            wsP.Range("D1").Value = b
            Application.Calculate
            Set sp = SxSpill(wsS.Range("B7"))
            SxCopyBlock wsS.Range(wsS.Range("B6"), wsS.Cells(sp.Row + sp.Rows.Count - 1, 10)), ex, b & U(" ^0441^0445^0435^043C^0430")
            Set sp = SxSpill(wsS.Range("M6"))
            SxCopyBlock wsS.Range(wsS.Range("L6"), wsS.Cells(15, sp.Column + sp.Columns.Count - 1)), ex, b & U(" ^0438^0442^043E^0433^0438")
            Set sp = SxSpill(wsJ.Range("A5"))
            SxCopyBlock wsJ.Range(wsJ.Range("A2"), wsJ.Cells(sp.Row + sp.Rows.Count - 1, 13)), ex, b & U(" ^0436^0443^0440^043D^0430^043B")
            Set sp = SxSpill(wsP.Range("B3"))
            SxCopyBlock wsP.Range(wsP.Range("B2"), wsP.Cells(sp.Row + sp.Rows.Count - 1, 10)), ex, b & U(" ^0441^043F^0435^0446")
            n = n + 1
        End If
    Next
    If n > 0 Then ex.Worksheets(1).Delete
    wsS.Range("C3").Value = oS
    wsJ.Range("E1").Value = oJ
    wsP.Range("D1").Value = oP
    Application.Calculate
    ex.SaveAs p, 51
    ex.Close False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Application.Visible Then
        MsgBox U("^0412^044B^0433^0440^0443^0436^0435^043D^043E ^0449^0438^0442^043E^0432: ") & n & vbCrLf & p, vbInformation
    Else
        Application.StatusBar = U("^0412^044B^0433^0440^0443^0436^0435^043D^043E ^0449^0438^0442^043E^0432: ") & n & " | " & p
    End If
    Exit Sub
fail:
    msg = Err.Description
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error Resume Next
    wsS.Range("C3").Value = oS
    wsJ.Range("E1").Value = oJ
    wsP.Range("D1").Value = oP
    If Not ex Is Nothing Then ex.Close False
    If Application.Visible Then
        MsgBox U("^0412^044B^0433^0440^0443^0437^043A^0430 ^043D^0435 ^0443^0434^0430^043B^0430^0441^044C: ") & msg, vbExclamation
    Else
        Application.StatusBar = U("^0412^044B^0433^0440^0443^0437^043A^0430 ^043D^0435 ^0443^0434^0430^043B^0430^0441^044C: ") & msg
    End If
End Sub

' ===== point 4: sort board, freeze phases, snapshot of the issue, labels
Private Function SxCol(ByVal lo As ListObject, ByVal h As String) As Long
    On Error Resume Next
    SxCol = 0
    SxCol = lo.ListColumns(h).Index
    Err.Clear
End Function

Private Function SxBoardOf(ByVal lo As ListObject) As String
    Dim c As Range, r As Long, b As String
    On Error Resume Next
    Set c = ActiveCell
    If Not c Is Nothing Then
        If c.Worksheet.Name = lo.Parent.Name Then
            If Not Application.Intersect(c, lo.DataBodyRange) Is Nothing Then
                r = c.Row - lo.DataBodyRange.Row + 1
                b = CStr(lo.DataBodyRange.Cells(r, SxCol(lo, U("^0429^0438^0442"))).Value)
            End If
        End If
    End If
    On Error GoTo 0
    If b = "" And Application.Visible Then b = InputBox(U("^0429^0438^0442 (^0438^043C^044F, ^043A^0430^043A ^0432 ^00AB^041B^0438^043D^0438^044F^0445^00BB):"), U("^041E^0434^043D^043E^043B^0438^043D^0435^0439^043A^0430 v5"))
    SxBoardOf = Trim$(b)
End Function

Private Function SxCtlOrder(ByVal c As String) As String
    c = LCase$(Trim$(c))
    If c = "relay" Then
        SxCtlOrder = "1"
    ElseIf c = "dim" Then
        SxCtlOrder = "2"
    ElseIf c = "dali" Then
        SxCtlOrder = "3"
    ElseIf c = LCase$(U("^0448^0442^043E^0440^044B")) Then
        SxCtlOrder = "4"
    ElseIf c = LCase$(U("^043A^043B^0430^043F^0430^043D")) Then
        SxCtlOrder = "5"
    ElseIf c = "" Then
        SxCtlOrder = "9"
    Else
        SxCtlOrder = "6"
    End If
End Function

' rows of the board: input node keeps its order; then by control, code, room, line. Groups "gr", RCD blocks and manual modules stay together.
Public Sub SX_SortBoard(Optional ByVal board As String = "")
    Dim lo As ListObject, v As Variant, n As Long, i As Long, j As Long, k As Long, nu As Long, first As Long, t As String
    Dim cB As Long, cL As Long, cMode As Long, cSt As Long, cUz As Long, cMst As Long, cZone As Long, cCtl As Long, cCode As Long, cRoom As Long, cKey As Long
    Dim unitOf() As Long, uKey() As String, ord() As Long, rk() As Long, keys() As Variant, posIn() As Long
    Dim isNode As Boolean, isCont As Boolean, prevNode As Boolean, ln As String
    Set lo = SxTable()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If board = "" Then board = SxBoardOf(lo)
    If board = "" Then Exit Sub
    cB = SxCol(lo, U("^0429^0438^0442")): cL = SxCol(lo, U("^041D^043E^043C^0435^0440 ^043B^0438^043D^0438^0438")): cMode = SxCol(lo, U("^0420^0435^0436^0438^043C")): cSt = SxCol(lo, U("_^0441^0432^043E^0439 ^0430^0432^0442^043E^043C^0430^0442"))
    cUz = SxCol(lo, U("_^043D^0430 ^0448^0438^043D^0435 ^0423^0417^041E")): cMst = SxCol(lo, U("_^0431^043B^043E^043A ^043C^043E^0434^0443^043B^044F")): cZone = SxCol(lo, U("_^0437^043E^043D^0430")): cCtl = SxCol(lo, U("_^0443^043F^0440^0430^0432^043B^0435^043D^0438^0435 (^0438^0442^043E^0433)"))
    cCode = SxCol(lo, U("^041A^043E^0434")): cRoom = SxCol(lo, U("^041F^043E^043C^0435^0449^0435^043D^0438^0435")): cKey = SxCol(lo, U("_^043A^043B^044E^0447 ^0443^043F^043E^0440^044F^0434^043E^0447^0438^0432^0430^043D^0438^044F"))
    If cKey = 0 Or cB = 0 Then Exit Sub
    v = lo.DataBodyRange.Value
    n = UBound(v, 1)
    ReDim unitOf(1 To n): ReDim uKey(1 To n): ReDim posIn(1 To n)
    For i = 1 To n
        If CStr(v(i, cB)) = board Then
            If first = 0 Then first = i
            ln = Trim$(CStr(v(i, cL)))
            isNode = (CStr(v(i, cMode)) = U("^0432^0440^0443^0447^043D^0443^044E")) Or (CStr(v(i, cZone)) = U("^0432^0432^043E^0434^043D^043E^0439 ^0443^0437^0435^043B"))
            isCont = False
            If nu > 0 And i > first Then
                If ln = "" Then
                    isCont = True
                ElseIf Not isNode And Not prevNode Then
                    If Val(CStr(v(i, cSt))) = 0 Then isCont = True
                    If Val(CStr(v(i, cUz))) = 1 Then isCont = True
                    If CStr(v(i, cMst)) = "M" And CStr(v(i - 1, cMst)) = "M" Then isCont = True
                End If
            End If
            If isCont Then
                unitOf(i) = nu
            Else
                nu = nu + 1
                unitOf(i) = nu
                If isNode Then
                    uKey(nu) = "0" & Chr$(1) & Format$(i, "000000")
                Else
                    uKey(nu) = "1" & Chr$(1) & SxCtlOrder(CStr(v(i, cCtl))) & Chr$(1) & UCase$(CStr(v(i, cCode))) & Chr$(1) & UCase$(CStr(v(i, cRoom))) & Chr$(1) & UCase$(ln) & Chr$(1) & Format$(i, "000000")
                End If
            End If
            If ln <> "" Or isNode Then prevNode = isNode
        End If
    Next
    If nu < 2 Then Exit Sub
    ReDim ord(1 To nu)
    For i = 1 To nu: ord(i) = i: Next
    For i = 2 To nu
        k = ord(i): t = uKey(k): j = i - 1
        Do While j >= 1
            If uKey(ord(j)) <= t Then Exit Do
            ord(j + 1) = ord(j): j = j - 1
        Loop
        ord(j + 1) = k
    Next
    ReDim rk(1 To nu)
    For i = 1 To nu: rk(ord(i)) = i: Next
    ReDim keys(1 To n, 1 To 1)
    For i = 1 To n
        If CStr(v(i, cB)) = board Then
            posIn(unitOf(i)) = posIn(unitOf(i)) + 1
            keys(i, 1) = first + (rk(unitOf(i)) * 1000# + posIn(unitOf(i))) / 10000000#
        Else
            keys(i, 1) = CDbl(i)
        End If
    Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    lo.ListColumns(cKey).DataBodyRange.Value = keys
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns(cKey).DataBodyRange, SortOn:=0, Order:=1
        .Header = 1
        .Apply
    End With
    lo.ListColumns(cKey).DataBodyRange.ClearContents
    Application.EnableEvents = True
    SX_RepairFormulas
    Application.ScreenUpdating = True
    Application.StatusBar = U("^0429^0438^0442 ^0443^043F^043E^0440^044F^0434^043E^0447^0435^043D: ") & board
End Sub

' gray (auto) phases of the board become blue (manual): they no longer change after edits
Public Sub SX_FreezePhases(Optional ByVal board As String = "")
    Dim lo As ListObject, cB As Long, cP As Long, i As Long, c As Range, n As Long
    Set lo = SxTable()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If board = "" Then board = SxBoardOf(lo)
    If board = "" Then Exit Sub
    cB = SxCol(lo, U("^0429^0438^0442")): cP = SxCol(lo, U("^0424^0430^0437^0430"))
    If cP = 0 Then Exit Sub
    Application.EnableEvents = False
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, cB).Value) = board Or board = "*" Then
            Set c = lo.DataBodyRange.Cells(i, cP)
            If c.HasFormula Then
                If CStr(c.Value) <> "" Then
                    c.Value = CStr(c.Value)
                    n = n + 1
                End If
            End If
        End If
    Next
    Application.EnableEvents = True
    Application.StatusBar = U("^0424^0430^0437^044B ^0437^0430^043A^0440^0435^043F^043B^0435^043D^044B, ^0441^0442^0440^043E^043A: ") & n
End Sub

Private Function SxSnapSheet() As Worksheet
    On Error Resume Next
    Set SxSnapSheet = ThisWorkbook.Worksheets(U("_^0432^044B^0434^0430^0447^0430"))
End Function

' snapshot of the issue: key fields of every line -> hidden sheet; sheet "Changes" compares with it
Public Sub SX_Snapshot(Optional ByVal rev As String = "")
    Dim lo As ListObject, ws As Worksheet, v As Variant, i As Long, j As Long, n As Long, m As Long, outA() As Variant, keysA() As Variant, cols As Variant, idx() As Long
    Set lo = SxTable()
    Set ws = SxSnapSheet()
    If lo Is Nothing Or ws Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If rev = "" Then
        If Application.Visible Then
            rev = InputBox(U("^041D^043E^043C^0435^0440 ^0438^0437^043C^0435^043D^0435^043D^0438^044F (0 ^2014 ^043F^0435^0440^0432^0430^044F ^0432^044B^0434^0430^0447^0430):"), U("^041E^0434^043D^043E^043B^0438^043D^0435^0439^043A^0430 v5"), IIf(CStr(ws.Range("B1").Value) = "", "0", CStr(Val(ws.Range("B1").Value) + 1)))
        Else
            rev = IIf(CStr(ws.Range("B1").Value) = "", "0", CStr(Val(ws.Range("B1").Value) + 1))
        End If
    End If
    If Trim$(rev) = "" Then Exit Sub
    If Application.Visible Then
        If MsgBox(U("^0417^0430^043A^0440^0435^043F^0438^0442^044C ^0444^0430^0437^044B ^0432^0441^0435^0445 ^0449^0438^0442^043E^0432? ^0422^043E^0433^0434^0430 ^043F^0440^0430^0432^043A^0438 ^043F^043E^0441^043B^0435 ^0432^044B^0434^0430^0447^0438 ^043D^0435 ^043F^0435^0440^0435^0441^0442^0430^0432^044F^0442 ^0444^0430^0437^044B."), vbYesNo + vbQuestion, U("^041E^0434^043D^043E^043B^0438^043D^0435^0439^043A^0430 v5")) = vbYes Then SX_FreezePhases "*"
    End If
    cols = Array(U("^0429^0438^0442"), U("^041D^043E^043C^0435^0440 ^043B^0438^043D^0438^0438"), U("^041D^0430^0438^043C^0435^043D^043E^0432^0430^043D^0438^0435"), U("^0420, ^043A^0412^0442"), U("^0410^0432^0442^043E^043C^0430^0442"), U("^0421^0435^0447^0435^043D^0438^0435, ^043C^043C^00B2"), U("^0414^043B^0438^043D^0430, ^043C"), U("^0424^0430^0437^0430"), U("^041A^0430^0431^0435^043B^044C"), U("_^041C^043E^0434^0443^043B^044C (^0441^0445^0435^043C^0430)"), U("^041A^0430^043D^0430^043B (^0438^0442^043E^0433)"), U("^041A^043B^0435^043C^043C^0430"), U("^041B^0438^0441^0442 / ^043C^0435^0441^0442^043E"), U("^041B^0438^0441^0442 ^0441^0445^0435^043C^044B"))
    m = UBound(cols) + 1
    ReDim idx(1 To m)
    For j = 1 To m: idx(j) = SxCol(lo, CStr(cols(j - 1))): Next
    v = lo.DataBodyRange.Value
    ReDim outA(1 To UBound(v, 1), 1 To m)
    ReDim keysA(1 To UBound(v, 1), 1 To 1)
    For i = 1 To UBound(v, 1)
        If Trim$(CStr(v(i, idx(2)))) <> "" Then
            n = n + 1
            keysA(n, 1) = CStr(v(i, idx(1))) & "|" & CStr(v(i, idx(2)))
            For j = 1 To m
                If idx(j) > 0 Then
                    If IsError(v(i, idx(j))) Then outA(n, j) = "" Else outA(n, j) = v(i, idx(j))
                End If
            Next
        End If
    Next
    Application.EnableEvents = False
    ws.Range(ws.Cells(5, 1), ws.Cells(20004, 15)).ClearContents
    If n > 0 Then
        ws.Range(ws.Cells(5, 1), ws.Cells(4 + n, m)).Value = outA
        ws.Range(ws.Cells(5, 15), ws.Cells(4 + n, 15)).Value = keysA
    End If
    ws.Range("B1").Value = rev
    ws.Range("B2").Value = Date
    ws.Cells(ws.Rows.Count, 16).End(-4162).Offset(1, 0).Value = rev & " | " & Format$(Date, "dd.mm.yyyy") & " | " & n
    Application.EnableEvents = True
    Application.Calculate
    If Application.Visible Then MsgBox U("^0412^044B^0434^0430^0447^0430 ^0437^0430^0444^0438^043A^0441^0438^0440^043E^0432^0430^043D^0430, ^043B^0438^043D^0438^0439: ") & n, vbInformation
End Sub

Private Function SxModWidth(ByVal model As String, ByVal dflt As Double) As Double
    Dim r As Variant, w As Variant
    SxModWidth = dflt
    If Trim$(model) = "" Then Exit Function
    On Error Resume Next
    r = Application.Match(model, ThisWorkbook.Worksheets(U("^041A^0430^0442^0430^043B^043E^0433")).Range("C4:C500"), 0)
    If Not IsError(r) Then
        w = ThisWorkbook.Worksheets(U("^041A^0430^0442^0430^043B^043E^0433")).Range("L4:L500").Cells(r, 1).Value
        If IsNumeric(w) Then If w > 0 Then SxModWidth = w
    End If
End Function

Private Function SxPoles(ByVal s As String) As Double
    s = UCase$(s)
    If InStr(s, "4P") > 0 Then
        SxPoles = 4
    ElseIf InStr(s, "3P") > 0 Then
        SxPoles = 3
    ElseIf InStr(s, "2P") > 0 Or InStr(s, "1P+N") > 0 Then
        SxPoles = 2
    Else
        SxPoles = 1
    End If
End Function

' caption row above a label strip (rail number)
Private Sub SxLblCap(ByVal ws As Worksheet, ByVal r As Long, ByVal s As String)
    ws.Cells(r, 1).Value = s
    ws.Cells(r, 1).Font.Size = 6.5
    ws.Cells(r, 1).Font.Color = 8421504
    ws.Rows(r).RowHeight = 12
End Sub

' font for a label: 7 pt; 6 pt when the text would take more than 4 lines, 5.5 pt - more than 5 (about 13 characters per module and line)
Private Function SxLblFont(ByVal s As String, ByVal w As Double) As Double
    Dim p As Variant, i As Long, n As Long, per As Double
    per = 13 * IIf(w < 1, 1, w)
    p = Split(s, vbLf)
    For i = 0 To UBound(p)
        n = n + IIf(Len(CStr(p(i))) = 0, 1, -Int(-Len(CStr(p(i))) / per))
    Next
    If n > 5 Then
        SxLblFont = 5.5
    ElseIf n > 4 Then
        SxLblFont = 6
    Else
        SxLblFont = 7
    End If
End Function

' labels for the board: strip on A4, one cell = 1 DIN module (18 mm on paper); designation + line + name under every device.
' Excel prints columns wider than on screen (printer font metrics): the factor F = printed / screen width is taken
' on one wide column (page layout view Width = printed width; rounding to screen pixels is negligible there);
' widths go in screen-pixel steps, so the strip mixes the two nearest screen widths around 18 mm / F.
' A rail longer than the sheet goes on several strips; a caption above each strip names the rail.
Public Sub SX_Labels(Optional ByVal board As String = "")
    Dim lo As ListObject, ws As Worksheet, v As Variant, i As Long, j As Long, n As Long, k As Long, perRail As Double
    Dim cB As Long, cL As Long, cPos As Long, cSt As Long, cQF As Long, cH As Long, cD(1 To 3) As Long, cMS As Long, cMM As Long, cNm As Long, cMode As Long
    Dim items() As Variant, ni As Long, w As Double, col As Long, row As Long, used As Double, txt As String, cw As Double, a As String
    Dim perLine As Double, lu As Double, rail As Long, tgt As Double, x As Double, cwLo As Double, cwHi As Double
    Dim wLo As Double, wHi As Double, acc As Double, vw As Long, nCols As Long, c As Long
    Dim fk As Double, tn As Double, wN As Double, wP As Double
    Set lo = SxTable()
    If lo Is Nothing Then Exit Sub
    If board = "" Then board = Trim$(CStr(ThisWorkbook.Worksheets(U("^0429^0438^0442")).Range("C3").Value))
    If board = "" Then Exit Sub
    cB = SxCol(lo, U("^0429^0438^0442")): cL = SxCol(lo, U("^041D^043E^043C^0435^0440 ^043B^0438^043D^0438^0438")): cPos = SxCol(lo, U("_^2116 ^043C^0435^0441^0442^0430 (^0438^0442^043E^0433)")): cSt = SxCol(lo, U("_^0441^0432^043E^0439 ^0430^0432^0442^043E^043C^0430^0442"))
    cQF = SxCol(lo, U("_QF ^0433^0440.")): cH = SxCol(lo, U("^0410^0432^0442^043E^043C^0430^0442 (^0441^0445^0435^043C^0430)")): cMS = SxCol(lo, U("_^043C^043E^0434^0443^043B^044C ^043D^0430^0447.")): cMM = SxCol(lo, U("_^043C^043E^0434^0435^043B^044C ^0433^0440."))
    cNm = SxCol(lo, U("^041D^0430^0438^043C^0435^043D^043E^0432^0430^043D^0438^0435 (^0438^0442^043E^0433)")): cMode = SxCol(lo, U("^0420^0435^0436^0438^043C"))
    cD(1) = SxCol(lo, U("^0410^043F^043F^0430^0440^0430^0442 1")): cD(2) = SxCol(lo, U("^0410^043F^043F^0430^0440^0430^0442 2")): cD(3) = SxCol(lo, U("^0410^043F^043F^0430^0440^0430^0442 3"))
    v = lo.DataBodyRange.Value
    ReDim items(1 To 3, 1 To 4 * UBound(v, 1) + 1)
    For i = 1 To UBound(v, 1)
        If CStr(v(i, cB)) = board And IsNumeric(v(i, cPos)) And CStr(v(i, cPos)) <> "" Then
            If Val(CStr(v(i, cSt))) = 1 And CStr(v(i, cH)) <> "" Then
                ni = ni + 1
                items(1, ni) = SxPoles(CStr(v(i, cH)))
                a = CStr(v(i, cQF)): If a = "" Then a = CStr(v(i, cH))
                items(2, ni) = a & vbLf & CStr(v(i, cL)) & vbLf & CStr(v(i, cNm))
            End If
            If Val(CStr(v(i, cSt))) = 1 Then
                For k = 1 To 3
                    If cD(k) > 0 Then
                        If Trim$(CStr(v(i, cD(k)))) <> "" Then
                            ni = ni + 1
                            a = Trim$(CStr(v(i, cD(k))))
                            If Left$(UCase$(a), 3) = U("^0423^0417^041E") Then items(1, ni) = SxPoles(a) Else items(1, ni) = SxModWidth(a, 1)
                            items(2, ni) = a & vbLf & CStr(v(i, cL))
                        End If
                    End If
                Next
            End If
            If cMS > 0 Then
                If CStr(v(i, cMS)) <> "" Then
                    ni = ni + 1
                    items(1, ni) = SxModWidth(CStr(v(i, cMM)), 4)
                    items(2, ni) = CStr(v(i, cMS)) & vbLf & CStr(v(i, cMM))
                End If
            End If
        End If
    Next
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(U("^041D^0430^043A^043B^0435^0439^043A^0438"))
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    perRail = Val(CStr(ws.Range("C2").Value))
    If perRail <= 0 Then perRail = 24
    ws.Range(ws.Cells(3, 1), ws.Cells(400, 80)).UnMerge
    ws.Range(ws.Cells(3, 1), ws.Cells(400, 80)).Clear
    ws.Range("A1").Value = U("^041D^0410^041A^041B^0415^0419^041A^0418 ^0449^0438^0442^0430 ") & board
    On Error Resume Next
    With ws.PageSetup
        .PaperSize = 9
        .Orientation = 2
        .Zoom = 100
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .CenterHorizontally = False
    End With
    On Error GoTo 0
    tgt = Application.CentimetersToPoints(1.8)
    perLine = Int((Application.CentimetersToPoints(29.7) - ws.PageSetup.LeftMargin - ws.PageSetup.RightMargin) / tgt)
    If perLine < 1 Then perLine = 1
    nCols = perLine
    If perRail < nCols Then nCols = perRail
    ' width: F on a wide column, then screen widths around tgt / F
    vw = 1
    fk = 1
    On Error Resume Next
    ThisWorkbook.Activate
    ws.Activate
    vw = ActiveWindow.View
    ActiveWindow.View = 1
    ws.Columns(1).ColumnWidth = 100
    wN = ws.Columns(1).Width
    ActiveWindow.View = 3
    wP = ws.Columns(1).Width
    ActiveWindow.View = 1
    On Error GoTo 0
    If wN > 0 And wP > 0 Then fk = wP / wN
    If fk < 0.8 Or fk > 1.25 Then fk = 1
    tn = tgt / fk
    x = 5
    Do While x < 15
        ws.Columns(1).ColumnWidth = x
        w = ws.Columns(1).Width
        If w <= tn Then
            cwLo = x: wLo = w
        Else
            cwHi = x: wHi = w
            Exit Do
        End If
        x = x + 0.02
    Loop
    If cwHi = 0 Then cwHi = cwLo: wHi = wLo
    If cwLo = 0 Then cwLo = cwHi: wLo = wHi
    acc = 0
    For c = 1 To nCols
        If Abs(acc + wLo - c * tn) <= Abs(acc + wHi - c * tn) Then
            ws.Columns(c).ColumnWidth = cwLo: acc = acc + wLo
        Else
            ws.Columns(c).ColumnWidth = cwHi: acc = acc + wHi
        End If
    Next
    On Error Resume Next
    ActiveWindow.View = vw
    On Error GoTo 0
    Application.ScreenUpdating = False
    row = 4: rail = 1: used = 0: lu = 0: col = 1
    SxLblCap ws, row, U("^0420^0435^0439^043A^0430 ") & rail
    row = row + 1
    For k = 1 To ni
        w = items(1, k)
        If w < 1 Then w = 1
        If used + w > perRail And used > 0 Then
            rail = rail + 1: used = 0: lu = 0: col = 1
            SxLblCap ws, row + 1, U("^0420^0435^0439^043A^0430 ") & rail
            row = row + 2
        ElseIf lu + w > perLine And lu > 0 Then
            lu = 0: col = 1
            SxLblCap ws, row + 1, U("^0420^0435^0439^043A^0430 ") & rail & U(", ^043F^0440^043E^0434^043E^043B^0436^0435^043D^0438^0435")
            row = row + 2
        End If
        With ws.Range(ws.Cells(row, col), ws.Cells(row, col + CLng(w) - 1))
            If w > 1 Then .Merge
            .Value = items(2, k)
            .WrapText = True
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4108
            .Font.Size = SxLblFont(CStr(items(2, k)), w)
            .Borders.LineStyle = 1
        End With
        ws.Rows(row).RowHeight = 48
        col = col + CLng(w): used = used + w: lu = lu + w
    Next
    On Error Resume Next
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(row, nCols)).Address
    On Error GoTo 0
    Application.ScreenUpdating = True
    ws.Activate
    Application.StatusBar = U("^041D^0430^043A^043B^0435^0435^043A: ") & ni
End Sub
