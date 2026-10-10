Attribute VB_Name = "SX_Addin"
' Надстройка «Однолинейка» для таблицы 1.08: спецификация, выдача и изменения, бирки, наклейки.
' Работает с активной книгой (листы «Исходные данные», «Разбивка_по_щитам», «Спецификация»).
Option Explicit

Private Const RZ As String = "Разбивка_по_щитам"
Private Const SRC As String = "Исходные данные"
Private Const SPEC As String = "Спецификация"
Private Const SNAP As String = "_снимок"
Private Const CHG As String = "Изменения"
Private Const R0 As Long = 3
Private Const R1 As Long = 1000
Private Const CLR_CHG As Long = 10284031   ' RGB(255,235,156) — изменено
Private Const CLR_NEW As Long = 13561798   ' RGB(198,239,206) — новое
Public SX_Quiet As Boolean                  ' True — без окон сообщений (для проверки программой)
Public SX_LastMsg As String
Private CX As Object      ' заголовок «Разбивки» (строка 2) -> номер столбца

Public Function SX_GetLastMsg() As String
    SX_GetLastMsg = SX_LastMsg
End Function

Public Sub SX_SetQuiet(Optional q As Boolean = True)
    SX_Quiet = q
End Sub

Private Sub Say(txt As String, Optional st As VbMsgBoxStyle = vbInformation)
    SX_LastMsg = txt
    If SX_Quiet Then
        Debug.Print txt
    Else
        MsgBox txt, st, "Однолинейка"
    End If
End Sub

' ---------------------------------------------------------------- лента
Public Sub SX_Ribbon(control As Object)
    Select Case control.ID
        Case "sxSpec": SX_UpdateSpec
        Case "sxFix": SX_FixIssue
        Case "sxChanges": SX_ShowChanges
        Case "sxTags": SX_Tags
        Case "sxLabels": SX_Labels
        Case "sxHelp": SX_Help
    End Select
End Sub

 ---------------------------------------------------------------- общее

' прочитать «Разбивку» целиком; столбцы — по заголовкам (порядок столбцов может меняться)
Private Function LoadRz(wb As Workbook) As Variant
    Dim h As Variant, j As Long, k As String, lastC As Long
    Set CX = CreateObject("Scripting.Dictionary")
    CX.CompareMode = vbBinaryCompare
    h = wb.Sheets(RZ).Range("A2:EZ2").Value
    For j = 1 To UBound(h, 2)
        k = S(h(1, j))
        If k <> "" Then
            If Not CX.Exists(k) Then CX.Add k, j
            lastC = j
        End If
    Next j
    LoadRz = wb.Sheets(RZ).Range(wb.Sheets(RZ).Cells(R0, 1), wb.Sheets(RZ).Cells(R1, lastC)).Value
End Function

Private Function Kol(nm As String) As Long
    If CX Is Nothing Then Err.Raise 1004, , "столбцы «Разбивки» не прочитаны"
    If Not CX.Exists(nm) Then Err.Raise 1004, , "в «Разбивке» нет столбца «" & nm & "»"
    Kol = CX(nm)
End Function

Private Function Book() As Workbook
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then GoTo bad
    If Not HasSheet(wb, RZ) Or Not HasSheet(wb, SRC) Then GoTo bad
    Set Book = wb
    Exit Function
bad:
    Say "Откройте таблицу однолинейки 1.08 (листы «Исходные данные» и «Разбивка_по_щитам»).", vbExclamation
    Set Book = Nothing
End Function

Private Function HasSheet(wb As Workbook, nm As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Sheets(nm)
    HasSheet = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function S(v As Variant) As String
    If IsError(v) Then
        S = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        S = ""
    Else
        S = Trim$(CStr(v))
    End If
End Function

Private Function NumV(v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsNumeric(v) And Not IsEmpty(v) Then NumV = CDbl(v)
End Function

Private Function NameVal(wb As Workbook, nm As String, dflt As Variant) As Variant
    On Error GoTo no
    NameVal = wb.Names(nm).RefersToRange.Value
    If IsEmpty(NameVal) Or IsError(NameVal) Then NameVal = dflt
    Exit Function
no:
    NameVal = dflt
End Function

Private Function RoundUpTo(x As Double, stp As Double) As Double
    If stp <= 0 Then RoundUpTo = x: Exit Function
    RoundUpTo = -Int(-x / stp) * stp
End Function

Private Sub Fast(on_ As Boolean)
    Application.ScreenUpdating = Not on_
    Application.EnableEvents = Not on_
    If on_ Then Application.Calculation = xlCalculationManual Else Application.Calculation = xlCalculationAutomatic
End Sub

' строки исходных: номер линии -> индекс в массиве
Private Function SrcIndex(asrc As Variant) As Object
    Dim d As Object, i As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(asrc, 1)
        k = S(asrc(i, 4))
        If k <> "" Then If Not d.Exists(k) Then d.Add k, i
    Next i
    Set SrcIndex = d
End Function

Private Sub AddCount(d As Object, k As String, n As Double)
    If k = "" Then Exit Sub
    If d.Exists(k) Then d(k) = d(k) + n Else d.Add k, n
End Sub

Private Function Lookup2(wb As Workbook, rngName As String, key As String, col As Long) As String
    Dim r As Range, v As Variant
    On Error GoTo no
    Set r = wb.Names(rngName).RefersToRange
    v = Application.VLookup(key, r, col, False)
    If Not IsError(v) Then Lookup2 = S(v)
no:
End Function

' ---------------------------------------------------------------- спецификация
Public Sub SX_UpdateSpec()
    Dim wb As Workbook: Set wb = Book(): If wb Is Nothing Then Exit Sub
    Dim ws As Worksheet, arz As Variant, asrc As Variant, idx As Object, board As String
    Dim i As Long, j As Long, k As String, pos As Long, sec As Long, rowOut As Long
    Dim autos As Object, devs As Object, mods As Object, modSeen As Object, x3 As Double, xp As Double
    Dim cab As Object, gof As Object, clips As Object, manual As Collection
    Dim zap As Double, rndM As Double, kcl As Double, si As Long, cores As Double, n3 As Double
    Dim parts() As String, item As Variant, nm As String, t As String

    Set ws = wb.Sheets(SPEC)
    board = S(ws.Range("H1").Value)
    zap = NumV(NameVal(wb, "Запас_кабеля", 0.2))
    rndM = NumV(NameVal(wb, "Округл_м", 10))
    kcl = NumV(NameVal(wb, "Клипсы_на_м", 2))
    Fast True
    On Error GoTo fin
    wb.Application.Calculate
    arz = LoadRz(wb)
    asrc = wb.Sheets(SRC).Range("A" & R0 & ":BH" & R1).Value
    Set idx = SrcIndex(asrc)
    Set autos = CreateObject("Scripting.Dictionary")
    Set devs = CreateObject("Scripting.Dictionary")
    Set mods = CreateObject("Scripting.Dictionary")
    Set modSeen = CreateObject("Scripting.Dictionary")
    Set cab = CreateObject("Scripting.Dictionary")
    Set gof = CreateObject("Scripting.Dictionary")
    Set clips = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(arz, 1)
        If S(arz(i, Kol("Щит"))) <> "" And (board = "" Or S(arz(i, Kol("Щит"))) = board) Then
            ' аппараты: автомат и «Аппараты» — у строки начала группы
            If NumV(arz(i, Kol("_начало группы"))) = 1 Then
                AddCount autos, S(arz(i, Kol("Автомат (итог)"))), 1
                t = Replace(S(arz(i, Kol("Аппараты"))), "; ", ";")
                If t <> "" Then
                    parts = Split(t, ";")
                    For Each item In parts
                        AddCount devs, Trim$(CStr(item)), 1
                    Next item
                End If
            End If
            ' модули: по одному на имя (R.1, Di.1...)
            If S(arz(i, Kol("Модуль (итог)"))) <> "" Then
                k = S(arz(i, Kol("Щит"))) & "|" & S(arz(i, Kol("Модуль (итог)")))
                If Not modSeen.Exists(k) Then modSeen.Add k, 1: AddCount mods, S(arz(i, Kol("Модель"))), 1
            End If
            ' клеммы и кабель — у строк с линией
            If NumV(arz(i, Kol("_есть линия"))) = 1 And idx.Exists(S(arz(i, Kol("Номер линии")))) Then
                si = idx(S(arz(i, Kol("Номер линии"))))
                If S(arz(i, Kol("Клемма (итог)"))) <> "" Then
                    cores = NumV(asrc(si, 55)): If cores = 0 Then cores = 3
                    n3 = 1
                    x3 = x3 + n3
                    If cores > 3 Then xp = xp + cores - 3
                End If
                t = S(asrc(si, 31))                          ' кабель (марка, жилы, сечение)
                If t <> "" And NumV(asrc(si, 13)) > 0 Then AddCount cab, t, NumV(asrc(si, 13))
                nm = Lookup2(wb, "Кабель_диаметр", t, 4)    ' диаметр гофры ПВХ
                If NumV(asrc(si, 11)) > 0 Then AddCount gof, "ПВХ|" & IIf(nm = "", ChrW(216) & "?", nm), NumV(asrc(si, 11))
                nm = Lookup2(wb, "Кабель_диаметр", t, 5)    ' диаметр гофры ПНД
                If NumV(asrc(si, 12)) > 0 Then AddCount gof, "ПНД|" & IIf(nm = "", ChrW(216) & "?", nm), NumV(asrc(si, 12))
            End If
        End If
    Next i

    ' ручные строки (Примечание = «вручную») сохраняем
    Set manual = New Collection
    Dim old As Variant
    old = ws.Range("B3:J1000").Value
    For i = 1 To UBound(old, 1)
        If LCase$(S(old(i, 9))) = "вручную" Then manual.Add Array(old(i, 2), old(i, 3), old(i, 4), old(i, 5), old(i, 6), old(i, 7), old(i, 8), old(i, 9))
    Next i
    ws.Range("B3:J1000").ClearContents
    ws.Range("B3:J1000").Font.Bold = False
    rowOut = 3: sec = 0

    sec = sec + 1: pos = 0
    PutHead ws, rowOut, sec, "Электрооборудование щита" & IIf(board = "", "", " " & board) & " (по однолинейной схеме)"
    For Each item In SortedKeys(autos)
        pos = pos + 1
        nm = IIf(InStr(1, CStr(item), "мА", vbTextCompare) > 0, "Автоматический выключатель дифференциального тока", "Автоматический выключатель")
        PutRow ws, rowOut, sec & "." & pos, nm, CStr(item), "", "", "шт.", autos(item), ""
    Next item
    For Each item In SortedKeys(devs)
        pos = pos + 1
        PutRow ws, rowOut, sec & "." & pos, CStr(item), CStr(item), "", "", "шт.", devs(item), ""
    Next item
    For Each item In SortedKeys(mods)
        pos = pos + 1
        nm = Lookup2(wb, "Модули_спр", CStr(item), 4)
        PutRow ws, rowOut, sec & "." & pos, "Модуль " & IIf(nm = "", "", nm & " ") & CStr(item), CStr(item), "", "", "шт.", mods(item), ""
    Next item
    If x3 > 0 Then pos = pos + 1: PutRow ws, rowOut, sec & "." & pos, "Клемма трёхуровневая (L, N, PE)", "", "", "", "шт.", x3, ""
    If xp > 0 Then pos = pos + 1: PutRow ws, rowOut, sec & "." & pos, "Клемма проходная", "", "", "", "шт.", xp, ""

    If cab.Count > 0 Then
        sec = sec + 1: pos = 0
        PutHead ws, rowOut, sec, "Кабельная продукция (длина по журналу + " & Format(zap * 100, "0") & " %, вверх до " & rndM & " м)"
        For Each item In SortedKeys(cab)
            pos = pos + 1
            PutRow ws, rowOut, sec & "." & pos, "Кабель " & CStr(item), CStr(item), Lookup2(wb, "Справочная_ПТО", CStr(item), 2), "", "м", RoundUpTo(cab(item) * (1 + zap), rndM), ""
        Next item
    End If
    If gof.Count > 0 Then
        sec = sec + 1: pos = 0
        PutHead ws, rowOut, sec, "Трубы гофрированные и крепёж"
        For Each item In SortedKeys(gof)
            parts = Split(CStr(item), "|")
            pos = pos + 1
            PutRow ws, rowOut, sec & "." & pos, "Труба гофрированная " & parts(0) & " " & parts(1), parts(1), "", "", "м", RoundUpTo(gof(item), rndM), ""
            AddCount clips, parts(1), gof(item) * kcl
        Next item
        For Each item In SortedKeys(clips)
            pos = pos + 1
            PutRow ws, rowOut, sec & "." & pos, "Держатель (клипса) для трубы " & CStr(item), CStr(item), "", "", "шт.", RoundUpTo(clips(item), 1), "норма " & kcl & " шт/м"
        Next item
    End If
    If manual.Count > 0 Then
        sec = sec + 1: pos = 0
        PutHead ws, rowOut, sec, "Корпус и прочее (вписано вручную)"
        For Each item In manual
            pos = pos + 1
            ws.Range("B" & rowOut).Value = sec & "." & pos
            ws.Range("C" & rowOut & ":J" & rowOut).Value = Array(item(0), item(1), item(2), item(3), item(4), item(5), item(6), item(7))
            rowOut = rowOut + 1
        Next item
    End If
    ws.Range("J1").Value = "Обновлено " & Format(Now, "dd.mm.yyyy hh:nn")
fin:
    Fast False
    If Err.Number <> 0 Then Say "Спецификация: " & Err.Description & " (строка " & Erl & ")", vbExclamation Else ws.Activate
End Sub

Private Sub PutHead(ws As Worksheet, ByRef r As Long, sec As Long, txt As String)
    ws.Range("B" & r).Value = CStr(sec)
    ws.Range("C" & r).Value = txt
    ws.Range("B" & r & ":C" & r).Font.Bold = True
    r = r + 1
End Sub

Private Sub PutRow(ws As Worksheet, ByRef r As Long, p As String, nm As String, tp As String, code As String, mf As String, unit As String, qty As Variant, note As String)
    ws.Range("B" & r).NumberFormat = "@"
    ws.Range("B" & r & ":J" & r).Value = Array(p, nm, tp, code, mf, unit, qty, "", note)
    r = r + 1
End Sub

Private Function SortedKeys(d As Object) As Variant
    Dim a() As String, i As Long, j As Long, t As String, k As Variant
    If d.Count = 0 Then SortedKeys = Array(): Exit Function
    ReDim a(0 To d.Count - 1)
    i = 0
    For Each k In d.Keys: a(i) = CStr(k): i = i + 1: Next k
    For i = 0 To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If StrComp(a(j), a(i), vbTextCompare) < 0 Then t = a(i): a(i) = a(j): a(j) = t
        Next j
    Next i
    SortedKeys = a
End Function

' ---------------------------------------------------------------- снимок и изменения
' ключевые поля: лист | столбец ключа | список (заголовок, столбец)
Private Function Fields() As Variant
    Fields = Array( _
        Array(SRC, 4, Array(Array("Наименование", 36), Array("Р, кВт", 5), Array("Напряжение", 6), Array("Фаз", 8), _
              Array("Длина итог, м", 13), Array("Кабель", 31), Array("Автомат (номинал)", 21), Array("Сечение", 22), _
              Array("Потолок/Пол", 19))), _
        Array(RZ, "Номер линии", Array(Array("Щит", "Щит"), Array("Место", "Место на схеме"), Array("Лист схемы", "Лист схемы"), _
              Array("QF", "QF"), Array("Автомат", "Автомат (итог)"), Array("Модуль", "Модуль (итог)"), Array("Канал", "Канал (итог)"), _
              Array("Клемма", "Клемма (итог)"), Array("Фаза", "фаза"))), _
        Array(SPEC, 0, Array(Array("Кол-во", 8))))
End Function

Private Function SnapKey(sh As String, arr As Variant, i As Long, keyCol As Variant) As String
    If sh = SPEC Then
        SnapKey = S(arr(i, 3)) & " | " & S(arr(i, 4))
        If SnapKey = " | " Then SnapKey = ""
    Else
        SnapKey = S(arr(i, ColOf(sh, keyCol)))
    End If
End Function

Private Function ColOf(sh As String, colSpec As Variant) As Long
    If VarType(colSpec) = vbString Then ColOf = Kol(CStr(colSpec)) Else ColOf = CLng(colSpec)
End Function

Private Function ReadState(wb As Workbook) As Object
    Dim d As Object, f As Variant, fl As Variant, arr As Variant, i As Long, k As String, j As Long
    Set d = CreateObject("Scripting.Dictionary")
    For Each f In Fields()
        If HasSheet(wb, CStr(f(0))) Then
            If CStr(f(0)) = RZ Then
                arr = LoadRz(wb)
            Else
                arr = wb.Sheets(CStr(f(0))).Range("A" & R0 & ":AO" & R1).Value
            End If
            For i = 1 To UBound(arr, 1)
                k = SnapKey(CStr(f(0)), arr, i, f(1))
                If k <> "" Then
                    If CStr(f(0)) <> SPEC Or (S(arr(i, 2)) <> "" And InStr(S(arr(i, 2)), ".") > 0) Then
                        For Each fl In f(2)
                            j = ColOf(CStr(f(0)), fl(1))
                            If Not d.Exists(f(0) & "¦" & k & "¦" & fl(0)) Then d.Add f(0) & "¦" & k & "¦" & fl(0), S(arr(i, j)) & "¦" & (i + R0 - 1) & "¦" & j
                        Next fl
                    End If
                End If
            Next i
        End If
    Next f
    Set ReadState = d
End Function

Public Sub SX_FixIssue()
    Dim wb As Workbook: Set wb = Book(): If wb Is Nothing Then Exit Sub
    Dim ws As Worksheet, d As Object, k As Variant, r As Long, rev As Long, data() As Variant
    Fast True
    On Error GoTo fin
    wb.Application.Calculate
    If Not HasSheet(wb, SNAP) Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = SNAP
    Else
        Set ws = wb.Sheets(SNAP)
        rev = NumV(ws.Range("B1").Value)
    End If
    ClearMarks wb
    Set d = ReadState(wb)
    ws.Cells.ClearContents
    rev = rev + 1
    ws.Range("A1:C1").Value = Array("Выдача №", rev, Format(Now, "dd.mm.yyyy hh:nn"))
    If d.Count > 0 Then
        ReDim data(1 To d.Count, 1 To 2)
        r = 0
        For Each k In d.Keys
            r = r + 1
            data(r, 1) = k
            data(r, 2) = Split(d(k), "¦")(0)
        Next k
        ws.Range("A3").Resize(d.Count, 2).NumberFormat = "@"
        ws.Range("A3").Resize(d.Count, 2).Value = data
    End If
    ws.Visible = xlSheetHidden
fin:
    Fast False
    If Err.Number <> 0 Then
        Say "Зафиксировать выдачу: " & Err.Description & " (строка " & Erl & ")", vbExclamation
    Else
        Say "Выдача № " & rev & " зафиксирована. Дальше «Показать изменения» подсветит всё, что поменяется после неё.", vbInformation
    End If
End Sub

Private Sub ClearMarks(wb As Workbook)
    Dim ws As Worksheet, a As Variant, i As Long
    If Not HasSheet(wb, SNAP) Then Exit Sub
    Set ws = wb.Sheets(SNAP)
    a = ws.Range("E3:E5000").Value
    For i = 1 To UBound(a, 1)
        If S(a(i, 1)) = "" Then Exit For
        On Error Resume Next
        With wb.Sheets(Split(a(i, 1), "!")(0)).Range(Split(a(i, 1), "!")(1)).Interior
            If .Color = CLR_CHG Or .Color = CLR_NEW Then .ColorIndex = xlColorIndexNone
        End With
        On Error GoTo 0
    Next i
    ws.Range("E3:E5000").ClearContents
End Sub

Public Sub SX_ShowChanges()
    Dim wb As Workbook: Set wb = Book(): If wb Is Nothing Then Exit Sub
    Dim ws As Worksheet, sn As Worksheet, cur As Object, old As Object, a As Variant, i As Long, k As Variant
    Dim p() As String, q() As String, out() As Variant, n As Long, marks As Collection, rev As Long, dt As String
    Dim keyLine As String, sheetNo As String, sh As String, fld As String, rowsDone As Object
    If Not HasSheet(wb, SNAP) Then
        Say "Сначала «Зафиксировать выдачу» — изменения считаются от неё.", vbInformation
        Exit Sub
    End If
    Fast True
    On Error GoTo fin
    wb.Application.Calculate
    Set sn = wb.Sheets(SNAP)
    rev = NumV(sn.Range("B1").Value): dt = S(sn.Range("C1").Value)
    ClearMarks wb
    Set old = CreateObject("Scripting.Dictionary")
    a = sn.Range("A3:B60000").Value
    For i = 1 To UBound(a, 1)
        If S(a(i, 1)) = "" Then Exit For
        old(S(a(i, 1))) = S(a(i, 2))
    Next i
    Set cur = ReadState(wb)
    Dim oldRows As Object, curRows As Object
    Set oldRows = RowSet(old): Set curRows = RowSet(cur)
    Set marks = New Collection
    Set rowsDone = CreateObject("Scripting.Dictionary")
    ReDim out(1 To cur.Count + old.Count + 1, 1 To 6)
    For Each k In cur.Keys
        p = Split(k, "¦"): q = Split(cur(k), "¦")
        If old.Exists(k) Then
            If old(k) <> q(0) Then
                n = n + 1
                out(n, 1) = p(0): out(n, 2) = p(1): out(n, 3) = p(2): out(n, 4) = old(k): out(n, 5) = q(0)
                out(n, 6) = SchemeSheet(wb, p(0), p(1))
                Mark wb, marks, p(0), CLng(q(1)), CLng(q(2)), CLR_CHG
            End If
        ElseIf Not rowsDone.Exists(p(0) & "¦" & p(1)) Then
            If Not oldRows.Exists(p(0) & "¦" & p(1)) Then
                rowsDone.Add p(0) & "¦" & p(1), 1
                n = n + 1
                out(n, 1) = p(0): out(n, 2) = p(1): out(n, 3) = "добавлено": out(n, 4) = "": out(n, 5) = q(0)
                out(n, 6) = SchemeSheet(wb, p(0), p(1))
                Mark wb, marks, p(0), CLng(q(1)), IIf(p(0) = SPEC, 3, IIf(p(0) = RZ, Kol("Номер линии"), 4)), CLR_NEW
            End If
        End If
    Next k
    For Each k In old.Keys
        p = Split(k, "¦")
        If Not curRows.Exists(p(0) & "¦" & p(1)) And Not rowsDone.Exists("-" & p(0) & "¦" & p(1)) Then
            rowsDone.Add "-" & p(0) & "¦" & p(1), 1
            n = n + 1
            out(n, 1) = p(0): out(n, 2) = p(1): out(n, 3) = "удалено": out(n, 4) = old(k): out(n, 5) = ""
        End If
    Next k
    If Not HasSheet(wb, CHG) Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = CHG
    Else
        Set ws = wb.Sheets(CHG)
    End If
    ws.Cells.Clear
    ws.Range("A1").Value = "Изменения после выдачи № " & rev & " от " & dt & "  (проверено " & Format(Now, "dd.mm.yyyy hh:nn") & ")"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2:F2").Value = Array("Лист", "Линия / позиция", "Что", "Было", "Стало", "Лист схемы")
    ws.Range("A2:F2").Font.Bold = True
    If n > 0 Then ws.Range("A3").Resize(n, 6).Value = out
    ws.Columns("A:F").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("D:E").ColumnWidth = 30
    ' адреса подсветки — чтобы снять в следующий раз
    If marks.Count > 0 Then
        For i = 1 To marks.Count
            sn.Range("E" & (i + 2)).Value = marks(i)
        Next i
    End If
fin:
    Fast False
    If Err.Number <> 0 Then
        Say "Показать изменения: " & Err.Description & " (строка " & Erl & ")", vbExclamation
    Else
        ws.Activate
        Say IIf(n = 0, "Изменений после выдачи № " & rev & " нет.", "Изменений: " & n & ". Список — лист «Изменения», ячейки подсвечены (жёлтые — изменено, зелёные — новое)."), vbInformation
    End If
End Sub

Private Function RowSet(d As Object) As Object
    Dim r As Object, k As Variant, p() As String
    Set r = CreateObject("Scripting.Dictionary")
    For Each k In d.Keys
        p = Split(k, "¦")
        If Not r.Exists(p(0) & "¦" & p(1)) Then r.Add p(0) & "¦" & p(1), 1
    Next k
    Set RowSet = r
End Function

Private Sub Mark(wb As Workbook, marks As Collection, sh As String, r As Long, c As Long, clr As Long)
    On Error Resume Next
    wb.Sheets(sh).Cells(r, c).Interior.Color = clr
    marks.Add sh & "!" & wb.Sheets(sh).Cells(r, c).Address(False, False)
End Sub

Private Function SchemeSheet(wb As Workbook, sh As String, key As String) As String
    Dim r As Variant
    If sh = SPEC Then Exit Function
    On Error GoTo no
    Dim w As Worksheet: Set w = wb.Sheets(RZ)
    If CX Is Nothing Then LoadRz wb
    r = Application.Match(key, w.Range(w.Cells(R0, Kol("Номер линии")), w.Cells(R1, Kol("Номер линии"))), 0)
    If Not IsError(r) Then SchemeSheet = S(w.Cells(R0 + r - 1, Kol("Щит")).Value) & " / лист " & S(w.Cells(R0 + r - 1, Kol("Лист схемы")).Value)
no:
End Function

' ---------------------------------------------------------------- бирки
Public Sub SX_Tags()
    Dim wb As Workbook: Set wb = Book(): If wb Is Nothing Then Exit Sub
    Dim arz As Variant, asrc As Variant, idx As Object, ws As Worksheet, i As Long, j As Long, n As Long
    Dim nodes() As String, out() As Variant, cabTxt As String, nm As String, si As Long
    Fast True
    On Error GoTo fin
    wb.Application.Calculate
    arz = LoadRz(wb)
    asrc = wb.Sheets(SRC).Range("A" & R0 & ":BH" & R1).Value
    Set idx = SrcIndex(asrc)
    ReDim out(1 To 6000, 1 To 6)
    For i = 1 To UBound(arz, 1)
        If NumV(arz(i, Kol("_есть линия"))) = 1 And S(arz(i, Kol("_цепь"))) <> "" Then
            nodes = Split(S(arz(i, Kol("_цепь"))), "|")
            cabTxt = "": nm = S(arz(i, Kol("Наименование")))
            If idx.Exists(S(arz(i, Kol("Номер линии")))) Then si = idx(S(arz(i, Kol("Номер линии")))): cabTxt = S(asrc(si, 31))
            For j = 0 To UBound(nodes)
                If n >= 5999 Then Exit For
                n = n + 1
                out(n, 1) = S(arz(i, Kol("Щит"))): out(n, 2) = n: out(n, 3) = S(arz(i, Kol("Номер линии"))): out(n, 4) = nodes(j)
                If j < UBound(nodes) Then
                    out(n, 5) = nodes(j + 1)
                Else
                    out(n, 5) = "кабель " & cabTxt & " " & ChrW(8594) & " " & nm
                End If
                out(n, 6) = nm
            Next j
        End If
    Next i
    Set ws = SheetReset(wb, "Бирки")
    ws.Range("A1").Value = "БИРКИ: строка на каждый провод (откуда - куда). Обновлено " & Format(Now, "dd.mm.yyyy hh:nn")
    ws.Range("A1").Font.Bold = True
    ws.Range("A2:F2").Value = Array("Щит", "№", "Линия", "Откуда", "Куда", "Наименование")
    ws.Range("A2:F2").Font.Bold = True
    If n > 0 Then ws.Range("A3").Resize(n, 6).Value = out
    ws.Columns("A:C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 45
    ws.Columns("F").ColumnWidth = 45
fin:
    Fast False
    If Err.Number <> 0 Then Say "Бирки: " & Err.Description & " (строка " & Erl & ")", vbExclamation Else ws.Activate
End Sub

Private Function SheetReset(wb As Workbook, nm As String) As Worksheet
    Dim ws As Worksheet
    If HasSheet(wb, nm) Then
        Set ws = wb.Sheets(nm)
        ws.Cells.Clear
        On Error Resume Next
        ws.Cells.UnMerge
        On Error GoTo 0
    Else
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = nm
    End If
    Set SheetReset = ws
End Function

' ---------------------------------------------------------------- наклейки
Private Function DevWidth(t As String) As Long
    Dim u As String: u = UCase$(t)
    If Left$(u, 2) = "4P" Then DevWidth = 4: Exit Function
    If Left$(u, 2) = "3P" Then DevWidth = 3: Exit Function
    If Left$(u, 2) = "2P" Or InStr(u, "1P+N") > 0 Then DevWidth = 2: Exit Function
    If Left$(u, 2) = "1P" Then DevWidth = 1: Exit Function
    If Left$(u, 2) = "БП" Then DevWidth = 4: Exit Function
    If InStr(u, "УЗО") > 0 Or InStr(u, "УЗИП") > 0 Or InStr(u, "РЕЛЕ НАПР") > 0 Then DevWidth = 2: Exit Function
    DevWidth = 1
End Function

Public Sub SX_Labels()
    Dim wb As Workbook: Set wb = Book(): If wb Is Nothing Then Exit Sub
    Dim arz As Variant, ws As Worksheet, i As Long, items As Collection, it As Variant, curBoard As String
    Dim r As Long, c As Long, w As Long, perRow As Long, t As String, parts() As String, x As Variant, nm As String
    Dim modSeen As Object, wmod As Long
    perRow = 10
    Fast True
    On Error GoTo fin
    wb.Application.Calculate
    arz = LoadRz(wb)
    Set ws = SheetReset(wb, "Наклейки")
    ws.Columns("A:J").ColumnWidth = 9.14          ' ~ 18 мм: 1 модуль DIN (проверьте пробной печатью)
    ws.Range("A1").Value = "НАКЛЕЙКИ под аппаратами: клетка = 1 модуль DIN (18 мм). Обновлено " & Format(Now, "dd.mm.yyyy hh:nn")
    ws.Range("A1").Font.Bold = True
    r = 2: c = 0: curBoard = ""
    Set modSeen = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(arz, 1)
        If S(arz(i, Kol("Щит"))) <> "" And NumV(arz(i, Kol("_начало группы"))) = 1 And S(arz(i, Kol("Автомат (итог)"))) <> "" And S(arz(i, Kol("Зона"))) <> "вводной узел" Then
            If S(arz(i, Kol("Щит"))) <> curBoard Then
                curBoard = S(arz(i, Kol("Щит")))
                r = r + 2: c = 0
                ws.Cells(r, 1).Value = "Щит " & curBoard
                ws.Cells(r, 1).Font.Bold = True
                r = r + 1
            End If
            nm = S(arz(i, Kol("Наименование")))
            PutLabel ws, r, c, perRow, DevWidth(S(arz(i, Kol("Автомат (итог)")))), S(arz(i, Kol("QF"))) & vbLf & nm
            t = Replace(S(arz(i, Kol("Аппараты"))), "; ", ";")
            If t <> "" Then
                parts = Split(t, ";")
                For Each x In parts
                    PutLabel ws, r, c, perRow, DevWidth(CStr(x)), Trim$(CStr(x))
                Next x
            End If
        End If
        If S(arz(i, Kol("Модуль (итог)"))) <> "" Then
            If Not modSeen.Exists(S(arz(i, Kol("Щит"))) & "|" & S(arz(i, Kol("Модуль (итог)")))) Then
                modSeen.Add S(arz(i, Kol("Щит"))) & "|" & S(arz(i, Kol("Модуль (итог)"))), 1
                wmod = NumV(Lookup2(wb, "Модули_спр", S(arz(i, Kol("Модель"))), 5)): If wmod = 0 Then wmod = 4
                PutLabel ws, r, c, perRow, wmod, S(arz(i, Kol("Модуль (итог)"))) & vbLf & S(arz(i, Kol("Модель")))
            End If
        End If
    Next i
    ws.PageSetup.Orientation = xlPortrait
    ws.PageSetup.Zoom = 100
fin:
    Fast False
    If Err.Number <> 0 Then Say "Наклейки: " & Err.Description & " (строка " & Erl & ")", vbExclamation Else ws.Activate
End Sub

Private Sub PutLabel(ws As Worksheet, ByRef r As Long, ByRef c As Long, perRow As Long, w As Long, txt As String)
    Dim rg As Range
    If w > perRow Then w = perRow
    If c + w > perRow Then r = r + 1: c = 0
    Set rg = ws.Range(ws.Cells(r, c + 1), ws.Cells(r, c + w))
    If w > 1 Then rg.Merge
    rg.Value = txt
    rg.WrapText = True
    rg.Font.Size = 7
    rg.HorizontalAlignment = xlCenter
    rg.VerticalAlignment = xlCenter
    rg.Borders.LineStyle = xlContinuous
    ws.Rows(r).RowHeight = 34
    c = c + w
End Sub

' ---------------------------------------------------------------- справка
Public Sub SX_Help()
    MsgBox "Однолинейка 1.08 — вкладка на ленте:" & vbLf & vbLf & _
        "• Обновить спецификацию — лист «Спецификация» по щиту из H1 (пусто — все): аппараты по схеме, кабели, гофра, клипсы. " & _
        "Строки с «вручную» в «Примечании» сохраняются." & vbLf & _
        "• Зафиксировать выдачу — запомнить состояние (выдача № N)." & vbLf & _
        "• Показать изменения — что поменялось после выдачи: лист «Изменения» + подсветка." & vbLf & _
        "• Бирки — строка на каждый провод." & vbLf & _
        "• Наклейки — лента под аппаратами, клетка = 1 модуль DIN.", vbInformation, "Однолинейка"
End Sub
