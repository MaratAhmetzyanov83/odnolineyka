Attribute VB_Name = "SX_Ribbon"
' Вкладка «Однолинейка» на ленте для таблицы v3.12: кнопки к SX_Spec и SX_Groups, бирки, наклейки.
' Работает с активной книгой (ActiveWorkbook), в самой книге макросов нет.
Option Explicit

Private Const ODN As String = "Однолинейка"
Private Const R0 As Long = 3
Private Const R1 As Long = 1000
Private Const PWD As String = "sx"
Public SX_Quiet As Boolean                  ' True — без окон сообщений (для проверки программой)
Public SX_LastMsg As String
Private gRoot As String                     ' папка объекта (Дерево щитов, ПЗ)

Public Sub SX_SetQuiet(Optional q As Boolean = True)
    SX_Quiet = q
    SX_SpecQuiet = q                        ' модуль SX_Spec
End Sub

Public Function SX_GetLastMsg() As String
    SX_GetLastMsg = SX_LastMsg
End Function

Private Sub Say(txt As String, Optional st As VbMsgBoxStyle = vbInformation)
    SX_LastMsg = txt
    If SX_Quiet Then
        Debug.Print txt
    Else
        MsgBox txt, st, "Однолинейка"
    End If
End Sub

' ---------------------------------------------------------------- лента
Public Sub SX_OnAction(control As Object)
    Select Case control.ID
        Case "sxSpec": SX_UpdateSpec
        Case "sxAccept": SX_AcceptChanges
        Case "sxGroup": SX_InsertGroup
        Case "sxSaveGroup": SX_SaveAsGroup
        Case "sxTags": SX_Tags
        Case "sxLabels": SX_Labels
        Case "sxLvl1": SX_Level 1
        Case "sxLvl2": SX_Level 2
        Case "sxLvl3": SX_Level 3
        Case "sxLink": SX_LinkUpstream
        Case "sxLinksUpd": SX_UpdateLinks
        Case "sxTree": SX_Tree
        Case "sxPZ": SX_PZ
        Case "sxKJ": SX_KJ
        Case "sxKJFile": SX_KJFile
        Case "sxKJStamp": SX_KJStamp
        Case "sxHelp": SX_Help
    End Select
End Sub

' ---------------------------------------------------------------- общее
Private Function Book() As Workbook
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If Not wb Is Nothing Then
        If HasSheet(wb, ODN) Then Set Book = wb: Exit Function
    End If
    Say "Откройте таблицу однолинейки v3.12 (лист «Однолинейка»).", vbExclamation
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

Private Sub Fast(on_ As Boolean)
    Application.ScreenUpdating = Not on_
    Application.EnableEvents = Not on_
End Sub

' ---------------------------------------------------------------- бирки
' лист «Бирки» по кнопке: значения из «Однолинейки», по порядку линий (BV = № линии)
' A Щит, G QF, P Клемма, Q Кабель, R Линия, S Наименование, BV № линии
Public Sub SX_Tags()
    Dim wb As Workbook: Set wb = Book(): If wb Is Nothing Then Exit Sub
    Dim a As Variant, ws As Worksheet, i As Long, k As Long, n As Long, mx As Long
    Dim idx() As Long, out() As Variant, tag As String, relock As Boolean
    Fast True
    On Error GoTo fin
    wb.Application.Calculate
    a = wb.Sheets(ODN).Range("A" & R0 & ":BV" & R1).Value
    For i = 1 To UBound(a, 1)
        If NumV(a(i, 74)) > mx Then mx = CLng(NumV(a(i, 74)))
    Next i
    If mx = 0 Then Err.Raise vbObjectError + 1, , "на листе «Однолинейка» нет линий"
    ReDim idx(1 To mx)
    For i = 1 To UBound(a, 1)
        k = CLng(NumV(a(i, 74)))
        If k > 0 Then If idx(k) = 0 Then idx(k) = i
    Next i
    ReDim out(1 To mx, 1 To 6)
    For k = 1 To mx
        i = idx(k)
        If i > 0 Then
            n = n + 1
            tag = S(a(i, 7))
            If S(a(i, 16)) <> "" Then tag = tag & "/" & S(a(i, 16))
            out(n, 1) = n
            out(n, 2) = S(a(i, 1))
            out(n, 3) = tag
            out(n, 4) = S(a(i, 18))
            out(n, 5) = S(a(i, 19))
            out(n, 6) = S(a(i, 17))
        End If
    Next k
    If wb.ProtectStructure Then
        wb.Unprotect PWD
        relock = True
    End If
    Set ws = SheetReset(wb, "Бирки")
    ws.Range("A1").Value = "БИРКИ ДЛЯ ПОДКЛЮЧЕНИЯ. Обновлено " & Format(Now, "dd.mm.yyyy hh:nn") & _
        " по листу «Однолинейка». Печатайте столбцы C и D. Изменили схему — нажмите «Бирки» снова."
    ws.Range("A1").Font.Bold = True
    ws.Range("A2:F2").Value = Array("№", "Щит", "Бирка автомат/клемма", "Бирка кабеля", "Наименование", "Кабель")
    ws.Range("A2:F2").Font.Bold = True
    ws.Range("A2:F2").Interior.Color = RGB(221, 235, 247)
    ws.Range("A3").Resize(n, 6).Value = out
    ws.Range("A2").Resize(n + 1, 6).Borders.LineStyle = xlContinuous
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 8
    ws.Columns("C:D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 40
    ws.Columns("F").ColumnWidth = 26
    ws.Range("C3:D" & n + 2).Font.Bold = True
    ws.PageSetup.PrintArea = "$A$1:$F$" & n + 2
    ws.PageSetup.PrintTitleRows = "$2:$2"
fin:
    If relock Then wb.Protect PWD, Structure:=True
    Fast False
    If Err.Number <> 0 Then
        Say "Бирки: " & Err.Description, vbExclamation
    Else
        ws.Activate
        Say "Бирки: " & n & " шт. на листе «Бирки»."
    End If
End Sub

' ---------------------------------------------------------------- наклейки
' ширина аппарата в модулях DIN: из «Каталога» (столбец «Ширина, мод. DIN»), иначе по полюсам / типу
Private Function DevWidth(t As String, cat As Object) As Long
    Dim u As String: u = UCase$(t)
    If cat.Exists(u) Then DevWidth = cat(u): Exit Function
    If InStr(u, "4P") > 0 Then DevWidth = 4: Exit Function
    If InStr(u, "3P") > 0 Then DevWidth = 3: Exit Function
    If InStr(u, "2P") > 0 Or InStr(u, "1P+N") > 0 Then DevWidth = 2: Exit Function
    If Left$(u, 2) = "БП" Or Left$(u, 2) = "T " Or Left$(u, 3) = "ТР " Then DevWidth = 4: Exit Function
    If InStr(u, "УЗО") > 0 Or InStr(u, "УЗИП") > 0 Or InStr(u, "РЕЛЕ НАПР") > 0 Then DevWidth = 2: Exit Function
    DevWidth = 1
End Function

' «Каталог»: модель (C) -> ширина, мод. DIN (L), если указана
Private Function CatWidths(wb As Workbook) As Object
    Dim d As Object, a As Variant, i As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    If HasSheet(wb, "Каталог") Then
        a = wb.Sheets("Каталог").Range("C4:L500").Value
        For i = 1 To UBound(a, 1)
            k = UCase$(S(a(i, 1)))
            If k <> "" And NumV(a(i, 10)) > 0 Then If Not d.Exists(k) Then d.Add k, CLng(NumV(a(i, 10)))
        Next i
    End If
    Set CatWidths = d
End Function

Public Sub SX_Labels()
    Dim wb As Workbook: Set wb = Book(): If wb Is Nothing Then Exit Sub
    Dim a As Variant, ws As Worksheet, i As Long, j As Long, cat As Object, curBoard As String
    Dim r As Long, c As Long, perRow As Long, t As String, nm As String, wmod As Long, n As Long
    Dim relock As Boolean
    perRow = 12
    Fast True
    On Error GoTo fin
    wb.Application.Calculate
    ' A Щит, G QF, H Автомат, I:K аппараты, M Модуль, S Наименование, Y Режим, AC Модель модуля, AJ QF нач., AW модуль нач.
    a = wb.Sheets(ODN).Range("A" & R0 & ":AW" & R1).Value
    Set cat = CatWidths(wb)
    If wb.ProtectStructure Then
        wb.Unprotect PWD
        relock = True
    End If
    Set ws = SheetReset(wb, "Наклейки")
    ws.Columns("A:L").ColumnWidth = 9.14          ' ~ 18 мм: 1 модуль DIN (проверьте пробной печатью)
    ws.Range("A1").Value = "НАКЛЕЙКИ под аппаратами: клетка = 1 модуль DIN (18 мм). Обновлено " & Format(Now, "dd.mm.yyyy hh:nn") & _
        ". Ширину аппарата можно задать в «Каталоге», столбец «Ширина, мод. DIN»."
    ws.Range("A1").Font.Bold = True
    r = 2: c = 0: curBoard = ""
    For i = 1 To UBound(a, 1)
        If S(a(i, 1)) <> "" And S(a(i, 25)) <> "вручную" Then
            If S(a(i, 1)) <> curBoard Then
                curBoard = S(a(i, 1))
                r = r + 2: c = 0
                ws.Cells(r, 1).Value = "Щит " & curBoard
                ws.Cells(r, 1).Font.Bold = True
                r = r + 1
            End If
            ' автомат: первая строка группы
            If S(a(i, 36)) <> "" And S(a(i, 8)) <> "" Then
                nm = S(a(i, 19))
                PutLabel ws, r, c, perRow, DevWidth(S(a(i, 8)), cat), S(a(i, 7)) & vbLf & nm
                n = n + 1
                For j = 9 To 11
                    t = S(a(i, j))
                    If t <> "" And t <> "-" Then PutLabel ws, r, c, perRow, DevWidth(t, cat), t: n = n + 1
                Next j
            End If
            ' модуль: первая строка модуля
            If S(a(i, 49)) <> "" Then
                wmod = 0
                If cat.Exists(UCase$(S(a(i, 29)))) Then wmod = cat(UCase$(S(a(i, 29))))
                If wmod = 0 Then wmod = 4
                PutLabel ws, r, c, perRow, wmod, S(a(i, 13)) & vbLf & S(a(i, 29))
                n = n + 1
            End If
        End If
    Next i
    ws.PageSetup.Orientation = xlPortrait
    ws.PageSetup.Zoom = 100
fin:
    If relock Then wb.Protect PWD, Structure:=True
    Fast False
    If Err.Number <> 0 Then
        Say "Наклейки: " & Err.Description & " (строка " & Erl & ")", vbExclamation
    Else
        ws.Activate
        Say "Наклейки: " & n & " шт. на листе «Наклейки»."
    End If
End Sub

Private Function SheetReset(wb As Workbook, nm As String) As Worksheet
    Dim ws As Worksheet, after As Object
    If HasSheet(wb, nm) Then
        Set ws = wb.Sheets(nm)
        ws.Cells.Clear
        On Error Resume Next
        ws.Cells.UnMerge
        On Error GoTo 0
    Else
        ' новый лист — за «Спецификацией» (и «Изменениями»), иначе в конце
        Set after = wb.Sheets(wb.Sheets.Count)
        If HasSheet(wb, "Спецификация") Then Set after = wb.Sheets("Спецификация")
        If HasSheet(wb, "Изменения") Then Set after = wb.Sheets("Изменения")
        Set ws = wb.Sheets.Add(After:=after)
        ws.Name = nm
    End If
    Set SheetReset = ws
End Function

Private Sub PutLabel(ws As Worksheet, ByRef r As Long, ByRef c As Long, perRow As Long, w As Long, txt As String)
    Dim rg As Range
    If w > perRow Then w = perRow
    If w < 1 Then w = 1
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

' ---------------------------------------------------------------- уровни столбцов
' кнопки «Ввод / Результаты / Всё»: уровни 1 / 2 / 3 на «Однолинейке» и «Исходных данных».
' Защищённый лист: защита снимается на время и ставится с теми же разрешениями.
Public Sub SX_Level(n As Integer)
    Dim ws As Worksheet, p As Protection, prot As Boolean
    Dim a(1 To 11) As Boolean
    Set ws = ActiveSheet
    If ws.Name <> ODN And ws.Name <> "Исходные данные" Then
        Say "Перейдите на лист «Однолинейка» или «Исходные данные».", vbExclamation
        Exit Sub
    End If
    prot = ws.ProtectContents
    If prot Then
        Set p = ws.Protection
        a(1) = p.AllowFormattingCells: a(2) = p.AllowFormattingColumns: a(3) = p.AllowFormattingRows
        a(4) = p.AllowInsertingColumns: a(5) = p.AllowInsertingRows: a(6) = p.AllowInsertingHyperlinks
        a(7) = p.AllowDeletingColumns: a(8) = p.AllowDeletingRows: a(9) = p.AllowSorting
        a(10) = p.AllowFiltering: a(11) = p.AllowUsingPivotTables
        ws.Unprotect PWD
    End If
    On Error Resume Next
    ws.Outline.ShowLevels ColumnLevels:=n
    On Error GoTo 0
    If prot Then
        ws.Protect PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
            AllowFormattingCells:=a(1), AllowFormattingColumns:=a(2), AllowFormattingRows:=a(3), _
            AllowInsertingColumns:=a(4), AllowInsertingRows:=a(5), AllowInsertingHyperlinks:=a(6), _
            AllowDeletingColumns:=a(7), AllowDeletingRows:=a(8), AllowSorting:=a(9), _
            AllowFiltering:=a(10), AllowUsingPivotTables:=a(11)
    End If
    SX_LastMsg = ws.Name & ": уровень " & n
End Sub

' ---------------------------------------------------------------- пути (объект в подпапках)
' В связях — путь относительно своей книги: «Корпус 1\ЩР-1.xlsx», «..\ГРЩ.xlsx»; просто имя — та же папка.
Private Function RelPath(fromFolder As String, toFull As String) As String
    Dim a As Variant, b As Variant, i As Long, j As Long, r As String
    a = Split(fromFolder, "\"): b = Split(toFull, "\")
    Do While i <= UBound(a) And i < UBound(b)
        If LCase$(a(i)) <> LCase$(b(i)) Then Exit Do
        i = i + 1
    Loop
    If i = 0 Then RelPath = toFull: Exit Function        ' другой диск
    For j = i To UBound(a)
        r = r & "..\"
    Next j
    For j = i To UBound(b)
        r = r & b(j) & IIf(j < UBound(b), "\", "")
    Next j
    RelPath = r
End Function

Private Function AbsPath(base As String, rel As String) As String
    Dim parts As Variant, st() As String, n As Long, i As Long, x As String
    If rel Like "?:\*" Or Left$(rel, 2) = "\\" Then AbsPath = rel: Exit Function
    parts = Split(base & "\" & rel, "\")
    ReDim st(0 To UBound(parts))
    For i = 0 To UBound(parts)
        x = parts(i)
        If x = ".." Then
            If n > 1 Then n = n - 1
        ElseIf x <> "." And x <> "" Then
            st(n) = x: n = n + 1
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve st(0 To n - 1)
    AbsPath = Join(st, "\")
End Function

' «путь / фидер; путь2 / фидер2» (абсолютные, строчные) -> вид относительно папки объекта
Private Function Vid(t As String) As String
    Dim ref As Variant, p As Variant, r As String
    If t = "" Then Exit Function
    For Each ref In Split(t, "; ")
        p = Split(ref & " / ", " / ")
        r = r & IIf(r = "", "", "; ") & RelPath(gRoot, CStr(p(0))) & IIf(p(1) <> "", " / " & p(1), "")
    Next ref
    Vid = r
End Function

Private Sub FindBooks(folder As String, files As Collection, Optional nameOnly As String = "")
    Dim fso As Object, fo As Object, f As Object, sf As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Exit Sub
    Set fo = fso.GetFolder(folder)
    For Each f In fo.Files
        If LCase$(f.Name) Like "*.xls[xm]" And Left$(f.Name, 2) <> "~$" And Left$(f.Name, 3) <> "ПЗ_" Then
            If nameOnly = "" Or LCase$(f.Name) = LCase$(nameOnly) Then files.Add f.Path
        End If
    Next f
    For Each sf In fo.SubFolders
        If Left$(sf.Name, 1) <> "." Then FindBooks sf.Path, files, nameOnly
    Next sf
End Sub

' папка объекта: от этой книги вверх по связям «Ввод 1» до главного щита; его папка (с подпапками)
Private Function ObjectRoot(wb As Workbook) As String
    Dim cur As String, w As Workbook, opened As Boolean, net As Worksheet, t As String, nxt As String, i As Long
    cur = wb.FullName
    For i = 1 To 12
        opened = False
        Set w = Nothing
        On Error Resume Next
        Set w = OpenBook(cur, opened)
        On Error GoTo 0
        nxt = ""
        If Not w Is Nothing Then
            Set net = NetSheet(w)
            If Not net Is Nothing Then
                t = S(net.Range("G6").Value)
                If InStr(1, t, ".xls", vbTextCompare) > 0 Then nxt = AbsPath(w.Path, CStr(Split(t & " / ", " / ")(0)))
            End If
            If opened Then w.Close False
        End If
        If nxt = "" Then Exit For
        If Dir(nxt) = "" Then Exit For
        cur = nxt
    Next i
    ObjectRoot = Left$(cur, InStrRev(cur, "\") - 1)
End Function

' ---------------------------------------------------------------- связи щитов
' Один щит — одна книга. «Сеть_щита»: раздел А, строки 6–8 (щит 1–3) — от вышестоящего: C ΔU до шин, D Zп, E Iкз max, F In,
' G файл; п. 4, строки 45–47: C:I — нагрузка этого щита для вышестоящего. Раздел Б (строки 94–123) — нижестоящие:
' B фидер, C файл, D:J — нагрузка из файла нижестоящего, K:N — что отдать нижестоящему.
' Книги до объединения листов: «Питание_щита» (те же строки) и «Нижестоящие_щиты» (строки 5–34).

Private Function NetSheet(wb As Workbook) As Worksheet
    If HasSheet(wb, "Сеть_щита") Then
        Set NetSheet = wb.Sheets("Сеть_щита")
    ElseIf HasSheet(wb, "Питание_щита") Then
        Set NetSheet = wb.Sheets("Питание_щита")
    End If
End Function

Private Function SubSheet(wb As Workbook, ByRef r0 As Long) As Worksheet
    If HasSheet(wb, "Нижестоящие_щиты") Then
        Set SubSheet = wb.Sheets("Нижестоящие_щиты"): r0 = 5
    ElseIf HasSheet(wb, "Сеть_щита") Then
        Set SubSheet = wb.Sheets("Сеть_щита"): r0 = 94
    End If
End Function

Private Function OpenBook(path As String, ByRef opened As Boolean) As Workbook
    Dim w As Workbook
    For Each w In Application.Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then Set OpenBook = w: Exit Function
    Next w
    Set OpenBook = Application.Workbooks.Open(path, 0)
    opened = True
End Function

Private Function ExtRef(wb As Workbook, ws As Worksheet) As String
    ExtRef = "'" & wb.Path & "\[" & wb.Name & "]" & ws.Name & "'!"
End Function

' Связать активную книгу (нижестоящий щит № childIdx, 1–3) с вышестоящей книгой upPath через линию-фидер feeder.
' Пишет связи в обе книги и сохраняет вышестоящую. Возвращает текст итога (он же в SX_LastMsg).
Public Function SX_LinkUpstreamTo(ByVal upPathV As Variant, ByVal feederV As Variant, _
                                  Optional ByVal childIdxV As Variant = 1, Optional ByVal inputV As Variant = 1) As String
    ' параметры Variant — так макрос можно вызвать и из Application.Run (проверка программой)
    Dim child As Workbook, up As Workbook, opened As Boolean
    Dim cNet As Worksheet, uSub As Worksheet, r0 As Long, r As Long, i As Long, j As Long, free As Long
    Dim ids As Variant, found As Boolean, refC As String, refU As String, board As String
    Dim upPath As String, feeder As String, childIdx As Long, inp As Long, nIn As Long
    Dim sRow As Long, sCol As Long, dRow As Long, dCol As Long
    upPath = CStr(upPathV): feeder = Trim$(CStr(feederV)): childIdx = CLng(Val(CStr(childIdxV)))
    inp = CLng(Val(CStr(inputV))): If inp < 1 Then inp = 1
    On Error GoTo fail
    Set child = Book(): If child Is Nothing Then Exit Function
    If child.Path = "" Then Err.Raise vbObjectError + 1, , "сохраните эту книгу в папку объекта"
    Set cNet = NetSheet(child)
    If cNet Is Nothing Then Err.Raise vbObjectError + 2, , "в этой книге нет листа «Сеть_щита»"
    If childIdx < 1 Or childIdx > 3 Then childIdx = 1
    board = S(cNet.Cells(5 + childIdx, 2).Value)
    If board = "" Then Err.Raise vbObjectError + 3, , "в этой книге нет щита № " & childIdx
    nIn = CLng(NumV(cNet.Cells(65, 2 + childIdx).Value)): If nIn < 1 Then nIn = 1
    If inp > nIn Then Err.Raise vbObjectError + 8, , "у щита «" & board & "» вводов: " & nIn & " («Сеть_щита», раздел 6)"
    ' откуда нагрузка для вышестоящего и куда питание от него
    If S(cNet.Range("U44").Value) = "ДЛЯ ВЫШЕСТОЯЩЕГО ПО ВВОДАМ" Then
        If inp > 2 Then Err.Raise vbObjectError + 9, , "нагрузка по Вводу 3 пока не считается (в «Нагрузка_щитов» две секции)"
        sRow = 46 + (childIdx - 1) * 2 + (inp - 1): sCol = 23          ' W:AC
    Else
        If inp > 1 Then Err.Raise vbObjectError + 10, , "в книге нет таблиц вводов — обновите её (patch_inputs.py)"
        sRow = 44 + childIdx: sCol = 3                                   ' п. 4 C:I
    End If
    If inp = 1 Then
        dRow = 5 + childIdx: dCol = 3                                    ' раздел А C:F, файл — G
    Else
        If S(cNet.Range("U4").Value) <> "ВВОДЫ 2 И 3" Then Err.Raise vbObjectError + 11, , "в книге нет таблицы «Вводы 2 и 3» — обновите её (patch_inputs.py)"
        dRow = 6 + (childIdx - 1) * 2 + (inp - 2): dCol = 23            ' W:Z, файл — AA
    End If
    Set up = OpenBook(upPath, opened)
    If up Is child Then Err.Raise vbObjectError + 4, , "выбрана эта же книга"
    Set uSub = SubSheet(up, r0)
    If uSub Is Nothing Or Not HasSheet(up, "Исходные данные") Then Err.Raise vbObjectError + 5, , "«" & up.Name & "» — не книга однолинейки"
    ids = up.Sheets("Исходные данные").Range("D3:D1000").Value
    For i = 1 To UBound(ids, 1)
        If S(ids(i, 1)) = feeder Then found = True: Exit For
    Next i
    If Not found Then Err.Raise vbObjectError + 6, , "в «" & up.Name & "» нет линии «" & feeder & "» («Исходные данные» D)"
    r = 0: free = 0
    For i = r0 To r0 + 29
        If S(uSub.Cells(i, 2).Value) = feeder Then r = i: Exit For
        If free = 0 And S(uSub.Cells(i, 2).Value) = "" Then free = i
    Next i
    If r = 0 Then r = free
    If r = 0 Then Err.Raise vbObjectError + 7, , "в «" & up.Name & "» нет свободной строки нижестоящих (30 шт.)"
    refC = ExtRef(child, cNet)
    refU = ExtRef(up, uSub)
    uSub.Cells(r, 2).Value = feeder
    uSub.Cells(r, 3).Value = RelPath(up.Path, child.FullName)
    For j = 0 To 6          ' нагрузка нижестоящего (весь щит или своя секция) -> вышестоящий D:J
        uSub.Cells(r, 4 + j).Formula = "=" & refC & cNet.Cells(sRow, sCol + j).Address
    Next j
    For j = 0 To 3          ' вышестоящий K:N -> питание этого ввода
        cNet.Cells(dRow, dCol + j).Formula = "=" & refU & uSub.Cells(r, 11 + j).Address
    Next j
    cNet.Cells(dRow, dCol + 4).Value = RelPath(child.Path, up.FullName) & " / " & feeder
    ' для схемы: номер автомата фидера, автомат, кабель, длина, щит-источник (BK:BO вышестоящего)
    If S(uSub.Range("BK92").Value) = "ДЛЯ СХЕМЫ НИЖЕСТОЯЩЕГО" Then
        For j = 0 To 4
            cNet.Cells(dRow, IIf(inp = 1, 8, 38) + j).Formula = "=" & refU & uSub.Cells(r, 63 + j).Address
        Next j
    End If
    Application.Calculate
    up.Save                 ' вышестоящую не закрываем: закрытие сразу после связи вешает скрытый Excel
    child.Activate          ' активной остаётся книга щита (следующая связь — снова от неё)
    SX_LastMsg = "Связано: щит «" & board & "»" & IIf(nIn > 1, ", ввод " & inp, "") & " ← «" & up.Name & "», линия " & feeder & "." & vbLf & _
        "Вышестоящая: раздел Б, строка " & r & " — сохранена и открыта. Эту книгу сохраните (Ctrl+S)."
    SX_LinkUpstreamTo = SX_LastMsg
    Exit Function
fail:
    SX_LastMsg = "Связь не создана: " & Err.Description
    SX_LinkUpstreamTo = SX_LastMsg
    If opened Then
        On Error Resume Next
        up.Close False
    End If
End Function

' кнопка «Связать с вышестоящим»: выбрать щит этой книги, файл вышестоящего, линию-фидер
Public Sub SX_LinkUpstream()
    Dim child As Workbook, cNet As Worksheet, f As Variant, up As Workbook, opened As Boolean
    Dim i As Long, k As Long, list As String, board As String, ans As String, feeder As String, a As Variant
    Set child = Book(): If child Is Nothing Then Exit Sub
    If child.Path = "" Then Say "Сначала сохраните эту книгу в папку объекта.", vbExclamation: Exit Sub
    Set cNet = NetSheet(child)
    If cNet Is Nothing Then Say "В этой книге нет листа «Сеть_щита».", vbExclamation: Exit Sub
    k = 1
    For i = 1 To 3
        If S(cNet.Cells(5 + i, 2).Value) <> "" Then list = list & vbLf & i & " — " & S(cNet.Cells(5 + i, 2).Value)
    Next i
    If S(cNet.Cells(7, 2).Value) <> "" Then
        ans = InputBox("Какой щит этой книги подключить?" & list, "Связать с вышестоящим", "1")
        If ans = "" Then Exit Sub
        k = Val(ans)
    End If
    board = S(cNet.Cells(5 + k, 2).Value)
    Dim inp As Long, nIn As Long
    inp = 1: nIn = CLng(NumV(cNet.Cells(65, 2 + k).Value))
    If nIn >= 2 Then
        ans = InputBox("У щита «" & board & "» вводов: " & nIn & ". Какой ввод подключить к этому фидеру?", "Связать с вышестоящим", "1")
        If ans = "" Then Exit Sub
        inp = Val(ans)
    End If
    On Error Resume Next
    ChDrive Left$(child.Path, 1)
    ChDir child.Path
    On Error GoTo 0
    f = Application.GetOpenFilename("Книги Excel (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , "Файл вышестоящего щита для «" & board & "»")
    If VarType(f) = vbBoolean Then Exit Sub
    ' подсказка: линия вышестоящего, в наименовании которой есть имя щита
    Set up = OpenBook(CStr(f), opened)
    If HasSheet(up, "Исходные данные") Then
        a = up.Sheets("Исходные данные").Range("D3:AJ1000").Value
        For i = 1 To UBound(a, 1)
            If S(a(i, 1)) <> "" And InStr(1, S(a(i, 33)), board, vbTextCompare) > 0 Then feeder = S(a(i, 1)): Exit For
        Next i
    End If
    If opened Then up.Close False
    feeder = Trim$(InputBox("Линия-фидер в «" & Dir(CStr(f)) & "», которая питает щит «" & board & "»" & vbLf & _
        "(номер как в «Исходных данных», столбец D):", "Связать с вышестоящим", feeder))
    If feeder = "" Then Exit Sub
    SX_LinkUpstreamTo CStr(f), feeder, k, inp
    Say SX_LastMsg, IIf(Left$(SX_LastMsg, 7) = "Связано", vbInformation, vbExclamation)
End Sub

' после ChangeLink: подписи «файл / фидер» и раздел Б C — на новый путь
Private Sub FixLinkTexts(wb As Workbook, oldFull As String, newFull As String)
    Dim net As Worksheet, sb As Worksheet, r0 As Long, i As Long, t As String, p As Variant, c As Variant
    Set net = NetSheet(wb): Set sb = SubSheet(wb, r0)
    If Not net Is Nothing Then
        For Each c In Array("G6", "G7", "G8", "AA6", "AA7", "AA8", "AA9", "AA10", "AA11")
            t = S(net.Range(c).Value)
            If InStr(1, t, ".xls", vbTextCompare) > 0 Then
                p = Split(t & " / ", " / ")
                If LCase$(AbsPath(wb.Path, CStr(p(0)))) = LCase$(oldFull) Then _
                    net.Range(c).Value = RelPath(wb.Path, newFull) & IIf(p(1) <> "", " / " & p(1), "")
            End If
        Next c
    End If
    If Not sb Is Nothing Then
        For i = r0 To r0 + 29
            t = S(sb.Cells(i, 3).Value)
            If t <> "" Then
                If LCase$(AbsPath(wb.Path, t)) = LCase$(oldFull) Then sb.Cells(i, 3).Value = RelPath(wb.Path, newFull)
            End If
        Next i
    End If
End Sub

' кнопка «Обновить связи»: обновить связи с другими книгами, показать, что изменилось
Public Sub SX_UpdateLinks()
    Dim wb As Workbook, links As Variant, i As Long, net As Worksheet, sb As Worksheet, r0 As Long
    Dim b1 As Variant, b2 As Variant, a1 As Variant, a2 As Variant, msg As String, bad As String
    Dim n As Long, r As Long, c As Long
    Set wb = Book(): If wb Is Nothing Then Exit Sub
    links = wb.LinkSources(xlExcelLinks)
    If IsEmpty(links) Then Say "Связей с другими книгами нет. Создайте: «Связать с вышестоящим».": Exit Sub
    Set net = NetSheet(wb): Set sb = SubSheet(wb, r0)
    If Not net Is Nothing Then a1 = net.Range("C6:F8").Value
    If Not sb Is Nothing Then b1 = sb.Range(sb.Cells(r0, 2), sb.Cells(r0 + 29, 10)).Value
    Dim cands As Collection, fixd As String, root As String
    root = ObjectRoot(wb)                    ' файл переехал — ищем по имени в папке объекта, затем уровнем выше
    For i = LBound(links) To UBound(links)
        If Dir(links(i)) = "" Then
            Set cands = New Collection
            FindBooks root, cands, Mid$(links(i), InStrRev(links(i), "\") + 1)
            If cands.Count = 0 Then FindBooks Left$(root, InStrRev(root, "\") - 1), cands, Mid$(links(i), InStrRev(links(i), "\") + 1)
            If cands.Count = 1 Then
                On Error Resume Next
                wb.ChangeLink links(i), cands(1), xlExcelLinks
                If Err.Number = 0 Then
                    FixLinkTexts wb, CStr(links(i)), CStr(cands(1))
                    fixd = fixd & vbLf & "  " & RelPath(wb.Path, CStr(links(i))) & " → " & RelPath(wb.Path, CStr(cands(1)))
                Else
                    bad = bad & vbLf & "  не перенаправлено: " & links(i)
                End If
                On Error GoTo 0
            Else
                bad = bad & vbLf & "  нет файла: " & links(i) & IIf(cands.Count > 1, " (найдено " & cands.Count & " с тем же именем — укажите вручную)", "")
            End If
        Else
            On Error Resume Next
            Err.Clear
            wb.UpdateLink links(i), xlExcelLinks
            If Err.Number <> 0 Then bad = bad & vbLf & "  не обновлено: " & links(i)
            On Error GoTo 0
        End If
    Next i
    Application.Calculate
    If Not net Is Nothing Then
        a2 = net.Range("C6:F8").Value
        For r = 1 To 3
            For c = 1 To 4
                If S(a1(r, c)) <> S(a2(r, c)) Then
                    n = n + 1
                    If n <= 15 Then msg = msg & vbLf & "  щит «" & S(net.Cells(5 + r, 2).Value) & "», " & _
                        S(net.Cells(5, 2 + c).Value) & ": " & S(a1(r, c)) & " → " & S(a2(r, c))
                End If
            Next c
        Next r
    End If
    If Not sb Is Nothing Then
        b2 = sb.Range(sb.Cells(r0, 2), sb.Cells(r0 + 29, 10)).Value
        For r = 1 To 30
            For c = 3 To 9
                If S(b1(r, c)) <> S(b2(r, c)) Then
                    n = n + 1
                    If n <= 15 Then msg = msg & vbLf & "  фидер " & S(b2(r, 1)) & ", " & _
                        S(sb.Cells(r0 - 1, 1 + c).Value) & ": " & S(b1(r, c)) & " → " & S(b2(r, c))
                End If
            Next c
        Next r
    End If
    SX_LastMsg = "Связей: " & (UBound(links) - LBound(links) + 1) & ". Изменилось значений: " & n & "." & msg & _
        IIf(n > 15, vbLf & "  … и ещё " & (n - 15), "") & IIf(fixd <> "", vbLf & "Перенаправлено (файл переехал):" & fixd, "") & _
        IIf(bad <> "", vbLf & "Не обновлено:" & bad, "")
    Say SX_LastMsg, IIf(bad <> "", vbExclamation, vbInformation)
End Sub

' все книги однолинеек папки: файл -> данные щита (для дерева и ПЗ)
Private Function CollectInfo(folder As String, full As Boolean) As Object
    Dim f As String, files As Collection, info As Object, k As Variant, w As Workbook, opened As Boolean
    Dim net As Worksheet, sb As Worksheet, r0 As Long, i As Long, rec As Object, md As Worksheet, a As Variant
    Dim kz As Variant, n As Long, leaf As Collection, src As Worksheet
    Set info = CreateObject("Scripting.Dictionary")
    Set files = New Collection
    gRoot = folder
    FindBooks folder, files
    For Each k In files
        opened = False
        Set w = Nothing
        On Error Resume Next
        Set w = OpenBook(CStr(k), opened)
        On Error GoTo 0
        If Not w Is Nothing Then
            Set net = NetSheet(w): Set sb = SubSheet(w, r0)
            If HasSheet(w, "Однолинейка") And Not net Is Nothing Then
                Set rec = CreateObject("Scripting.Dictionary")
                rec("щиты") = "": rec("от") = "": rec("ниже") = "": rec("подпись") = ""
                rec("имя") = RelPath(folder, CStr(k))
                For i = 1 To 3
                    If S(net.Cells(5 + i, 2).Value) <> "" Then
                        rec("щиты") = rec("щиты") & IIf(rec("щиты") = "", "", ", ") & S(net.Cells(5 + i, 2).Value)
                        AddFrom rec, S(net.Cells(5 + i, 7).Value), w.Path
                    End If
                Next i
                If S(net.Range("G6").Value) = "" And S(net.Range("C12").Value) <> "" Then _
                    AddFrom rec, "Ввод 1: от " & S(net.Range("C12").Value) & " (раздел 2)"
                If S(net.Range("U4").Value) = "ВВОДЫ 2 И 3" Then          ' вводы 2 и 3 — свой вышестоящий
                    For i = 6 To 11
                        f = S(net.Cells(i, 27).Value)
                        If f <> "" And InStr(1, f, ".xls", vbTextCompare) = 0 Then f = "Ввод " & net.Cells(i, 22).Value & ": " & f
                        AddFrom rec, f, w.Path
                    Next i
                End If
                rec("Руст") = net.Cells(45, 3).Value: rec("Iр") = net.Cells(45, 9).Value
                rec("dU") = net.Cells(33, 3).Value: rec("Iкз") = net.Cells(35, 3).Value
                rec("пр_апп") = S(net.Range("C38").Value)
                If Not sb Is Nothing Then
                    For i = r0 To r0 + 29
                        If S(sb.Cells(i, 3).Value) <> "" Then rec("ниже") = rec("ниже") & "|" & _
                            LCase$(AbsPath(w.Path, S(sb.Cells(i, 3).Value))) & "#" & S(sb.Cells(i, 2).Value)
                    Next i
                End If
                If full Then
                    rec("вводов") = net.Range("C65").Value: rec("АВР") = S(net.Range("C66").Value)
                    rec("кат") = S(net.Range("C68").Value)
                    rec("Zп") = net.Range("C34").Value: rec("Iкз3") = net.Range("C36").Value
                    rec("пр_откл") = S(net.Range("C39").Value)
                    rec("пр_терм") = S(net.Range("C40").Value): rec("пр_сел") = S(net.Range("C41").Value)
                    If HasSheet(w, "Режимы") Then
                        Set md = w.Sheets("Режимы")
                        a = md.Range("H14:H21").Value              ' нормальный режим, «Итого»
                        rec("Кс") = a(2, 1): rec("cos") = a(3, 1): rec("Pр") = a(5, 1)
                        rec("Qр") = a(6, 1): rec("Sр") = a(7, 1): rec("Iрн") = a(8, 1)
                        a = md.Range("H59:H62").Value              ' максимум аварийного
                        rec("Pр_ав") = a(1, 1): rec("Sр_ав") = a(3, 1): rec("Iр_ав") = a(4, 1)
                        rec("реж") = md.Range("B11:I26").Value
                        rec("ав") = md.Range("B58:I63").Value
                    End If
                    ' токи КЗ по линиям — раздел В
                    n = 0
                    Do While S(net.Cells(132 + n, 1).Value) <> "" And n < 998
                        n = n + 1
                    Loop
                    If n > 0 Then rec("kz") = net.Range(net.Cells(132, 1), net.Cells(131 + n, 13)).Value
                    ' линии потребителей (не фидеры щитов): категория для объекта
                    Set leaf = New Collection
                    Set src = w.Sheets("Исходные данные")
                    a = src.Range("A3:BV1000").Value
                    For i = 1 To UBound(a, 1)
                        If S(a(i, 4)) <> "" And S(a(i, 55)) <> "Щит" And S(a(i, 55)) <> "Нет" Then
                            leaf.Add Array(S(net.Cells(6, 2).Value), S(a(i, 4)), S(a(i, 36)), NumV(a(i, 38)), _
                                S(a(i, 68)), S(a(i, 69)), S(a(i, 74)))
                        End If
                    Next i
                    Set rec("leaf") = leaf
                    If HasSheet(w, "Расчёт_нагрузок") Then rec("РН") = w.Sheets("Расчёт_нагрузок").Range("A1:J61").Value
                End If
                Set info(LCase$(CStr(k))) = rec
            End If
            If opened Then w.Close False
        End If
    Next k
    Set CollectInfo = info
End Function

' дерево в лист ws с строки 4; возвращает следующую свободную строку
Private Function WriteTree(ws As Worksheet, info As Object) As Long
    Dim row As Long, k As Variant, p As Variant, kids As String
    ws.Range("A3:H3").Value = Array("Щит (дерево)", "Файл", "Питается от", "Руст, кВт", "Iр, А", "ΔU до шин, %", _
        "Iкз min на шинах, А", "Проверка")
    ws.Range("A3:H3").Font.Bold = True
    ws.Range("A3:H3").Interior.Color = RGB(221, 235, 247)
    row = 4
    For Each k In info.Keys
        kids = kids & info(k)("ниже") & "|"
    Next k
    For Each k In info.Keys          ' корни: вышестоящего нет (или нет в папке) и никто не указал книгу нижестоящей
        p = Split(info(k)("от") & " / ", " / ")
        If (info(k)("от") = "" Or Not info.Exists(LCase$(p(0)))) And InStr(kids, "|" & k & "#") = 0 Then
            TreeRow ws, info, CStr(k), 0, row, ""
        End If
    Next k
    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B:C").ColumnWidth = 30
    ws.Columns("D:G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 60
    If row > 4 Then ws.Range("A3:H" & row - 1).Borders.LineStyle = xlContinuous
    WriteTree = row
End Function

' ключи книг в порядке дерева (сверху вниз), затем остальные
Private Function OrderKeys(info As Object) As Collection
    Dim c As New Collection, seen As Object, k As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    If RootKey(info) <> "" Then OrderWalk info, RootKey(info), c, seen
    For Each k In info.Keys
        If Not seen.Exists(CStr(k)) Then OrderWalk info, CStr(k), c, seen
    Next k
    Set OrderKeys = c
End Function

Private Sub OrderWalk(info As Object, k As String, c As Collection, seen As Object)
    Dim kids As Variant, i As Long, ch As String
    If seen.Exists(k) Or Not info.Exists(k) Then Exit Sub
    seen(k) = 1: c.Add k
    kids = Split(info(k)("ниже"), "|")
    For i = 1 To UBound(kids)
        ch = Split(kids(i) & "#", "#")(0)
        OrderWalk info, ch, c, seen
    Next i
End Sub

' корневая книга (ГРЩ / ВРУ): нет вышестоящего в папке и никто не питает её
Private Function RootKey(info As Object) As String
    Dim k As Variant, p As Variant, kids As String
    For Each k In info.Keys
        kids = kids & info(k)("ниже") & "|"
    Next k
    For Each k In info.Keys
        p = Split(info(k)("от") & " / ", " / ")
        If (info(k)("от") = "" Or Not info.Exists(LCase$(p(0)))) And InStr(kids, "|" & k & "#") = 0 Then RootKey = CStr(k): Exit Function
    Next k
End Function

' кнопка «Дерево щитов»: все книги однолинеек в папке этой книги -> лист «Дерево_щитов»
Public Sub SX_Tree()
    Dim wb As Workbook, info As Object, ws As Worksheet, relock As Boolean
    Set wb = Book(): If wb Is Nothing Then Exit Sub
    If wb.Path = "" Then Say "Сначала сохраните книгу в папку объекта.", vbExclamation: Exit Sub
    Fast True
    On Error GoTo fin
    Set info = CollectInfo(ObjectRoot(wb), False)
    wb.Activate
    If wb.ProtectStructure Then wb.Unprotect PWD: relock = True
    Set ws = SheetReset(wb, "Дерево_щитов")
    ws.Range("A1").Value = "ДЕРЕВО ЩИТОВ — папка объекта " & gRoot & ". Обновлено " & Format(Now, "dd.mm.yyyy hh:nn") & _
        ". Связи ставит кнопка «Связать с вышестоящим»."
    ws.Range("A1").Font.Bold = True
    WriteTree ws, info
fin:
    If relock Then wb.Protect PWD, Structure:=True
    Fast False
    If Err.Number <> 0 Then
        Say "Дерево щитов: " & Err.Description, vbExclamation
    Else
        ws.Activate
        Say "Дерево щитов: книг однолинеек — " & info.Count & ", лист «Дерево_щитов»."
    End If
End Sub

' ---------------------------------------------------------------- ПЗ по объекту
Private Function F2(v As Variant, Optional d As Long = 2) As String
    If IsNumeric(v) And S(v) <> "" Then F2 = FormatNumber(CDbl(v), d, vbTrue, vbFalse, vbFalse) Else F2 = S(v)
End Function

Private Function CatRank(c As String) As Long
    Select Case c
        Case "I особая": CatRank = 4
        Case "I": CatRank = 3
        Case "II": CatRank = 2
        Case "III": CatRank = 1
    End Select
End Function

' лист ПЗ: заголовок, шапка, строки (массив массивов)
Private Function PzSheet(pz As Workbook, nm As String, title As String, hdr As Variant, rows As Collection) As Worksheet
    Dim ws As Worksheet, i As Long, j As Long, r As Variant
    Set ws = pz.Sheets.Add(After:=pz.Sheets(pz.Sheets.Count))
    ws.Name = nm
    ws.Range("A1").Value = title
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    ws.Range("A3").Resize(1, UBound(hdr) + 1).Value = hdr
    With ws.Range("A3").Resize(1, UBound(hdr) + 1)
        .Font.Bold = True: .WrapText = True: .Interior.Color = RGB(221, 235, 247): .VerticalAlignment = xlCenter
    End With
    i = 4
    For Each r In rows
        For j = 0 To UBound(r)
            If VarType(r(j)) = vbString And IsNumeric(r(j)) And r(j) <> "" Then
                ws.Cells(i, j + 1).Value = CDbl(r(j))       ' «0,594» -> число (иначе «1,000» станет тысячей)
            Else
                ws.Cells(i, j + 1).Value = r(j)
            End If
        Next j
        i = i + 1
    Next r
    ws.Range("A3").Resize(i - 3, UBound(hdr) + 1).Borders.LineStyle = xlContinuous
    ws.Columns("A:Z").ColumnWidth = 13
    ws.Columns("A").ColumnWidth = 18
    Set PzSheet = ws
End Function

Private Sub WordPara(doc As Object, txt As String, Optional bold As Boolean = False, Optional size As Single = 11)
    Dim rg As Object
    Set rg = doc.Content
    rg.InsertParagraphAfter
    Set rg = doc.Paragraphs(doc.Paragraphs.Count).Range
    rg.Text = txt
    rg.Font.Bold = bold
    rg.Font.Size = size
    rg.Font.Name = "Times New Roman"
End Sub

Private Sub WordTable(doc As Object, hdr As Variant, rows As Collection, Optional caption As String = "")
    Dim rg As Object, t As Object, i As Long, j As Long, r As Variant
    If caption <> "" Then WordPara doc, caption
    doc.Content.InsertParagraphAfter
    Set rg = doc.Paragraphs(doc.Paragraphs.Count).Range
    Set t = doc.Tables.Add(rg, rows.Count + 1, UBound(hdr) + 1)
    t.Borders.Enable = True
    t.Range.Font.Size = 9
    t.Range.Font.Name = "Times New Roman"
    For j = 0 To UBound(hdr)
        t.Cell(1, j + 1).Range.Text = hdr(j)
        t.Cell(1, j + 1).Range.Font.Bold = True
    Next j
    i = 2
    For Each r In rows
        For j = 0 To UBound(r)
            t.Cell(i, j + 1).Range.Text = CStr(r(j))
        Next j
        i = i + 1
    Next r
End Sub

' кнопка «ПЗ по объекту»: из всех книг папки — ПЗ_<объект>.xlsx (таблицы) и ПЗ_<объект>.docx (расчётная часть)
Public Sub SX_PZ()
    Dim wb As Workbook, folder As String, obj As String, info As Object, k As Variant, rec As Object, root As String
    Dim pz As Workbook, ws As Worksheet, rows As Collection, hdr As Variant, r As Variant, i As Long, j As Long
    Dim catP As Object, catN As Object, firstCat As Collection, worst As String, issues As Collection, v As Variant
    Dim wd As Object, doc As Object, newWord As Boolean, xlsPath As String, docPath As String, a As Variant
    Set wb = Book(): If wb Is Nothing Then Exit Sub
    If wb.Path = "" Then Say "Сначала сохраните книгу в папку объекта.", vbExclamation: Exit Sub
    folder = ObjectRoot(wb)
    obj = Mid$(folder, InStrRev(folder, "\") + 1)
    Fast True
    On Error GoTo fin
    Set info = CollectInfo(folder, True)
    If info.Count = 0 Then Err.Raise vbObjectError + 20, , "в папке нет книг однолинеек"
    root = RootKey(info)
    Set issues = New Collection
    Set catP = CreateObject("Scripting.Dictionary"): Set catN = CreateObject("Scripting.Dictionary")
    Set firstCat = New Collection
    For Each v In Array("I особая", "I", "II", "III")
        catP(v) = 0#: catN(v) = 0
    Next v
    For Each k In info.Keys
        Set rec = info(k)
        If rec.Exists("leaf") Then
            For Each r In rec("leaf")
                If catP.Exists(r(4)) Then catP(r(4)) = catP(r(4)) + r(3): catN(r(4)) = catN(r(4)) + 1
                If CatRank(CStr(r(4))) > CatRank(worst) Then worst = CStr(r(4))
                If CatRank(CStr(r(4))) >= 3 Then firstCat.Add Array(r(0), r(1), r(2), F2(r(3)), r(4), r(6))
                If r(6) <> "" And r(6) <> "ОК" Then issues.Add "щит " & r(0) & ", линия " & r(1) & ": категория " & r(4) & " — " & r(6)
            Next r
        End If
        For Each v In Array(rec("пр_апп"), rec("пр_откл"), rec("пр_терм"))
            If v <> "" And LCase$(Left$(v, 2)) <> "ok" Then issues.Add "щит " & rec("щиты") & ", питающая линия: " & v
        Next v
        If IsNumeric(rec("dU")) And S(rec("dU")) <> "" Then If CDbl(rec("dU")) > 4 Then issues.Add "щит " & rec("щиты") & ": ΔU до шин " & F2(rec("dU")) & " % > 4 %"
        If rec.Exists("kz") Then
            a = rec("kz")
            For i = 1 To UBound(a, 1)
                If Left$(S(a(i, 13)), 4) = "мало" Then issues.Add "щит " & rec("щиты") & ", линия " & S(a(i, 2)) & ": " & S(a(i, 13))
                If Left$(S(a(i, 13)), 12) = "не проверено" Then issues.Add "щит " & rec("щиты") & ", линия " & S(a(i, 2)) & ": КЗ не проверен (нет питания)"
            Next i
        End If
    Next k
    If worst = "" Then worst = "III"

    ' ---------- Excel
    Set pz = Application.Workbooks.Add
    Set rec = info(root)
    Set rows = New Collection
    rows.Add Array("Объект (папка)", obj)
    rows.Add Array("Щитов (книг)", info.Count)
    rows.Add Array("Главный щит", rec("щиты") & " (" & rec("имя") & ")")
    rows.Add Array("Питание главного щита", rec("подпись") & IIf(rec("от") <> "", " " & Vid(rec("от")), ""))
    rows.Add Array("Вводов / АВР", S(rec("вводов")) & " / " & rec("АВР"))
    rows.Add Array("Руст объекта, кВт", F2(rec("Руст")))
    rows.Add Array("Нормальный режим: Рр, кВт / Sр, кВА / Iр, А", F2(rec("Pр")) & " / " & F2(rec("Sр")) & " / " & F2(rec("Iрн"), 1))
    rows.Add Array("Аварийный (макс.): Рр, кВт / Sр, кВА / Iр, А", F2(rec("Pр_ав")) & " / " & F2(rec("Sр_ав")) & " / " & F2(rec("Iр_ав"), 1))
    rows.Add Array("Кс / cos φ (нормальный)", F2(rec("Кс"), 3) & " / " & F2(rec("cos"), 2))
    rows.Add Array("Высшая категория потребителей", worst)
    rows.Add Array("Обеспечено вводами главного щита", rec("кат"))
    rows.Add Array("Замечаний", issues.Count)
    Set ws = PzSheet(pz, "Сводка", "ПЗ ПО ОБЪЕКТУ «" & obj & "» — сводка. Создано " & Format(Now, "dd.mm.yyyy hh:nn"), Array("Показатель", "Значение"), rows)
    ws.Columns("A").ColumnWidth = 48: ws.Columns("B").ColumnWidth = 50
    Application.DisplayAlerts = False
    Do While pz.Sheets.Count > 1
        pz.Sheets(1).Delete
    Loop
    Application.DisplayAlerts = True

    hdr = Array("Щит", "Файл", "Питается от", "Вводов / АВР", "Руст, кВт", "Кс", "cos φ", "Рр, кВт", "Sр, кВА", "Iр, А", _
                "Iр авар., А", "Категория (вводы)", "ΔU до шин, %", "Iкз min на шинах, А")
    Set rows = New Collection
    For Each k In OrderKeys(info)
        Set rec = info(k)
        rows.Add Array(rec("щиты"), rec("имя"), Vid(rec("от")) & IIf(rec("от") <> "" And rec("подпись") <> "", "; ", "") & rec("подпись"), _
            S(rec("вводов")) & " / " & rec("АВР"), F2(rec("Руст")), F2(rec("Кс"), 3), F2(rec("cos")), F2(rec("Pр")), _
            F2(rec("Sр")), F2(rec("Iрн"), 1), F2(rec("Iр_ав"), 1), rec("кат"), F2(rec("dU")), F2(rec("Iкз"), 0))
    Next k
    Set ws = PzSheet(pz, "Нагрузки щитов", "Расчётные нагрузки щитов (нормальный режим; аварийный — максимум при потере ввода)", hdr, rows)
    Dim tbLoads As Collection: Set tbLoads = rows

    Set rec = info(root)
    If rec.Exists("РН") Then
        Set ws = pz.Sheets.Add(After:=pz.Sheets(pz.Sheets.Count))
        ws.Name = "Расчёт нагрузок ГЩ"
        ws.Range("A1:J61").Value = rec("РН")
        ws.Range("A1").Font.Bold = True
        ws.Columns("A").ColumnWidth = 6: ws.Columns("B").ColumnWidth = 40: ws.Columns("C:J").ColumnWidth = 12
    End If
    If rec.Exists("реж") Then
        Set ws = pz.Sheets.Add(After:=pz.Sheets(pz.Sheets.Count))
        ws.Name = "Режимы ГЩ"
        ws.Range("A1").Value = "Режимы питания главного щита «" & rec("щиты") & "» (нормальный и аварийный)"
        ws.Range("A1").Font.Bold = True
        ws.Range("A3:H18").Value = rec("реж")
        ws.Range("A20:H25").Value = rec("ав")
        ws.Columns("A").ColumnWidth = 40: ws.Columns("B:H").ColumnWidth = 12
    End If

    hdr = Array("Категория", "Линий", "Руст, кВт", "Доля, %")
    Set rows = New Collection
    Dim tot As Double: tot = 0
    For Each v In catP.Keys
        tot = tot + catP(v)
    Next v
    For Each v In Array("I особая", "I", "II", "III")
        rows.Add Array(v, catN(v), F2(catP(v)), IIf(tot > 0, F2(catP(v) / tot * 100, 1), ""))
    Next v
    rows.Add Array("Итого", catN("I особая") + catN("I") + catN("II") + catN("III"), F2(tot), "100")
    Set ws = PzSheet(pz, "Категории", "Категории надёжности по объекту (линии потребителей всех щитов, без фидеров щитов)", hdr, rows)
    Dim tbCat As Collection: Set tbCat = rows
    i = 5 + rows.Count + 2
    ws.Cells(i, 1).Value = "Потребители I категории и выше": ws.Cells(i, 1).Font.Bold = True
    ws.Cells(i + 1, 1).Resize(1, 6).Value = Array("Щит", "Линия", "Наименование", "Руст, кВт", "Категория", "Проверка")
    ws.Cells(i + 1, 1).Resize(1, 6).Font.Bold = True
    j = i + 2
    For Each r In firstCat
        ws.Cells(j, 1).Resize(1, 6).Value = r
        j = j + 1
    Next r

    hdr = Array("Щит", "ΔU до шин, %", "Zп на шинах, Ом", "Iкз min, А", "Iкз max, кА", "Аппарат и кабель", "Отключение КЗ", "Термостойкость", "Селективность")
    Set rows = New Collection
    For Each k In OrderKeys(info)
        Set rec = info(k)
        rows.Add Array(rec("щиты"), F2(rec("dU")), F2(rec("Zп"), 4), F2(rec("Iкз"), 0), F2(rec("Iкз3"), 2), _
            rec("пр_апп"), rec("пр_откл"), rec("пр_терм"), rec("пр_сел"))
    Next k
    Set ws = PzSheet(pz, "Сеть и КЗ", "Потери напряжения и токи КЗ на шинах щитов", hdr, rows)
    Dim tbNet As Collection: Set tbNet = rows

    hdr = Array("Щит", "№", "Линия", "L, м", "In, А", "Наименование", "Кабель", "Zп в конце, Ом", "Iкз min, А", "Iкз min / In", "Iкз max, кА", "Отключение ≤ 0,4 с")
    Set rows = New Collection
    For Each k In OrderKeys(info)
        Set rec = info(k)
        If rec.Exists("kz") Then
            a = rec("kz")
            For i = 1 To UBound(a, 1)
                rows.Add Array(rec("щиты"), a(i, 1), a(i, 2), a(i, 4), a(i, 5), a(i, 7), a(i, 8), a(i, 9), a(i, 10), a(i, 11), a(i, 12), a(i, 13))
            Next i
        End If
    Next k
    Set ws = PzSheet(pz, "КЗ по линиям", "Токи короткого замыкания в конце линий (ГОСТ 28249, ПУЭ 1.7.79)", hdr, rows)
    ws.Columns("F").ColumnWidth = 36: ws.Columns("G").ColumnWidth = 24: ws.Columns("L").ColumnWidth = 30
    Dim nKZ As Long: nKZ = rows.Count

    Set ws = pz.Sheets.Add(After:=pz.Sheets(pz.Sheets.Count))
    ws.Name = "Дерево щитов"
    ws.Range("A1").Value = "Схема питания: щиты объекта «" & obj & "»"
    ws.Range("A1").Font.Bold = True
    WriteTree ws, info

    hdr = Array("№", "Замечание")
    Set rows = New Collection
    i = 1
    For Each v In issues
        rows.Add Array(i, v): i = i + 1
    Next v
    If issues.Count = 0 Then rows.Add Array("—", "замечаний нет: категории обеспечены, КЗ и ΔU в норме")
    Set ws = PzSheet(pz, "Замечания", "Замечания по расчётам", hdr, rows)
    ws.Columns("B").ColumnWidth = 100
    pz.Sheets(1).Activate
    xlsPath = folder & "\ПЗ_" & obj & ".xlsx"
    Application.DisplayAlerts = False
    pz.SaveAs xlsPath, 51
    Application.DisplayAlerts = True

    ' ---------- Word
    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    On Error GoTo fin
    If wd Is Nothing Then Set wd = CreateObject("Word.Application"): newWord = True
    Set doc = wd.Documents.Add
    Set rec = info(root)
    doc.Content.Text = "РАСЧЁТНАЯ ЧАСТЬ. ЭЛЕКТРОСНАБЖЕНИЕ"
    doc.Paragraphs(1).Range.Font.Bold = True
    doc.Paragraphs(1).Range.Font.Size = 14
    doc.Paragraphs(1).Range.Font.Name = "Times New Roman"
    WordPara doc, "Объект: " & obj & ". Расчёт выполнен по однолинейным схемам щитов (" & info.Count & " шт.), " & Format(Now, "dd.mm.yyyy") & "."
    WordPara doc, "1. Источник питания и схема", True, 12
    WordPara doc, "Главный щит " & rec("щиты") & " питается: " & rec("подпись") & IIf(rec("от") <> "", " " & Vid(rec("от")), "") & _
        ". Число вводов — " & S(rec("вводов")) & ", АВР — " & rec("АВР") & ". Обеспечиваемая категория надёжности: " & rec("кат") & _
        ". Высшая категория потребителей объекта: " & worst & "."
    WordPara doc, "2. Расчётные нагрузки", True, 12
    WordPara doc, "Расчётная нагрузка определена по СП 256.1325800: Рр = Руст.быт · Кс + Ко · Σ(Рсил · Кс) + Σ Рпост " & _
        "(Ко = 0,9 при наличии бытовой нагрузки). Установленная мощность объекта Руст = " & F2(rec("Руст")) & " кВт. " & _
        "Нормальный режим: Рр = " & F2(rec("Pр")) & " кВт, Sр = " & F2(rec("Sр")) & " кВА, Iр = " & F2(rec("Iрн"), 1) & " А, cos φ = " & F2(rec("cos")) & ". " & _
        "Аварийный режим (потеря одного ввода, вся нагрузка на оставшемся): Рр = " & F2(rec("Pр_ав")) & " кВт, Iр = " & F2(rec("Iр_ав"), 1) & " А."
    Dim t2 As New Collection
    For Each r In tbLoads
        t2.Add Array(r(0), r(4), r(5), r(7), r(8), r(9), r(10), r(11))
    Next r
    WordTable doc, Array("Щит", "Руст, кВт", "Кс", "Рр, кВт", "Sр, кВА", "Iр, А", "Iр авар., А", "Категория"), t2, "Таблица 1 — Расчётные нагрузки щитов"
    WordPara doc, "3. Надёжность электроснабжения", True, 12
    WordTable doc, Array("Категория", "Линий", "Руст, кВт", "Доля, %"), tbCat, "Таблица 2 — Нагрузки по категориям надёжности"
    WordPara doc, "Потребители I категории питаются от щитов с двумя вводами и АВР; проверка обеспеченности — в таблице «Категории» (Excel)."
    WordPara doc, "4. Потери напряжения и токи короткого замыкания", True, 12
    WordPara doc, "Потери напряжения — по ГОСТ Р 50571.5.52 (прил. G); допустимые от ВРУ до потребителя: освещение 3 %, прочие 4 % (СП 256). " & _
        "Токи КЗ — по ГОСТ 28249 (Iкз min: с = 0,95, нагретые жилы, переходные сопротивления; Iкз max: с = 1,05). " & _
        "Время отключения групповых линий ≤ 0,4 с (ПУЭ 1.7.79); распределительных — ≤ 5 с. Токи КЗ по " & nKZ & " линиям — лист «КЗ по линиям» (Excel)."
    Dim t3 As New Collection
    For Each r In tbNet
        t3.Add Array(r(0), r(1), r(2), r(3), r(4), r(6))
    Next r
    WordTable doc, Array("Щит", "ΔU до шин, %", "Zп, Ом", "Iкз min, А", "Iкз max, кА", "Отключение КЗ"), t3, "Таблица 3 — ΔU и токи КЗ на шинах щитов"
    WordPara doc, "5. Выводы", True, 12
    If issues.Count = 0 Then
        WordPara doc, "Категории надёжности обеспечены. Потери напряжения и время отключения КЗ — в пределах норм. Замечаний нет."
    Else
        WordPara doc, "Есть замечания (" & issues.Count & "), см. лист «Замечания»:"
        i = 0
        For Each v In issues
            i = i + 1
            If i > 20 Then WordPara doc, "… и ещё " & (issues.Count - 20): Exit For
            WordPara doc, "— " & v
        Next v
    End If
    docPath = folder & "\ПЗ_" & obj & ".docx"
    doc.SaveAs2 docPath, 16
    doc.Close False
    If newWord Then wd.Quit
    Set wd = Nothing
fin:
    Fast False
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then
        SX_LastMsg = "ПЗ: " & Err.Description
        Say SX_LastMsg, vbExclamation
    Else
        SX_LastMsg = "ПЗ по объекту: «" & Dir(xlsPath) & "» и «" & Dir(docPath) & "» в папке объекта. Щитов " & info.Count & _
            ", замечаний " & issues.Count & "."
        Say SX_LastMsg
    End If
End Sub


' «файл.xlsx / фидер» — вышестоящая книга; другой текст (ТП, вручную) — только подпись
Private Sub AddFrom(rec As Object, ByVal t As String, Optional base As String = "")
    Dim key As String, p As Variant
    If t = "" Then Exit Sub
    key = IIf(InStr(1, t, ".xls", vbTextCompare) > 0, "от", "подпись")
    If key = "от" And base <> "" Then                       ' путь относительно книги -> полный, строчными
        p = Split(t & " / ", " / ")
        t = AbsPath(base, CStr(p(0))) & IIf(p(1) <> "", " / " & p(1), "")      ' регистр — для вида; сравнение — LCase
    End If
    If InStr(rec(key), t) = 0 Then rec(key) = rec(key) & IIf(rec(key) = "", "", "; ") & t
End Sub

Private Sub TreeRow(ws As Worksheet, info As Object, k As String, lvl As Long, ByRef row As Long, ByVal path As String, _
                    Optional parent As String = "")
    Dim rec As Object, p As Variant, ch As String, chk As String, up As String, kids As Variant, i As Long
    If InStr(path, "|" & k & "|") > 0 Then Exit Sub          ' петля в связях
    path = path & "|" & k & "|"
    Set rec = info(k)
    up = rec("от")
    Dim ref As Variant
    If up <> "" Then
        For Each ref In Split(up, "; ")
            p = Split(ref & " / ", " / ")
            If Not info.Exists(LCase$(p(0))) Then
                chk = chk & "вышестоящего «" & RelPath(gRoot, CStr(p(0))) & "» нет в папке объекта; "
            ElseIf InStr(info(LCase$(p(0)))("ниже"), "|" & k & "#" & p(1)) = 0 Then
                chk = chk & "у «" & RelPath(gRoot, CStr(p(0))) & "» нет фидера " & p(1) & " в разделе Б; "
            End If
        Next ref
    End If
    If up = "" And parent <> "" Then chk = chk & "в книге не указан вышестоящий («Связать с вышестоящим»); "
    If IsNumeric(rec("dU")) And S(rec("dU")) <> "" Then
        If CDbl(rec("dU")) > 4 Then chk = chk & "ΔU до шин > 4 %; "
    End If
    If S(rec("Iкз")) = "" Then chk = chk & "Iкз не посчитан (нет данных питания); "
    ws.Cells(row, 1).Value = String(lvl * 4, " ") & IIf(lvl > 0, "└ ", "") & rec("щиты")
    ws.Cells(row, 2).Value = rec("имя")
    ws.Cells(row, 3).Value = Vid(up) & IIf(up <> "" And rec("подпись") <> "", "; ", "") & rec("подпись")
    If rec("пр_апп") <> "" And LCase$(Left$(rec("пр_апп"), 2)) <> "ok" Then chk = chk & rec("пр_апп") & "; "
    ws.Cells(row, 4).Value = rec("Руст")
    If IsNumeric(rec("Iр")) And S(rec("Iр")) <> "" Then ws.Cells(row, 5).Value = Round(CDbl(rec("Iр")), 1)
    If IsNumeric(rec("dU")) And S(rec("dU")) <> "" Then ws.Cells(row, 6).Value = Round(CDbl(rec("dU")), 2)
    If IsNumeric(rec("Iкз")) And S(rec("Iкз")) <> "" Then ws.Cells(row, 7).Value = Round(CDbl(rec("Iкз")), 0)
    ws.Cells(row, 8).Value = IIf(chk = "", "ok", chk)
    If chk <> "" Then ws.Cells(row, 8).Interior.Color = RGB(255, 242, 204)
    row = row + 1
    kids = Split(rec("ниже"), "|")
    Dim done As String
    For i = 1 To UBound(kids)
        ch = Split(kids(i), "#")(0)
        If InStr(done, "|" & ch & "|") > 0 Then GoTo nxt      ' щит с двумя фидерами от этого же — один раз
        done = done & "|" & ch & "|"
        If info.Exists(ch) Then
            TreeRow ws, info, ch, lvl + 1, row, path, k
        Else
            ws.Cells(row, 1).Value = String((lvl + 1) * 4, " ") & "└ ?"
            ws.Cells(row, 2).Value = RelPath(gRoot, ch)
            ws.Cells(row, 3).Value = rec("имя") & " / " & Split(kids(i) & "#", "#")(1)
            ws.Cells(row, 8).Value = "файла нет в папке объекта"
            ws.Cells(row, 8).Interior.Color = RGB(255, 199, 206)
            row = row + 1
        End If
nxt:
    Next i
End Sub

' ---------------------------------------------------------------- справка
Public Sub SX_Help()
    MsgBox "Вкладка «Однолинейка» (таблица v3.12):" & vbLf & vbLf & _
        "• Обновить спецификацию — «Спецификация_авто» переносится в «Спецификация», изменения против принятой ревизии подсвечиваются, удалённые — в листе «Изменения»." & vbLf & _
        "• Принять изменения — когда снабжение получило: подсветка снимается, ревизия +1." & vbLf & _
        "• Бирки — создаёт лист «Бирки» по «Однолинейке» (значения). Изменили схему — нажмите снова." & vbLf & _
        "• Наклейки — лента под аппаратами, клетка = 1 модуль DIN. Проверьте пробной печатью." & vbLf & _
        "• Связать с вышестоящим — один щит = одна книга: выбрать файл вышестоящего и линию-фидер; связи ставятся в обе книги." & vbLf & _
        "• Обновить связи — подтянуть данные из связанных книг и показать, что изменилось. Дерево щитов — схема и проверки по папке объекта." & vbLf & _
        "• ПЗ по объекту — из всех книг папки: ПЗ_<объект>.xlsx (таблицы) и ПЗ_<объект>.docx (расчётная часть)." & vbLf & _
        "• Только ввод / Ввод + расчёт / Все столбцы — какие столбцы видны на «Однолинейке» и «Исходных данных» (работает и на защищённом листе)." & vbLf & vbLf & _
        "Подробно — 4_Инструкции/Надстройка_на_ленте.md", vbInformation, "Однолинейка"
End Sub
