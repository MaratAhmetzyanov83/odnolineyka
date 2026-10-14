Attribute VB_Name = "SX_Catalog"
' Общий каталог: 1_Таблицы_Excel\Каталог_общий.xlsx -> лист «Каталог» активной книги.
' Ключ позиции — Категория (A) + Производитель (B) + Модель (C) + Наименование (G).
' Пишутся только данные изделия A–M и блок Q. Запас O, формулы N, P, U–AH и списки R, S в книге не трогаются.
' Формулы листа протянуты до строки 500: новые позиции — в первые пустые строки, строки не вставляются.
Option Explicit

Private Const CAT As String = "Каталог"
Private Const CAT_FILE As String = "Каталог_общий.xlsx"
Private Const HDR As Long = 3
Private Const R0 As Long = 4
Private Const R1 As Long = 500
Private Const NCOL As Long = 17                 ' A..Q
Private Const PWD As String = "sx"
Private Const REG_APP As String = "Odnolineyka"

' ---------------------------------------------------------------- кнопки
Public Sub SX_CatUpdate()
    Cat_Run ""
End Sub

' Книга -> общий каталог: новые позиции книги дописываются, разные по данным — по отдельному вопросу.
Public Sub SX_CatPush()
    Cat_Push "", False
End Sub

' Для проверки программой: путь к общему каталогу задан, окон нет (SX_SetQuiet). Возвращает итоговое сообщение.
Public Function SX_CatUpdateFrom(path As Variant) As String
    Cat_Run CStr(path)
    SX_CatUpdateFrom = SX_LastMsg
End Function

' Для проверки программой: withChanged — записать в общий и разные по данным позиции.
Public Function SX_CatPushTo(path As Variant, Optional withChanged As Boolean) As String
    Cat_Push CStr(path), withChanged
    SX_CatPushTo = SX_LastMsg
End Function

Private Sub CSay(txt As String, Optional st As VbMsgBoxStyle = vbInformation)
    SX_LastMsg = txt
    If SX_Quiet Then
        Debug.Print txt
    Else
        MsgBox txt, st, "Каталог"
    End If
End Sub

' столбцы с данными изделия: A..M и Q
Private Function DataCol(c As Long) As Boolean
    DataCol = (c <= 13 Or c = 17)
End Function

Private Function NormKey(a As Variant, b As Variant, c As Variant, g As Variant) As String
    NormKey = LCase$(Trim$(CStr(a))) & "|" & LCase$(Trim$(CStr(b))) & "|" & _
              LCase$(Trim$(CStr(c))) & "|" & LCase$(Trim$(CStr(g)))
End Function

' Текст сравнивается как текст (артикул «001234» <> «1234»), числа — как числа.
Private Function Same(x As Variant, y As Variant) As Boolean
    If IsError(x) Or IsError(y) Then Same = False: Exit Function
    If VarType(x) = vbString Or VarType(y) = vbString Then
        Same = (Trim$(CStr(x)) = Trim$(CStr(y))): Exit Function
    End If
    If IsNumeric(x) And IsNumeric(y) And Not IsEmpty(x) And Not IsEmpty(y) Then
        If Trim$(CStr(x)) <> "" And Trim$(CStr(y)) <> "" Then Same = (CDbl(x) = CDbl(y)): Exit Function
    End If
    Same = (Trim$(CStr(x)) = Trim$(CStr(y)))
End Function

' Записать значение; текст, похожий на число («001234»), остаётся текстом.
Private Sub PutVal(cel As Range, v As Variant)
    If VarType(v) = vbString Then
        If IsNumeric(v) Then cel.Value = "'" & v: Exit Sub
    End If
    cel.Value = v
End Sub

' ---------------------------------------------------------------- путь к общему каталогу
Private Function CatPath(forced As String) As String
    Dim p As String, f As Variant
    If forced <> "" Then CatPath = forced: Exit Function
    p = GetSetting(REG_APP, "Catalog", "Path", "")
    If p <> "" Then
        If Dir(p) <> "" Then CatPath = p: Exit Function
    End If
    p = ThisWorkbook.Path & "\..\1_Таблицы_Excel\" & CAT_FILE
    If Dir(p) <> "" Then CatPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(p): Exit Function
    If SX_Quiet Then Exit Function
    f = Application.GetOpenFilename("Excel (*.xlsx),*.xlsx", , "Где лежит " & CAT_FILE & "?")
    If VarType(f) = vbBoolean Then Exit Function
    SaveSetting REG_APP, "Catalog", "Path", CStr(f)
    CatPath = CStr(f)
End Function

' Уже открытая книга с этим путём -> wbOut. False — открыт одноимённый файл из другой папки (его путь в other).
Private Function FindOpen(path As String, ByRef wbOut As Workbook, ByRef other As String) As Boolean
    Dim fso As Object
    Set wbOut = Nothing
    On Error Resume Next
    Set wbOut = Workbooks(Mid$(path, InStrRev(path, "\") + 1))
    On Error GoTo 0
    FindOpen = True
    If wbOut Is Nothing Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If LCase$(fso.GetAbsolutePathName(wbOut.FullName)) <> LCase$(fso.GetAbsolutePathName(path)) Then
        other = wbOut.FullName
        Set wbOut = Nothing
        FindOpen = False
    End If
End Function

' Прочитать общий каталог (A3:Q…). Открытый пользователем файл не закрывается. Ошибка -> errTxt.
Private Function ReadCommon(path As String, ByRef hdrRow As Variant, ByRef errTxt As String) As Variant
    Dim wb As Workbook, opened As Boolean, ws As Worksheet, last As Long, other As String
    If Not FindOpen(path, wb, other) Then
        errTxt = "Открыт другой файл с тем же именем:" & vbLf & other & vbLf & "Закройте его и повторите."
        Exit Function
    End If
    On Error GoTo fail
    If wb Is Nothing Then
        Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True)
        opened = True
    End If
    On Error Resume Next
    Set ws = wb.Worksheets(CAT)
    On Error GoTo fail
    If ws Is Nothing Then
        errTxt = "В общем каталоге нет листа «" & CAT & "»: " & path
    Else
        last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        hdrRow = ws.Range(ws.Cells(HDR, 1), ws.Cells(HDR, NCOL)).Value
        If last > HDR Then ReadCommon = ws.Range(ws.Cells(R0, 1), ws.Cells(last, NCOL)).Value
    End If
    If opened Then wb.Close False
    Exit Function
fail:
    errTxt = "Не удалось прочитать общий каталог: " & Err.Description & vbLf & path
    On Error Resume Next
    If opened Then wb.Close False
End Function

' ---------------------------------------------------------------- главная
Private Sub Cat_Run(forced As String)
    Dim wb As Workbook, ws As Worksheet, path As String
    Dim src As Variant, sh As Variant, bk As Variant, bh As Variant
    Dim idx As Object, seen As Object, k As String
    Dim i As Long, c As Long, r As Long, nSrc As Long
    Dim addRows As Collection, updCells As Collection, upd As Object
    Dim freeRows As Collection, onlyBook As Long, onlyList As String, dupList As String
    Dim addList As String, updList As String, msg As String, item As Variant, errTxt As String
    Dim calc As XlCalculation, prot As Boolean, su As Boolean, j As Long, blank As Boolean

    Set wb = ActiveWorkbook
    If wb Is Nothing Then CSay "Откройте таблицу однолинейки.", vbExclamation: Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets(CAT)
    On Error GoTo 0
    If ws Is Nothing Then CSay "В книге нет листа «" & CAT & "».", vbExclamation: Exit Sub
    If wb.Name = CAT_FILE Then CSay "Это сам общий каталог. Откройте книгу однолинейки.", vbExclamation: Exit Sub

    path = CatPath(forced)
    If path = "" Then CSay "Общий каталог не найден: " & CAT_FILE, vbExclamation: Exit Sub
    If Dir(path) = "" Then CSay "Нет файла: " & path, vbExclamation: Exit Sub

    su = Application.ScreenUpdating
    Application.ScreenUpdating = False
    src = ReadCommon(path, sh, errTxt)
    wb.Activate
    Application.ScreenUpdating = su
    If errTxt <> "" Then CSay errTxt & vbLf & "Ничего не изменено.", vbExclamation: Exit Sub
    If IsEmpty(src) Then CSay "Общий каталог пуст: " & path, vbExclamation: Exit Sub

    ' шапки должны совпадать по столбцам данных
    bh = ws.Range(ws.Cells(HDR, 1), ws.Cells(HDR, NCOL)).Value
    For c = 1 To NCOL
        If DataCol(c) Then
            If LCase$(Trim$(CStr(sh(1, c)))) <> LCase$(Trim$(CStr(bh(1, c)))) Then
                CSay "Шапка «Каталога» в книге не совпадает с общим каталогом, столбец " & _
                     Split(ws.Cells(1, c).Address, "$")(1) & ":" & vbLf & _
                     "книга: «" & bh(1, c) & "»" & vbLf & "общий: «" & sh(1, c) & "»" & vbLf & _
                     "Ничего не изменено.", vbExclamation
                Exit Sub
            End If
        End If
    Next c

    ' книга: ключ -> строка
    bk = ws.Range(ws.Cells(R0, 1), ws.Cells(R1, NCOL)).Value
    Set idx = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(bk, 1)
        If Trim$(CStr(bk(i, 1))) <> "" Then
            k = NormKey(bk(i, 1), bk(i, 2), bk(i, 3), bk(i, 7))
            If idx.Exists(k) Then
                If Len(dupList) < 400 Then dupList = dupList & vbLf & "  строки " & (R0 + idx(k) - 1) & " и " & _
                    (R0 + i - 1) & ": " & bk(i, 3) & " — " & bk(i, 7)
            Else
                idx.Add k, i
            End If
        End If
    Next i
    If dupList <> "" Then
        CSay "В «Каталоге» книги повторяется позиция (Категория + Производитель + Модель + Наименование):" & _
             dupList & vbLf & vbLf & "Уберите повтор или измените Наименование. Ничего не изменено.", vbExclamation
        Exit Sub
    End If

    ' план: что добавить, что изменить
    Set seen = CreateObject("Scripting.Dictionary")
    Set addRows = New Collection
    Set updCells = New Collection
    Set upd = CreateObject("Scripting.Dictionary")
    nSrc = 0
    For i = 1 To UBound(src, 1)
        If Trim$(CStr(src(i, 1))) <> "" Then
            nSrc = nSrc + 1
            k = NormKey(src(i, 1), src(i, 2), src(i, 3), src(i, 7))
            If seen.Exists(k) Then
                ' повтор в общем: одинаковые данные — пропустить, разные — остановиться
                r = seen(k)
                For c = 1 To NCOL
                    If DataCol(c) Then
                        If Not Same(src(i, c), src(r, c)) Then
                            If Len(dupList) < 400 Then dupList = dupList & vbLf & "  строки " & (R0 + r - 1) & _
                                " и " & (R0 + i - 1) & ": " & src(i, 3) & " — " & src(i, 7)
                            Exit For
                        End If
                    End If
                Next c
            Else
                seen.Add k, i
                If idx.Exists(k) Then
                    r = idx(k)
                    For c = 1 To NCOL
                        If DataCol(c) Then
                            If Not Same(src(i, c), bk(r, c)) Then
                                updCells.Add Array(r, c, src(i, c))
                                If Not upd.Exists(r) Then
                                    upd.Add r, 1
                                    If upd.Count <= 15 Then updList = updList & vbLf & "  " & src(i, 3) & " — " & _
                                        Split(ws.Cells(1, c).Address, "$")(1)
                                End If
                            End If
                        End If
                    Next c
                Else
                    addRows.Add i
                    If addRows.Count <= 15 Then addList = addList & vbLf & "  " & src(i, 1) & ": " & src(i, 3)
                End If
            End If
        End If
    Next i

    If dupList <> "" Then
        CSay "В общем каталоге одна позиция записана дважды с разными данными:" & dupList & vbLf & vbLf & _
             "Исправьте " & path & ". Ничего не изменено.", vbExclamation
        Exit Sub
    End If

    ' только в книге
    For Each item In idx.Keys
        If Not seen.Exists(item) Then
            onlyBook = onlyBook + 1
            r = idx(item)
            If onlyBook <= 15 Then onlyList = onlyList & vbLf & "  " & bk(r, 1) & ": " & bk(r, 3)
        End If
    Next item

    ' место для новых строк: строки, где все столбцы данных A–M, Q пусты (также пропуски внутри списка)
    Set freeRows = New Collection
    If addRows.Count > 0 Then
        For i = 1 To UBound(bk, 1)
            blank = True
            For c = 1 To NCOL
                If DataCol(c) Then
                    If IsError(bk(i, c)) Then blank = False: Exit For
                    If Trim$(CStr(bk(i, c))) <> "" Then blank = False: Exit For
                End If
            Next c
            If blank Then freeRows.Add i
            If freeRows.Count >= addRows.Count Then Exit For
        Next i
        If freeRows.Count < addRows.Count Then
            CSay "В «Каталоге» не хватает строк: нужно " & addRows.Count & ", свободно до строки " & R1 & _
                 " — " & freeRows.Count & "." & vbLf & "Ничего не изменено.", vbExclamation
            Exit Sub
        End If
    End If

    If addRows.Count = 0 And upd.Count = 0 Then
        msg = "Каталог книги совпадает с общим (" & nSrc & " позиций)."
        If onlyBook > 0 Then msg = msg & vbLf & vbLf & "Только в книге — " & onlyBook & ":" & onlyList
        CSay msg
        Exit Sub
    End If

    msg = "Общий каталог: " & path & vbLf & vbLf
    If addRows.Count > 0 Then msg = msg & "Добавить позиций: " & addRows.Count & addList & vbLf & vbLf
    If upd.Count > 0 Then msg = msg & "Изменить позиций: " & upd.Count & " (ячеек " & updCells.Count & ")" & updList & vbLf & vbLf
    If onlyBook > 0 Then msg = msg & "Только в книге (не трогаю): " & onlyBook & onlyList & vbLf & vbLf
    If Not SX_Quiet Then
        If MsgBox(msg & "Запас (O) и формулы не меняются. Отменить через Ctrl+Z нельзя." & vbLf & _
                  "Обновить?", vbQuestion + vbYesNo, "Каталог") <> vbYes Then Exit Sub
    End If

    ' запись
    calc = Application.Calculation
    prot = ws.ProtectContents
    On Error GoTo fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    If prot Then ws.Unprotect PWD
    For Each item In updCells
        PutVal ws.Cells(R0 + item(0) - 1, item(1)), item(2)
    Next item
    j = 0
    For Each item In addRows
        j = j + 1
        r = freeRows(j)
        For c = 1 To NCOL
            If DataCol(c) Then PutVal ws.Cells(R0 + r - 1, c), src(item, c)
        Next c
    Next item
    If prot Then ws.Protect PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    Application.Calculation = calc
    Application.ScreenUpdating = su

    msg = "Каталог обновлён: добавлено " & addRows.Count & ", изменено " & upd.Count & "."
    If onlyBook > 0 Then msg = msg & vbLf & "Только в книге: " & onlyBook & "."
    msg = msg & vbLf & "Спецификацию обновите кнопкой «Обновить спецификацию»."
    CSay msg
    Exit Sub
fail:
    msg = "Ошибка при записи «Каталога» (строка " & Erl & "): " & Err.Description & vbLf & _
          "Часть позиций могла записаться. Закройте книгу без сохранения, если нужно."
    On Error Resume Next
    If prot Then ws.Protect PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    Application.Calculation = calc
    Application.ScreenUpdating = su
    CSay msg, vbCritical
End Sub

' ---------------------------------------------------------------- книга -> общий
Private Sub Cat_Push(forced As String, withChanged As Boolean)
    Dim wb As Workbook, ws As Worksheet, cwb As Workbook, cws As Worksheet, path As String, nm As String
    Dim bk As Variant, bh As Variant, src As Variant, sh As Variant
    Dim idx As Object, seen As Object, k As String
    Dim i As Long, c As Long, r As Long, last As Long, opened As Boolean, su As Boolean
    Dim addRows As Collection, updCells As Collection, upd As Object
    Dim addList As String, updList As String, msg As String, item As Variant, doChanged As Boolean

    Set wb = ActiveWorkbook
    If wb Is Nothing Then CSay "Откройте таблицу однолинейки.", vbExclamation: Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets(CAT)
    On Error GoTo 0
    If ws Is Nothing Then CSay "В книге нет листа «" & CAT & "».", vbExclamation: Exit Sub
    If wb.Name = CAT_FILE Then CSay "Это сам общий каталог. Откройте книгу однолинейки.", vbExclamation: Exit Sub

    path = CatPath(forced)
    If path = "" Then CSay "Общий каталог не найден: " & CAT_FILE, vbExclamation: Exit Sub
    If Dir(path) = "" Then CSay "Нет файла: " & path, vbExclamation: Exit Sub

    If Not FindOpen(path, cwb, nm) Then
        CSay "Открыт другой файл с тем же именем:" & vbLf & nm & vbLf & "Закройте его и повторите. Ничего не изменено.", vbExclamation
        Exit Sub
    End If
    su = Application.ScreenUpdating
    Application.ScreenUpdating = False
    If cwb Is Nothing Then
        On Error Resume Next
        Set cwb = Workbooks.Open(path, UpdateLinks:=0)
        On Error GoTo 0
        opened = True
    End If
    wb.Activate
    Application.ScreenUpdating = su
    If cwb Is Nothing Then CSay "Не удалось открыть общий каталог: " & path, vbExclamation: Exit Sub
    On Error Resume Next
    Set cws = cwb.Worksheets(CAT)
    On Error GoTo 0
    If cws Is Nothing Then
        If opened Then cwb.Close False
        CSay "В общем каталоге нет листа «" & CAT & "». Ничего не изменено.", vbExclamation
        Exit Sub
    End If
    If cwb.ReadOnly Then
        If opened Then cwb.Close False
        CSay "Общий каталог открыт только для чтения (занят?): " & path & vbLf & "Ничего не изменено.", vbExclamation
        Exit Sub
    End If
    last = cws.Cells(cws.Rows.Count, 1).End(xlUp).Row
    If last < HDR Then last = HDR
    sh = cws.Range(cws.Cells(HDR, 1), cws.Cells(HDR, NCOL)).Value
    If last > HDR Then src = cws.Range(cws.Cells(R0, 1), cws.Cells(last, NCOL)).Value

    bh = ws.Range(ws.Cells(HDR, 1), ws.Cells(HDR, NCOL)).Value
    For c = 1 To NCOL
        If DataCol(c) Then
            If LCase$(Trim$(CStr(sh(1, c)))) <> LCase$(Trim$(CStr(bh(1, c)))) Then
                If opened Then cwb.Close False
                CSay "Шапка «Каталога» в книге не совпадает с общим каталогом, столбец " & _
                     Split(ws.Cells(1, c).Address, "$")(1) & ". Ничего не изменено.", vbExclamation
                Exit Sub
            End If
        End If
    Next c

    ' общий: ключ -> строка
    Set idx = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(src) Then
        For i = 1 To UBound(src, 1)
            If Trim$(CStr(src(i, 1))) <> "" Then
                k = NormKey(src(i, 1), src(i, 2), src(i, 3), src(i, 7))
                If Not idx.Exists(k) Then idx.Add k, i
            End If
        Next i
    End If

    ' книга: новые и разные
    bk = ws.Range(ws.Cells(R0, 1), ws.Cells(R1, NCOL)).Value
    Set seen = CreateObject("Scripting.Dictionary")
    Set addRows = New Collection
    Set updCells = New Collection
    Set upd = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(bk, 1)
        If Trim$(CStr(bk(i, 1))) <> "" Then
            k = NormKey(bk(i, 1), bk(i, 2), bk(i, 3), bk(i, 7))
            If Not seen.Exists(k) Then
                seen.Add k, i
                If idx.Exists(k) Then
                    r = idx(k)
                    For c = 1 To NCOL
                        If DataCol(c) Then
                            If Not IsError(bk(i, c)) Then
                                If Not Same(bk(i, c), src(r, c)) Then
                                    updCells.Add Array(r, c, bk(i, c))
                                    If Not upd.Exists(r) Then
                                        upd.Add r, 1
                                        If upd.Count <= 15 Then updList = updList & vbLf & "  " & bk(i, 3) & " — " & _
                                            Split(ws.Cells(1, c).Address, "$")(1) & ": «" & src(r, c) & "» -> «" & bk(i, c) & "»"
                                    End If
                                End If
                            End If
                        End If
                    Next c
                Else
                    addRows.Add i
                    If addRows.Count <= 15 Then addList = addList & vbLf & "  " & bk(i, 1) & ": " & bk(i, 3) & " — " & bk(i, 7)
                End If
            End If
        End If
    Next i

    If addRows.Count = 0 And upd.Count = 0 Then
        If opened Then cwb.Close False
        CSay "Всё, что есть в книге, уже есть в общем каталоге."
        Exit Sub
    End If

    doChanged = withChanged
    If Not SX_Quiet Then
        If addRows.Count > 0 Then
            msg = "Общий каталог: " & path & vbLf & vbLf & "Дописать новых позиций: " & addRows.Count & addList & vbLf & vbLf
            If upd.Count > 0 Then msg = msg & "Ещё " & upd.Count & " позиций в книге отличаются — спрошу отдельно." & vbLf & vbLf
            If MsgBox(msg & "Дописать?", vbQuestion + vbYesNo, "Каталог") <> vbYes Then Set addRows = New Collection
        End If
        If upd.Count > 0 Then
            doChanged = (MsgBox("В книге иначе, чем в общем каталоге (" & upd.Count & " позиций, ячеек " & updCells.Count & "):" & _
                updList & vbLf & vbLf & "Записать данные книги в общий каталог?" & vbLf & _
                "«Нет» — общий не меняется (книгу можно обновить кнопкой «Каталог: обновить»).", _
                vbQuestion + vbYesNo + vbDefaultButton2, "Каталог") = vbYes)
        End If
    End If
    If addRows.Count = 0 And Not doChanged Then
        If opened Then cwb.Close False
        CSay "Общий каталог не изменён."
        Exit Sub
    End If

    On Error GoTo fail
    If doChanged Then
        For Each item In updCells
            PutVal cws.Cells(R0 + item(0) - 1, item(1)), item(2)
        Next item
    End If
    r = last + 1
    For Each item In addRows
        For c = 1 To NCOL
            If DataCol(c) Then
                cws.Cells(r, c).NumberFormat = ws.Cells(R0 + item - 1, c).NumberFormat
                PutVal cws.Cells(r, c), bk(item, c)
            End If
        Next c
        r = r + 1
    Next item
    cws.AutoFilterMode = False
    cws.Range(cws.Cells(HDR, 1), cws.Cells(r - 1, NCOL)).AutoFilter
    cwb.Save
    If opened Then cwb.Close False
    wb.Activate

    msg = "Общий каталог: дописано " & addRows.Count
    If doChanged Then msg = msg & ", изменено " & upd.Count
    msg = msg & "." & vbLf & "Отправьте в GitHub («Отправить.cmd»), чтобы каталог был и на другом ПК."
    CSay msg
    Exit Sub
fail:
    msg = "Ошибка записи в общий каталог (строка " & Erl & "): " & Err.Description & vbLf & "Файл не сохранён."
    On Error Resume Next
    If opened Then cwb.Close False
    CSay msg, vbCritical
End Sub
