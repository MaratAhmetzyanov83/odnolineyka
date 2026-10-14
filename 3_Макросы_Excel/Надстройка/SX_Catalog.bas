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

' Для проверки программой: путь к общему каталогу задан, окон нет (SX_SetQuiet). Возвращает итоговое сообщение.
Public Function SX_CatUpdateFrom(path As Variant) As String
    Cat_Run CStr(path)
    SX_CatUpdateFrom = SX_LastMsg
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

Private Function Same(x As Variant, y As Variant) As Boolean
    If IsError(x) Or IsError(y) Then Same = False: Exit Function
    If IsNumeric(x) And IsNumeric(y) And Not IsEmpty(x) And Not IsEmpty(y) Then
        If Trim$(CStr(x)) <> "" And Trim$(CStr(y)) <> "" Then Same = (CDbl(x) = CDbl(y)): Exit Function
    End If
    Same = (Trim$(CStr(x)) = Trim$(CStr(y)))
End Function

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

' Прочитать общий каталог (A3:Q…). Открытый файл не закрывается.
Private Function ReadCommon(path As String, ByRef hdrRow As Variant) As Variant
    Dim wb As Workbook, opened As Boolean, ws As Worksheet, last As Long, nm As String
    nm = Mid$(path, InStrRev(path, "\") + 1)
    On Error Resume Next
    Set wb = Workbooks(nm)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True)
        opened = True
    End If
    Set ws = wb.Worksheets(CAT)
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    hdrRow = ws.Range(ws.Cells(HDR, 1), ws.Cells(HDR, NCOL)).Value
    If last > HDR Then ReadCommon = ws.Range(ws.Cells(R0, 1), ws.Cells(last, NCOL)).Value
    If opened Then wb.Close False
End Function

' ---------------------------------------------------------------- главная
Private Sub Cat_Run(forced As String)
    Dim wb As Workbook, ws As Worksheet, path As String
    Dim src As Variant, sh As Variant, bk As Variant, bh As Variant
    Dim idx As Object, seen As Object, k As String
    Dim i As Long, c As Long, r As Long, nSrc As Long
    Dim addRows As Collection, updCells As Collection, upd As Object
    Dim freeRow As Long, onlyBook As Long, onlyList As String, dupList As String
    Dim addList As String, updList As String, msg As String, item As Variant
    Dim calc As XlCalculation, prot As Boolean

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

    Application.ScreenUpdating = False
    src = ReadCommon(path, sh)
    wb.Activate
    Application.ScreenUpdating = True
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
            If Not idx.Exists(k) Then idx.Add k, i
        End If
    Next i

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
                If Len(dupList) < 400 Then dupList = dupList & vbLf & "  " & src(i, 3) & " — " & src(i, 7)
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

    ' только в книге
    For Each item In idx.Keys
        If Not seen.Exists(item) Then
            onlyBook = onlyBook + 1
            r = idx(item)
            If onlyBook <= 15 Then onlyList = onlyList & vbLf & "  " & bk(r, 1) & ": " & bk(r, 3)
        End If
    Next item

    ' место для новых строк
    If addRows.Count > 0 Then
        r = 0
        For i = UBound(bk, 1) To 1 Step -1
            If Trim$(CStr(bk(i, 1))) <> "" Then r = i: Exit For
        Next i
        freeRow = r + 1                              ' индекс в bk, после последней заполненной
        If freeRow + addRows.Count - 1 > UBound(bk, 1) Then
            CSay "В «Каталоге» не хватает строк: нужно " & addRows.Count & ", свободно до строки " & R1 & _
                 " — " & (UBound(bk, 1) - freeRow + 1) & "." & vbLf & "Ничего не изменено.", vbExclamation
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
    If dupList <> "" Then msg = msg & "Повтор в общем каталоге (взята первая строка):" & dupList & vbLf & vbLf
    If Not SX_Quiet Then
        If MsgBox(msg & "Запас (O) и формулы не меняются. Отменить через Ctrl+Z нельзя." & vbLf & _
                  "Обновить?", vbQuestion + vbYesNo, "Каталог") <> vbYes Then Exit Sub
    End If

    ' запись
    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    prot = ws.ProtectContents
    On Error GoTo fail
    If prot Then ws.Unprotect PWD
    For Each item In updCells
        ws.Cells(R0 + item(0) - 1, item(1)).Value = item(2)
    Next item
    r = freeRow
    For Each item In addRows
        For c = 1 To NCOL
            If DataCol(c) Then ws.Cells(R0 + r - 1, c).Value = src(item, c)
        Next c
        r = r + 1
    Next item
    If prot Then ws.Protect PWD, DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    Application.Calculation = calc
    Application.ScreenUpdating = True

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
    Application.ScreenUpdating = True
    CSay msg, vbCritical
End Sub
