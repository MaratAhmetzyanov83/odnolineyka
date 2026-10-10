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

Public Sub SX_SetQuiet(Optional q As Boolean = True)
    SX_Quiet = q
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
Public Sub SX_Ribbon(control As Object)
    Select Case control.ID
        Case "sxSpec": SX_UpdateSpec
        Case "sxAccept": SX_AcceptChanges
        Case "sxGroup": SX_InsertGroup
        Case "sxSaveGroup": SX_SaveAsGroup
        Case "sxTags": SX_Tags
        Case "sxLabels": SX_Labels
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
' лист «Бирки» в v3.12 считается сам формулами — открыть его
Public Sub SX_Tags()
    Dim wb As Workbook: Set wb = Book(): If wb Is Nothing Then Exit Sub
    If Not HasSheet(wb, "Бирки") Then Say "В книге нет листа «Бирки».", vbExclamation: Exit Sub
    wb.Sheets("Бирки").Visible = xlSheetVisible
    wb.Sheets("Бирки").Activate
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

' ---------------------------------------------------------------- справка
Public Sub SX_Help()
    MsgBox "Вкладка «Однолинейка» (таблица v3.12):" & vbLf & vbLf & _
        "• Обновить спецификацию — «Спецификация_авто» переносится в «Спецификация», изменения против принятой ревизии подсвечиваются, удалённые — в листе «Изменения»." & vbLf & _
        "• Принять изменения — когда снабжение получило: подсветка снимается, ревизия +1." & vbLf & _
        "• Типовая группа — выделите строки на листе «Однолинейка» и вставьте группу; «Сохранить как группу» — запомнить свою." & vbLf & _
        "• Бирки — лист «Бирки» (считается сам)." & vbLf & _
        "• Наклейки — лента под аппаратами, клетка = 1 модуль DIN. Проверьте пробной печатью." & vbLf & vbLf & _
        "Подробно — 4_Инструкции/Надстройка_на_ленте.md", vbInformation, "Однолинейка"
End Sub
