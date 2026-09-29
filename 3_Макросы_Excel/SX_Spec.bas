Attribute VB_Name = "SX_Spec"
Option Explicit

' ============================================================================
' SX_Spec - обновление спецификации и подсветка изменений для снабжения
' Однолинейка (электрощит), книги v3.11 (.xlsx, макросов внутри книг нет).
'
' Куда ставить: PERSONAL.XLSB или надстройка .xlam. Все процедуры работают
' с ActiveWorkbook (никогда с ThisWorkbook), поэтому подходят любой книге
' однолинейки: шаблон, пример, щиты объекта, ГРЩ.
'
' Публичные макросы:
'   SX_UpdateSpec / ОбновитьСпецификацию      - обновить и подсветить
'   SX_AcceptChanges / ПринятьИзменения       - принять ревизию
'
' Как это работает:
'   1. Лист "Спецификация_авто" B4:J160 (черновик, формулы) переносится
'      в лист "Спецификация" B4:J160 как ЗНАЧЕНИЯ (присваивание массива .Value,
'      без буфера обмена). Форматы листа "Спецификация" не трогаются.
'   2. Скрытый лист "Спец_принятая" хранит последнее ПРИНЯТОЕ состояние
'      (те же адреса B4:J160, значения как текст). A1 = метка, B1 = номер
'      принятой ревизии, C1 = дата принятия.
'   3. После каждого обновления итог сравнивается с принятым снимком.
'      Ключ позиции: артикул (E) | наименование (C) | тип/марка (D),
'      без учёта регистра и лишних пробелов; одинаковые ключи различаются
'      порядковым номером. Если ключ не нашёлся, но артикул встречается по
'      одному разу среди несопоставленных строк и там, и там, строки
'      считаются одной позицией (например, правка наименования).
'      Колонка "Поз" (B) в сравнении не участвует: номер сдвигается при
'      вставке позиции выше и давал бы ложные изменения.
'   4. Изменённые ячейки - жёлтые, новые строки - зелёные. Удалённые позиции
'      попадают в журнал (лист "Изменения"). Подсветка копится, пока
'      не выполнена команда "Принять", т.к. сравнение всегда с принятым снимком.
'   5. Журнал "Изменения": принятые ревизии остаются навсегда, непринятые
'      отличия (в колонке "Ревизия" стоит "не принято") пересобираются при
'      каждом обновлении. При "Принять" они получают номер ревизии и дату.
'
' Кодировка файла: Windows-1251 (ANSI русской Windows), окончания строк CRLF.
' ============================================================================

Private Const SH_SRC As String = "Спецификация_авто"
Private Const SH_DST As String = "Спецификация"
Private Const SH_SNAP As String = "Спец_принятая"
Private Const SH_LOG As String = "Изменения"
Private Const SH_CAT As String = "Каталог"
Private Const SH_ODN As String = "Однолинейка"

Private Const SHEET_PWD As String = "sx"       ' пароль защиты листов однолинейки
Private Const ROW_FIRST As Long = 4            ' итоговая область B4:J160
Private Const ROW_LAST As Long = 160
Private Const COL_FIRST As Long = 2            ' B
Private Const COL_LAST As Long = 10            ' J
Private Const NCOL As Long = 9                 ' число колонок B:J
Private Const STAMP_ADDR As String = "C1"      ' заметная ячейка на листе "Спецификация"
Private Const SNAP_MARK As String = "SX_SNAPSHOT_V1"

' номера колонок внутри массивов B:J
Private Const IX_POS As Long = 1               ' B  Поз
Private Const IX_NAME As Long = 2              ' C  Наименование
Private Const IX_MODEL As Long = 3             ' D  Тип, марка
Private Const IX_ART As Long = 4               ' E  Код (артикул)
Private Const IX_UNIT As Long = 6              ' G  Ед. изм.
Private Const IX_QTY As Long = 7               ' H  Кол-во

' журнал "Изменения"
Private Const LOG_HDR_ROW As Long = 3
Private Const LOG_FIRST As Long = 4
Private Const LOGCOLS As Long = 9
Private Const PENDING As String = "не принято"
Private Const TYP_CHG As String = "изменено"
Private Const TYP_NEW As String = "добавлено"
Private Const TYP_DEL As String = "удалено"
Private Const FLD_ALL As String = "Позиция целиком"

Private Type AppState
    Calc As Long
    Scr As Boolean
    Evt As Boolean
    Bar As Variant
End Type

Private Type ProtInfo
    WasProtected As Boolean
    Unlocked As Boolean
    Pwd As String
    Contents As Boolean
    Draw As Boolean
    Scen As Boolean
    Allow(1 To 11) As Boolean
End Type

Private Type DiffResult
    NChanged As Long            ' позиций с изменёнными ячейками
    NCells As Long              ' изменённых ячеек
    NNew As Long
    NDel As Long
    RowState() As Byte          ' по строкам итога: 0 без изменений, 1 изменена, 2 новая
    CellMask() As Boolean       ' изменённые ячейки (строка, колонка 1..9)
    Entries() As Variant        ' строки журнала (запись, 1..9)
    LogN As Long
End Type

' ----------------------------------------------------------------------------
' Русские псевдонимы макросов (для окна Alt+F8 и кнопок)
' ----------------------------------------------------------------------------
Public Sub ОбновитьСпецификацию()
    SX_UpdateSpec
End Sub

Public Sub ПринятьИзменения()
    SX_AcceptChanges
End Sub

' ============================================================================
' ОБНОВИТЬ СПЕЦИФИКАЦИЮ
' ============================================================================
Public Sub SX_UpdateSpec()
    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsDst As Worksheet, wsSnap As Worksheet, wsLog As Worksheet
    Dim wsActive As Object
    Dim st As AppState
    Dim pDst As ProtInfo
    Dim d As DiffResult
    Dim rngDst As Range
    Dim arr As Variant, cur As Variant, old As Variant, hdr As Variant
    Dim problem As String, msg As String, errText As String, warnText As String
    Dim kind As Long                ' 0 успех, 1 ошибка, 2 отмена
    Dim stateSaved As Boolean, firstRun As Boolean
    Dim n As Long, m As Long, nItems As Long, nErr As Long, lastRow As Long
    Dim rev As Long
    Dim stamp As Date
    Dim ans As VbMsgBoxResult

    On Error Resume Next
    Set wb = ActiveWorkbook
    On Error GoTo 0
    If wb Is Nothing Then
        MsgBox "Нет открытой книги.", vbExclamation, "Обновить спецификацию"
        Exit Sub
    End If
    problem = CheckBook(wb)
    If Len(problem) > 0 Then
        MsgBox problem, vbExclamation, "Обновить спецификацию"
        Exit Sub
    End If

    On Error GoTo EH
    SaveState st
    stateSaved = True
    Set wsActive = Nothing
    On Error Resume Next
    Set wsActive = wb.ActiveSheet
    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Обновление спецификации..."
    Application.Calculation = xlCalculationManual

    Set wsSrc = SheetOf(wb, SH_SRC)
    Set wsDst = SheetOf(wb, SH_DST)

    ' --- полный пересчёт (черновик зависит от Каталога и Однолинейки) ---
    Application.CalculateFull

    ' --- читаем черновик одним массивом ---
    arr = wsSrc.Range(wsSrc.Cells(ROW_FIRST, COL_FIRST), wsSrc.Cells(ROW_LAST, COL_LAST)).Value
    lastRow = PrepareForWrite(arr, wsDst, nItems, nErr)
    warnText = Norm(wsSrc.Range("C1").Value)
    If InStr(1, warnText, "ВНИМАНИЕ", vbBinaryCompare) <> 1 Then warnText = ""

    If lastRow = 0 Then
        Application.ScreenUpdating = True
        ans = MsgBox("Черновик «Спецификация_авто» пуст. Очистить итоговую область " & _
                     "B4:J160 в листе «Спецификация»?", vbYesNo + vbQuestion + vbDefaultButton2, _
                     "Обновить спецификацию")
        Application.ScreenUpdating = False
        If ans <> vbYes Then
            kind = 2
            msg = "Обновление отменено, лист «Спецификация» не изменён."
            GoTo Fin
        End If
    End If

    ' --- служебные листы (создаются при первом запуске) ---
    If wb.ProtectStructure Then
        If SheetOf(wb, SH_LOG) Is Nothing Or SheetOf(wb, SH_SNAP) Is Nothing Then
            Err.Raise vbObjectError + 513, "SX_Spec", "Структура книги защищена, нельзя добавить " & _
                "служебные листы «" & SH_LOG & "» и «" & SH_SNAP & "». Снимите защиту структуры " & _
                "(Рецензирование - Защитить книгу) и повторите."
        End If
    End If
    Set wsSnap = SheetOf(wb, SH_SNAP)
    If Not wsSnap Is Nothing Then
        If Not SnapValid(wsSnap) Then
            Err.Raise vbObjectError + 515, "SX_Spec", "Лист «" & SH_SNAP & "» существует, но не является " & _
                "снимком SX_Spec (нет метки в A1). Переименуйте его и повторите."
        End If
    End If

    ' --- защита листа "Спецификация" ---
    If Not UnprotectSheet(wsDst, pDst) Then
        Err.Raise vbObjectError + 514, "SX_Spec", "Лист «" & SH_DST & "» защищён неизвестным паролем. " & _
            "Снимите защиту вручную и повторите."
    End If

    ' --- очистка ТОЛЬКО значений (форматы остаются), снятие своей подсветки ---
    Set rngDst = wsDst.Range(wsDst.Cells(ROW_FIRST, COL_FIRST), wsDst.Cells(ROW_LAST, COL_LAST))
    ClearOurFills rngDst
    rngDst.ClearContents

    ' --- перенос значений (не формул) ---
    rngDst.Value = arr

    ' --- сравнение с принятым снимком ---
    stamp = Now
    cur = rngDst.Value
    n = LastFilled(cur, ROW_LAST - ROW_FIRST + 1)
    Set wsLog = EnsureLog(wb, wsDst)
    If wsSnap Is Nothing Then
        Set wsSnap = CreateSnap(wb, wsLog)
        SaveSnap wsSnap, cur, 0, stamp
        firstRun = True
        rev = 0
        InitDiff d, n, 0
    Else
        rev = CLng(Val(CStr(wsSnap.Range("B1").Value)))
        old = wsSnap.Range(wsSnap.Cells(ROW_FIRST, COL_FIRST), wsSnap.Cells(ROW_LAST, COL_LAST)).Value
        m = LastFilled(old, ROW_LAST - ROW_FIRST + 1)
        hdr = wsDst.Range(wsDst.Cells(2, COL_FIRST), wsDst.Cells(2, COL_LAST)).Value
        BuildDiff cur, n, old, m, hdr, stamp, d
        ApplyFills wsDst, d, n
    End If
    WriteLog wsLog, d
    SetLogStatus wsLog, rev, d, stamp

    wsDst.Range(STAMP_ADDR).Value = StampText(nItems, rev, d, stamp, False)

    msg = nItems & " " & PosWord(nItems) & vbCrLf & "Обновлено: " & Format$(stamp, "dd.mm.yyyy hh:nn")
    If firstRun Then
        msg = msg & vbCrLf & vbCrLf & "Это первый запуск: текущее состояние записано как принятая " & _
              "ревизия 0 (базовый снимок), подсветки нет. Изменения будут подсвечиваться " & _
              "относительно него."
    Else
        msg = msg & vbCrLf & vbCrLf & "Принятая ревизия: " & rev & vbCrLf & _
              "Изменено позиций: " & d.NChanged & " (ячеек: " & d.NCells & ", жёлтые)" & vbCrLf & _
              "Добавлено: " & d.NNew & " (зелёные)" & vbCrLf & _
              "Удалено: " & d.NDel & " (см. лист «" & SH_LOG & "»)"
    End If
    If nErr > 0 Then msg = msg & vbCrLf & vbCrLf & "Ячеек с ошибками в черновике: " & nErr & _
        " (записаны текстом #ОШИБКА) - проверьте «Спецификация_авто»."
    If Len(warnText) > 0 Then msg = msg & vbCrLf & vbCrLf & warnText
    msg = msg & vbCrLf & vbCrLf & "Сохраните книгу (Ctrl+S): снимок и журнал хранятся в ней."
    kind = 0
    GoTo Fin

EH:
    kind = 1
    errText = "Ошибка " & Err.Number & ": " & Err.Description
Fin:
    On Error Resume Next
    If pDst.Unlocked Then ReprotectSheet wsDst, pDst
    If Not wsActive Is Nothing Then wsActive.Activate
    If stateSaved Then RestoreState st
    On Error GoTo 0

    Select Case kind
        Case 0
            MsgBox msg, vbInformation, "Обновить спецификацию"
        Case 2
            MsgBox msg, vbInformation, "Обновить спецификацию"
        Case Else
            MsgBox "Не удалось обновить спецификацию." & vbCrLf & vbCrLf & errText & vbCrLf & vbCrLf & _
                   "Настройки Excel (пересчёт, события, обновление экрана, защита листа) восстановлены. " & _
                   "Книга могла остаться в промежуточном состоянии: проверьте лист «Спецификация».", _
                   vbCritical, "Обновить спецификацию"
    End Select
End Sub

' ============================================================================
' ПРИНЯТЬ ИЗМЕНЕНИЯ
' ============================================================================
Public Sub SX_AcceptChanges()
    Dim wb As Workbook
    Dim wsDst As Worksheet, wsSnap As Worksheet, wsLog As Worksheet
    Dim wsActive As Object
    Dim st As AppState
    Dim pDst As ProtInfo
    Dim d As DiffResult
    Dim rngDst As Range
    Dim cur As Variant, old As Variant, hdr As Variant
    Dim problem As String, msg As String, errText As String
    Dim kind As Long                ' 0 принято, 1 ошибка, 2 без действий
    Dim stateSaved As Boolean
    Dim n As Long, m As Long, nItems As Long, rev As Long
    Dim stamp As Date
    Dim ans As VbMsgBoxResult

    On Error Resume Next
    Set wb = ActiveWorkbook
    On Error GoTo 0
    If wb Is Nothing Then
        MsgBox "Нет открытой книги.", vbExclamation, "Принять изменения"
        Exit Sub
    End If
    problem = CheckBook(wb)
    If Len(problem) > 0 Then
        MsgBox problem, vbExclamation, "Принять изменения"
        Exit Sub
    End If
    Set wsSnap = SheetOf(wb, SH_SNAP)
    If wsSnap Is Nothing Then
        MsgBox "Принятого снимка ещё нет. Сначала выполните «Обновить спецификацию» " & _
               "(при первом запуске снимок создаётся автоматически).", vbExclamation, "Принять изменения"
        Exit Sub
    End If
    If Not SnapValid(wsSnap) Then
        MsgBox "Лист «" & SH_SNAP & "» повреждён (нет метки в A1). Принять изменения нельзя.", _
               vbExclamation, "Принять изменения"
        Exit Sub
    End If

    On Error GoTo EH
    SaveState st
    stateSaved = True
    Set wsActive = Nothing
    On Error Resume Next
    Set wsActive = wb.ActiveSheet
    On Error GoTo EH

    Set wsDst = SheetOf(wb, SH_DST)
    Set wsLog = SheetOf(wb, SH_LOG)
    If wsLog Is Nothing Then
        If wb.ProtectStructure Then
            Err.Raise vbObjectError + 513, "SX_Spec", "Структура книги защищена, нельзя создать лист «" & SH_LOG & "»."
        End If
    End If

    ' --- сравнение текущего итога с принятым снимком (без переноса) ---
    Set rngDst = wsDst.Range(wsDst.Cells(ROW_FIRST, COL_FIRST), wsDst.Cells(ROW_LAST, COL_LAST))
    cur = rngDst.Value
    n = LastFilled(cur, ROW_LAST - ROW_FIRST + 1)
    nItems = CountItems(cur, n)
    old = wsSnap.Range(wsSnap.Cells(ROW_FIRST, COL_FIRST), wsSnap.Cells(ROW_LAST, COL_LAST)).Value
    m = LastFilled(old, ROW_LAST - ROW_FIRST + 1)
    hdr = wsDst.Range(wsDst.Cells(2, COL_FIRST), wsDst.Cells(2, COL_LAST)).Value
    rev = CLng(Val(CStr(wsSnap.Range("B1").Value)))
    stamp = Now
    BuildDiff cur, n, old, m, hdr, stamp, d

    If d.NChanged + d.NNew + d.NDel = 0 Then
        kind = 2
        msg = "Отличий от принятой ревизии " & rev & " нет, принимать нечего."
        GoTo Fin
    End If

    ans = MsgBox("Принять ревизию " & (rev + 1) & "?" & vbCrLf & vbCrLf & _
                 "Изменено позиций: " & d.NChanged & " (ячеек: " & d.NCells & ")" & vbCrLf & _
                 "Добавлено: " & d.NNew & vbCrLf & _
                 "Удалено: " & d.NDel & vbCrLf & vbCrLf & _
                 "Подсветка будет снята, текущий итог станет новой принятой ревизией, " & _
                 "записи журнала получат номер " & (rev + 1) & ".", _
                 vbYesNo + vbQuestion + vbDefaultButton2, "Принять изменения")
    If ans <> vbYes Then
        kind = 2
        msg = "Принятие отменено, ничего не изменено."
        GoTo Fin
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Принятие ревизии..."
    Application.Calculation = xlCalculationManual

    If Not UnprotectSheet(wsDst, pDst) Then
        Err.Raise vbObjectError + 514, "SX_Spec", "Лист «" & SH_DST & "» защищён неизвестным паролем. " & _
            "Снимите защиту вручную и повторите."
    End If

    Set wsLog = EnsureLog(wb, wsDst)
    WriteLog wsLog, d                       ' актуальный набор непринятых отличий
    rev = rev + 1
    SaveSnap wsSnap, cur, rev, stamp        ' новый принятый снимок
    FreezeLog wsLog, rev, stamp             ' строки "не принято" получают номер ревизии
    ClearOurFills rngDst                    ' снять жёлтую и зелёную заливку
    InitDiff d, 0, 0
    SetLogStatus wsLog, rev, d, stamp
    wsDst.Range(STAMP_ADDR).Value = StampText(nItems, rev, d, stamp, True)

    kind = 0
    msg = "Ревизия " & rev & " принята (" & Format$(stamp, "dd.mm.yyyy hh:nn") & ")." & vbCrLf & _
          "Подсветка снята." & vbCrLf & vbCrLf & "Сохраните книгу (Ctrl+S)."
    GoTo Fin

EH:
    kind = 1
    errText = "Ошибка " & Err.Number & ": " & Err.Description
Fin:
    On Error Resume Next
    If pDst.Unlocked Then ReprotectSheet wsDst, pDst
    If Not wsActive Is Nothing Then wsActive.Activate
    If stateSaved Then RestoreState st
    On Error GoTo 0

    Select Case kind
        Case 0, 2
            MsgBox msg, vbInformation, "Принять изменения"
        Case Else
            MsgBox "Не удалось принять изменения." & vbCrLf & vbCrLf & errText & vbCrLf & vbCrLf & _
                   "Настройки Excel восстановлены. Выполните «Обновить спецификацию» и проверьте " & _
                   "листы «" & SH_LOG & "» и «Спецификация».", vbCritical, "Принять изменения"
    End Select
End Sub

' ============================================================================
' ПРОВЕРКА КНИГИ
' ============================================================================
Private Function CheckBook(wb As Workbook) As String
    Dim nm As Variant
    Dim miss As String
    Dim wsA As Worksheet, wsB As Worksheet
    Dim c As Long

    For Each nm In Array(SH_SRC, SH_DST, SH_CAT, SH_ODN)
        If SheetOf(wb, CStr(nm)) Is Nothing Then miss = miss & vbCrLf & "   - " & CStr(nm)
    Next nm
    If Len(miss) > 0 Then
        CheckBook = "Активная книга «" & wb.Name & "» не похожа на однолинейку. Нет листов:" & miss & _
                    vbCrLf & vbCrLf & "Откройте книгу однолинейки v3.11 и запустите макрос снова."
        Exit Function
    End If

    Set wsA = SheetOf(wb, SH_SRC)
    Set wsB = SheetOf(wb, SH_DST)
    For c = COL_FIRST To COL_LAST
        If IsError(wsA.Cells(2, c).Value) Or IsError(wsB.Cells(2, c).Value) Then
            CheckBook = "В заголовках листов «" & SH_SRC & "» / «" & SH_DST & "» (строка 2) ошибки."
            Exit Function
        End If
        If Len(Trim$(CStr(wsA.Cells(2, c).Value))) = 0 Or _
           StrComp(Trim$(CStr(wsA.Cells(2, c).Value)), Trim$(CStr(wsB.Cells(2, c).Value)), vbTextCompare) <> 0 Then
            CheckBook = "Структура листов «" & SH_SRC & "» и «" & SH_DST & "» не совпадает " & _
                        "(заголовки в строке 2, колонки B:J). Макрос рассчитан на однолинейку v3.11."
            Exit Function
        End If
    Next c
End Function

Private Function SheetOf(wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            Set SheetOf = ws
            Exit Function
        End If
    Next ws
End Function

' ============================================================================
' НАСТРОЙКИ EXCEL
' ============================================================================
Private Sub SaveState(ByRef st As AppState)
    st.Calc = Application.Calculation
    st.Scr = Application.ScreenUpdating
    st.Evt = Application.EnableEvents
    st.Bar = Application.StatusBar
End Sub

Private Sub RestoreState(ByRef st As AppState)
    On Error Resume Next
    Application.Calculation = st.Calc
    Application.EnableEvents = st.Evt
    Application.ScreenUpdating = st.Scr
    Application.StatusBar = st.Bar
End Sub

' ============================================================================
' ЗАЩИТА ЛИСТА
' ============================================================================
Private Function IsProtected(ws As Worksheet) As Boolean
    IsProtected = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
End Function

' Снимает защиту (сначала без пароля, затем паролем "sx"), запоминает параметры.
Private Function UnprotectSheet(ws As Worksheet, ByRef p As ProtInfo) As Boolean
    p.WasProtected = IsProtected(ws)
    p.Unlocked = False
    If Not p.WasProtected Then
        UnprotectSheet = True
        Exit Function
    End If

    p.Contents = ws.ProtectContents
    p.Draw = ws.ProtectDrawingObjects
    p.Scen = ws.ProtectScenarios
    With ws.Protection
        p.Allow(1) = .AllowFormattingCells
        p.Allow(2) = .AllowFormattingColumns
        p.Allow(3) = .AllowFormattingRows
        p.Allow(4) = .AllowInsertingColumns
        p.Allow(5) = .AllowInsertingRows
        p.Allow(6) = .AllowInsertingHyperlinks
        p.Allow(7) = .AllowDeletingColumns
        p.Allow(8) = .AllowDeletingRows
        p.Allow(9) = .AllowSorting
        p.Allow(10) = .AllowFiltering
        p.Allow(11) = .AllowUsingPivotTables
    End With

    On Error Resume Next
    ws.Unprotect ""
    If IsProtected(ws) Then
        Err.Clear
        ws.Unprotect SHEET_PWD
        p.Pwd = SHEET_PWD
    Else
        p.Pwd = ""
    End If
    On Error GoTo 0

    p.Unlocked = Not IsProtected(ws)
    UnprotectSheet = p.Unlocked
End Function

Private Sub ReprotectSheet(ws As Worksheet, ByRef p As ProtInfo)
    ws.Protect Password:=p.Pwd, DrawingObjects:=p.Draw, Contents:=p.Contents, Scenarios:=p.Scen, _
        AllowFormattingCells:=p.Allow(1), AllowFormattingColumns:=p.Allow(2), _
        AllowFormattingRows:=p.Allow(3), AllowInsertingColumns:=p.Allow(4), _
        AllowInsertingRows:=p.Allow(5), AllowInsertingHyperlinks:=p.Allow(6), _
        AllowDeletingColumns:=p.Allow(7), AllowDeletingRows:=p.Allow(8), _
        AllowSorting:=p.Allow(9), AllowFiltering:=p.Allow(10), AllowUsingPivotTables:=p.Allow(11)
    p.Unlocked = False
End Sub

' ============================================================================
' ПОДГОТОВКА МАССИВА К ЗАПИСИ (значения, а не формулы)
' Возвращает номер последней непустой строки массива (0 - пусто).
' ============================================================================
Private Function PrepareForWrite(ByRef arr As Variant, wsDst As Worksheet, _
                                 ByRef items As Long, ByRef errs As Long) As Long
    Dim isText(1 To 9) As Boolean
    Dim i As Long, c As Long, nr As Long, last As Long
    Dim v As Variant
    Dim s As String

    nr = UBound(arr, 1)
    For c = 1 To NCOL
        isText(c) = (wsDst.Cells(ROW_FIRST, COL_FIRST + c - 1).NumberFormat = "@")
    Next c

    items = 0
    errs = 0
    last = 0
    For i = 1 To nr
        For c = 1 To NCOL
            v = arr(i, c)
            If IsError(v) Then
                arr(i, c) = "#ОШИБКА"
                errs = errs + 1
                last = i
            ElseIf IsEmpty(v) Then
                ' пусто
            Else
                s = Norm(v)
                If Len(s) > 0 Then last = i
                If VarType(v) = vbString Then
                    If Len(v) = 0 Then
                        arr(i, c) = Empty           ' "" из формулы -> пустая ячейка
                    ElseIf isText(c) Then
                        ' в текстовой колонке Excel съедает ведущий апостроф: удваиваем
                        If Left$(CStr(v), 1) = "'" Then arr(i, c) = "'" & CStr(v)
                    Else
                        If NeedsApos(CStr(v)) Then arr(i, c) = "'" & CStr(v)
                    End If
                End If
            End If
        Next c
        If Len(Norm(arr(i, IX_POS))) > 0 Then
            If Left$(CStr(arr(i, IX_POS)), 1) <> "#" Then items = items + 1
        End If
    Next i
    PrepareForWrite = last
End Function

' Строка, которую Excel при записи через .Value превратит в число, дату или формулу.
Private Function NeedsApos(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Select Case Left$(s, 1)
        Case "=", "+", "-", "@", "'"
            NeedsApos = True
            Exit Function
    End Select
    If IsNumeric(s) Then
        NeedsApos = True
    ElseIf IsDate(s) Then
        NeedsApos = True
    End If
End Function

' Нормализация значения для сравнения: текст без лишних пробелов и переводов строк.
Private Function Norm(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then
        Norm = "#ERR"
        Exit Function
    End If
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbInteger, vbLong, vbDecimal, vbByte
            s = Trim$(Str$(v))
        Case vbDate
            s = Format$(v, "yyyy-mm-dd hh:nn:ss")
        Case Else
            s = CStr(v)
    End Select
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Norm = Trim$(s)
End Function

Private Function RowEmpty(a As Variant, ByVal i As Long) As Boolean
    Dim c As Long
    For c = 1 To NCOL
        If Len(Norm(a(i, c))) > 0 Then Exit Function
    Next c
    RowEmpty = True
End Function

Private Function LastFilled(a As Variant, ByVal n As Long) As Long
    Dim i As Long
    For i = n To 1 Step -1
        If Not RowEmpty(a, i) Then
            LastFilled = i
            Exit Function
        End If
    Next i
End Function

Private Function CountItems(a As Variant, ByVal n As Long) As Long
    Dim i As Long, k As Long
    For i = 1 To n
        If Len(Norm(a(i, IX_POS))) > 0 Then k = k + 1
    Next i
    CountItems = k
End Function

Private Function PosWord(ByVal n As Long) As String
    Dim m100 As Long, m10 As Long
    m100 = n Mod 100
    m10 = n Mod 10
    If m100 >= 11 And m100 <= 14 Then
        PosWord = "позиций"
    ElseIf m10 = 1 Then
        PosWord = "позиция"
    ElseIf m10 >= 2 And m10 <= 4 Then
        PosWord = "позиции"
    Else
        PosWord = "позиций"
    End If
End Function

Private Function StampText(ByVal nItems As Long, ByVal rev As Long, ByRef d As DiffResult, _
                           ByVal stamp As Date, ByVal accepted As Boolean) As String
    If accepted Then
        StampText = "Ревизия " & rev & " принята " & Format$(stamp, "dd.mm.yyyy hh:nn") & " | " & _
                    nItems & " " & PosWord(nItems) & " | непринятых отличий нет"
    Else
        StampText = "Обновлено " & Format$(stamp, "dd.mm.yyyy hh:nn") & " | " & _
                    nItems & " " & PosWord(nItems) & " | принятая ревизия: " & rev & _
                    " | не принято: изменено " & d.NChanged & ", добавлено " & d.NNew & _
                    ", удалено " & d.NDel
    End If
End Function

' ============================================================================
' ПОДСВЕТКА (только своя заливка в области B4:J160)
' ============================================================================
Private Function ClrChanged() As Long
    ClrChanged = RGB(255, 255, 0)
End Function

Private Function ClrNew() As Long
    ClrNew = RGB(146, 208, 80)
End Function

Private Sub ClearOurFills(rng As Range)
    Dim v As Variant
    Dim cl As Range
    Dim col As Long

    v = rng.Interior.ColorIndex
    If Not IsNull(v) Then
        If v = xlColorIndexNone Then Exit Sub       ' заливки нет вообще
    End If
    For Each cl In rng.Cells
        If cl.Interior.ColorIndex <> xlColorIndexNone Then
            col = cl.Interior.Color
            If col = ClrChanged() Or col = ClrNew() Then cl.Interior.ColorIndex = xlColorIndexNone
        End If
    Next cl
End Sub

Private Sub ApplyFills(wsDst As Worksheet, ByRef d As DiffResult, ByVal n As Long)
    Dim i As Long, c As Long, r As Long
    For i = 1 To n
        r = ROW_FIRST + i - 1
        Select Case d.RowState(i)
            Case 2
                wsDst.Range(wsDst.Cells(r, COL_FIRST), wsDst.Cells(r, COL_LAST)).Interior.Color = ClrNew()
            Case 1
                For c = 1 To NCOL
                    If d.CellMask(i, c) Then wsDst.Cells(r, COL_FIRST + c - 1).Interior.Color = ClrChanged()
                Next c
        End Select
    Next i
End Sub

' ============================================================================
' СРАВНЕНИЕ С ПРИНЯТЫМ СНИМКОМ
' ============================================================================
Private Function KeyOf(a As Variant, ByVal i As Long) As String
    KeyOf = LCase$(Norm(a(i, IX_ART))) & "|" & LCase$(Norm(a(i, IX_NAME))) & "|" & LCase$(Norm(a(i, IX_MODEL)))
End Function

Private Sub InitDiff(ByRef d As DiffResult, ByVal n As Long, ByVal m As Long)
    Dim nn As Long, cap As Long
    nn = n
    If nn < 1 Then nn = 1
    cap = NCOL * nn + m + 5
    d.NChanged = 0
    d.NCells = 0
    d.NNew = 0
    d.NDel = 0
    d.LogN = 0
    ReDim d.RowState(1 To nn)
    ReDim d.CellMask(1 To nn, 1 To NCOL)
    ReDim d.Entries(1 To cap, 1 To LOGCOLS)
End Sub

Private Sub AddEntry(ByRef d As DiffResult, ByVal stamp As Date, a As Variant, ByVal i As Long, _
                     ByVal fld As String, ByVal was As String, ByVal becomes As String, ByVal typ As String)
    d.LogN = d.LogN + 1
    d.Entries(d.LogN, 1) = stamp
    d.Entries(d.LogN, 2) = PENDING
    d.Entries(d.LogN, 3) = Norm(a(i, IX_POS))
    d.Entries(d.LogN, 4) = Norm(a(i, IX_ART))
    d.Entries(d.LogN, 5) = Norm(a(i, IX_NAME))
    d.Entries(d.LogN, 6) = fld
    d.Entries(d.LogN, 7) = was
    d.Entries(d.LogN, 8) = becomes
    d.Entries(d.LogN, 9) = typ
End Sub

Private Function QtyText(a As Variant, ByVal i As Long) As String
    QtyText = "кол-во " & Norm(a(i, IX_QTY))
    If Len(Norm(a(i, IX_UNIT))) > 0 Then QtyText = QtyText & " " & Norm(a(i, IX_UNIT))
End Function

Private Function Shown(ByVal s As String) As String
    If Len(s) = 0 Then
        Shown = "(пусто)"
    Else
        Shown = s
    End If
End Function

' cur, old - массивы B:J (строки 1..157); n, m - число заполненных строк.
Private Sub BuildDiff(cur As Variant, ByVal n As Long, old As Variant, ByVal m As Long, _
                      hdr As Variant, ByVal stamp As Date, ByRef d As DiffResult)
    Dim i As Long, j As Long, c As Long, nn As Long, mm As Long
    Dim k As String, kk As String, art As String
    Dim dOld As Object, cntOld As Object, cntCur As Object
    Dim aOld As Object, aOldIdx As Object, aCur As Object
    Dim curMatch() As Long
    Dim oldUsed() As Boolean
    Dim rowChanged As Boolean

    InitDiff d, n, m
    nn = n
    If nn < 1 Then nn = 1
    mm = m
    If mm < 1 Then mm = 1
    ReDim curMatch(1 To nn)
    ReDim oldUsed(1 To mm)

    Set dOld = CreateObject("Scripting.Dictionary")
    Set cntOld = CreateObject("Scripting.Dictionary")
    Set cntCur = CreateObject("Scripting.Dictionary")

    ' --- проход 1: полный ключ (артикул | наименование | тип), дубли - по порядку ---
    For j = 1 To m
        If Not RowEmpty(old, j) Then
            k = KeyOf(old, j)
            If cntOld.Exists(k) Then
                cntOld(k) = cntOld(k) + 1
            Else
                cntOld.Add k, 1
            End If
            dOld.Add k & "#" & cntOld(k), j
        End If
    Next j
    For i = 1 To n
        If Not RowEmpty(cur, i) Then
            k = KeyOf(cur, i)
            If cntCur.Exists(k) Then
                cntCur(k) = cntCur(k) + 1
            Else
                cntCur.Add k, 1
            End If
            kk = k & "#" & cntCur(k)
            If dOld.Exists(kk) Then
                curMatch(i) = dOld(kk)
                oldUsed(curMatch(i)) = True
            End If
        End If
    Next i

    ' --- проход 2: по артикулу, если он один среди несопоставленных с обеих сторон ---
    Set aOld = CreateObject("Scripting.Dictionary")
    Set aOldIdx = CreateObject("Scripting.Dictionary")
    Set aCur = CreateObject("Scripting.Dictionary")
    For j = 1 To m
        If Not RowEmpty(old, j) Then
            If Not oldUsed(j) Then
                art = LCase$(Norm(old(j, IX_ART)))
                If Len(art) > 0 Then
                    If aOld.Exists(art) Then
                        aOld(art) = aOld(art) + 1
                    Else
                        aOld.Add art, 1
                    End If
                    aOldIdx(art) = j
                End If
            End If
        End If
    Next j
    For i = 1 To n
        If Not RowEmpty(cur, i) Then
            If curMatch(i) = 0 Then
                art = LCase$(Norm(cur(i, IX_ART)))
                If Len(art) > 0 Then
                    If aCur.Exists(art) Then
                        aCur(art) = aCur(art) + 1
                    Else
                        aCur.Add art, 1
                    End If
                End If
            End If
        End If
    Next i
    For i = 1 To n
        If Not RowEmpty(cur, i) Then
            If curMatch(i) = 0 Then
                art = LCase$(Norm(cur(i, IX_ART)))
                If Len(art) > 0 Then
                    If aOld.Exists(art) Then
                        If aCur(art) = 1 And aOld(art) = 1 Then
                            j = aOldIdx(art)
                            curMatch(i) = j
                            oldUsed(j) = True
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' --- классификация ---
    For i = 1 To n
        If Not RowEmpty(cur, i) Then
            If curMatch(i) = 0 Then
                d.RowState(i) = 2
                d.NNew = d.NNew + 1
                AddEntry d, stamp, cur, i, FLD_ALL, "", QtyText(cur, i), TYP_NEW
            Else
                j = curMatch(i)
                rowChanged = False
                For c = 2 To NCOL                       ' колонку "Поз" (1) не сравниваем
                    If StrComp(Norm(cur(i, c)), Norm(old(j, c)), vbBinaryCompare) <> 0 Then
                        d.CellMask(i, c) = True
                        d.NCells = d.NCells + 1
                        rowChanged = True
                        AddEntry d, stamp, cur, i, Norm(hdr(1, c)), Shown(Norm(old(j, c))), _
                                 Shown(Norm(cur(i, c))), TYP_CHG
                    End If
                Next c
                If rowChanged Then
                    d.RowState(i) = 1
                    d.NChanged = d.NChanged + 1
                End If
            End If
        End If
    Next i
    For j = 1 To m
        If Not RowEmpty(old, j) Then
            If Not oldUsed(j) Then
                d.NDel = d.NDel + 1
                AddEntry d, stamp, old, j, FLD_ALL, QtyText(old, j), "", TYP_DEL
            End If
        End If
    Next j
End Sub

' ============================================================================
' ЛИСТ-СНИМОК "Спец_принятая" (скрытый)
' ============================================================================
Private Function SnapValid(ws As Worksheet) As Boolean
    If IsError(ws.Range("A1").Value) Then Exit Function
    SnapValid = (CStr(ws.Range("A1").Value) = SNAP_MARK)
End Function

Private Function CreateSnap(wb As Workbook, wsAfter As Worksheet) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wsAfter)
    ws.Name = SH_SNAP
    ws.Range(ws.Cells(ROW_FIRST, COL_FIRST), ws.Cells(ROW_LAST, COL_LAST)).NumberFormat = "@"
    ws.Range("A1").Value = SNAP_MARK
    ws.Range("A2").Value = "Служебный лист SX_Spec: последняя ПРИНЯТАЯ версия листа «" & SH_DST & _
                           "» (B4:J160, значения текстом). B1 - номер ревизии, C1 - дата принятия. " & _
                           "Не редактировать и не удалять."
    ws.Range("C1").NumberFormat = "dd.mm.yyyy hh:mm"
    ws.Visible = xlSheetHidden
    Set CreateSnap = ws
End Function

' Записывает текущий итог как принятый снимок (значения как текст) и номер ревизии.
Private Sub SaveSnap(ws As Worksheet, cur As Variant, ByVal rev As Long, ByVal stamp As Date)
    Dim out() As Variant
    Dim i As Long, c As Long, nr As Long
    Dim s As String
    Dim rng As Range

    nr = UBound(cur, 1)
    ReDim out(1 To nr, 1 To NCOL)
    For i = 1 To nr
        For c = 1 To NCOL
            s = Norm(cur(i, c))
            If Len(s) > 0 Then out(i, c) = s
        Next c
    Next i
    Set rng = ws.Range(ws.Cells(ROW_FIRST, COL_FIRST), ws.Cells(ROW_FIRST + nr - 1, COL_LAST))
    rng.ClearContents
    rng.NumberFormat = "@"
    rng.Value = out
    ws.Range("A1").Value = SNAP_MARK
    ws.Range("B1").Value = rev
    ws.Range("C1").Value = stamp
End Sub

' ============================================================================
' ЖУРНАЛ "Изменения"
' ============================================================================
Private Function EnsureLog(wb As Workbook, wsAfter As Worksheet) As Worksheet
    Dim ws As Worksheet
    Dim h As Variant
    Dim c As Long

    Set ws = SheetOf(wb, SH_LOG)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wsAfter)
        ws.Name = SH_LOG
        ws.Columns("C:H").NumberFormat = "@"
        ws.Columns("A").NumberFormat = "dd.mm.yyyy hh:mm"
        ws.Range("A1").Value = "Журнал изменений спецификации"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 12
        h = Array("Дата-время", "Ревизия", "Поз", "Артикул", "Наименование", "Поле", "Было", "Стало", "Тип")
        For c = 0 To LOGCOLS - 1
            ws.Cells(LOG_HDR_ROW, c + 1).Value = h(c)
        Next c
        With ws.Range(ws.Cells(LOG_HDR_ROW, 1), ws.Cells(LOG_HDR_ROW, LOGCOLS))
            .Font.Bold = True
            .Interior.Color = RGB(221, 235, 247)
        End With
        ws.Columns(1).ColumnWidth = 17
        ws.Columns(2).ColumnWidth = 12
        ws.Columns(3).ColumnWidth = 8
        ws.Columns(4).ColumnWidth = 16
        ws.Columns(5).ColumnWidth = 60
        ws.Columns(6).ColumnWidth = 24
        ws.Columns(7).ColumnWidth = 30
        ws.Columns(8).ColumnWidth = 30
        ws.Columns(9).ColumnWidth = 12
        ws.Columns(5).WrapText = True
        ws.Columns(7).WrapText = True
        ws.Columns(8).WrapText = True
        ws.Tab.Color = RGB(255, 192, 0)
    End If
    Set EnsureLog = ws
End Function

Private Function LastLogRow(ws As Worksheet) As Long
    Dim a As Long, b As Long
    a = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    b = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If b > a Then a = b
    If a < LOG_FIRST - 1 Then a = LOG_FIRST - 1
    LastLogRow = a
End Function

' Первая строка хвостового блока "не принято" (или строка после последней записи).
Private Function PendingStart(ws As Worksheet) As Long
    Dim r As Long
    r = LastLogRow(ws)
    Do While r >= LOG_FIRST
        If CStr(ws.Cells(r, 2).Value) <> PENDING Then Exit Do
        r = r - 1
    Loop
    PendingStart = r + 1
End Function

' Пересобирает непринятый набор: старые строки "не принято" стираются, пишется текущий набор.
' Принятые ревизии выше не затрагиваются.
Private Sub WriteLog(ws As Worksheet, ByRef d As DiffResult)
    Dim r0 As Long, lastR As Long, k As Long, c As Long
    Dim out() As Variant

    r0 = PendingStart(ws)
    lastR = LastLogRow(ws)
    If lastR >= r0 Then ws.Range(ws.Cells(r0, 1), ws.Cells(lastR, LOGCOLS)).ClearContents
    If d.LogN > 0 Then
        ReDim out(1 To d.LogN, 1 To LOGCOLS)
        For k = 1 To d.LogN
            For c = 1 To LOGCOLS
                out(k, c) = d.Entries(k, c)
            Next c
        Next k
        ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + d.LogN - 1, LOGCOLS)).Value = out
    End If
End Sub

' "Принять": записи "не принято" получают номер ревизии и дату принятия.
Private Sub FreezeLog(ws As Worksheet, ByVal rev As Long, ByVal stamp As Date)
    Dim p As Long, lastR As Long
    p = PendingStart(ws)
    lastR = LastLogRow(ws)
    If lastR < p Then Exit Sub
    ws.Range(ws.Cells(p, 1), ws.Cells(lastR, 1)).Value = stamp
    ws.Range(ws.Cells(p, 2), ws.Cells(lastR, 2)).Value = rev
End Sub

Private Sub SetLogStatus(ws As Worksheet, ByVal rev As Long, ByRef d As DiffResult, ByVal stamp As Date)
    ws.Range("A2").Value = "Принятая ревизия: " & rev & ". Непринятые отличия (строки «" & PENDING & _
        "»): изменено позиций " & d.NChanged & ", добавлено " & d.NNew & ", удалено " & d.NDel & _
        ". Обновлено " & Format$(stamp, "dd.mm.yyyy hh:nn") & ". Не сортируйте журнал."
End Sub
