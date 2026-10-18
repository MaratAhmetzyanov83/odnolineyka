Attribute VB_Name = "SX_KJ"
' Кабельный журнал в DWG: лист «Кабельный_журнал» (P:AB) -> таблица AutoCAD по образцу КЖ_шаблон.dwg.
' DWG открывать не нужно: запись в фоне (ObjectDBX). Если DWG открыт в AutoCAD — правится и сохраняется открытый.
' В DWG одна таблица с разрывами: AutoCAD сам режет её на листы, программа добавляет или удаляет вкладки листов.
' Графы 9–13 (проложено, испытания, дата) заполняют в DWG вручную — при обновлении они сохраняются по маркировке.
' После записи — отметка <книга>.kj.txt для программы «Объект» (КЖ актуален / устарел).
Option Explicit

Private Const KJ_SHEET As String = "Кабельный_журнал"
Private Const KJ_NAME As String = "SX_KJ_DWG"
Private Const KJ_HEAD As String = "Маркировка кабеля"
Private Const KJ_STAMP As String = "SX_KJ_STAMP"   ' файл штампа (внешняя ссылка), относительно папки книги
Private Const STAMP_XREF As String = "Штамп"       ' имя внешней ссылки в блоке рамки
Private Const NC As Long = 13
Private Const MAN0 As Long = 8                  ' первая ручная графа (0 = маркировка)
Private Const FONT_PFX As String = "{\fGOST 2.304 type A|b0|i0|c0;"

Private kjErr As Long                           ' ячеек Excel с ошибкой (#Н/Д ...)
Private kjNote As String                        ' замечания для итогового сообщения
Private kjStamp As String                       ' полный путь к DWG штампа ("" — не менять)
Private kjDrop As Collection                    ' лишние вкладки: в фоне (ObjectDBX) AutoCAD их не удаляет

' ---------------------------------------------------------------- кнопки
Public Sub SX_KJMake()
    KJ_Run "", False
End Sub

Public Sub SX_KJFile()
    KJ_Run "", True
End Sub

' Выбрать файл штампа (DWG, на него ссылается рамка КЖ) и сразу обновить КЖ.
Public Sub SX_KJStamp()
    KJ_Run "", False, True
End Sub

' Для проверки программой: путь DWG (и штампа) задан, окон нет (SX_SetQuiet). Возвращает итоговое сообщение.
Public Function SX_KJTo(dwg As Variant, Optional stamp As Variant) As String
    If IsMissing(stamp) Then KJ_Run CStr(dwg), False Else KJ_Run CStr(dwg), False, False, CStr(stamp)
    SX_KJTo = SX_LastMsg
End Function

Private Sub KSay(txt As String, Optional st As VbMsgBoxStyle = vbInformation)
    SX_LastMsg = txt
    If SX_Quiet Then
        Debug.Print txt
    Else
        MsgBox txt, st, "Кабельный журнал"
    End If
End Sub

Private Function Fso() As Object
    Static f As Object
    If f Is Nothing Then Set f = CreateObject("Scripting.FileSystemObject")
    Set Fso = f
End Function

' ---------------------------------------------------------------- главная
Private Sub KJ_Run(forced As String, ask As Boolean, Optional pickStamp As Boolean, Optional stamp As String)
    Dim wb As Workbook, ws As Worksheet, rows As Collection, dwg As String, tpl As String, msg As String
    Set wb = ActiveWorkbook
    If wb Is Nothing Then KSay "Откройте таблицу щита.", vbExclamation: Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets(KJ_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then KSay "В книге нет листа «" & KJ_SHEET & "».", vbExclamation: Exit Sub

    kjErr = 0: kjNote = ""
    Set rows = ReadRows(ws)
    If rows.Count = 0 Then KSay "На листе «" & KJ_SHEET & "» нет кабелей (столбец P пуст).", vbExclamation: Exit Sub

    If forced <> "" Then dwg = forced Else dwg = DwgPath(wb, ask)
    If dwg = "" Then Exit Sub
    If stamp <> "" Then
        kjStamp = stamp
    Else
        kjStamp = StampPath(wb, pickStamp Or (Not Fso.FileExists(dwg) And Not SX_Quiet))
        If pickStamp And kjStamp = "" Then Exit Sub
    End If
    tpl = Fso.GetAbsolutePathName(ThisWorkbook.Path & "\..\2_AutoCAD\Кабельный_журнал\КЖ_шаблон.dwg")
    If Not Fso.FileExists(dwg) And Not Fso.FileExists(tpl) Then
        KSay "Нет образца КЖ:" & vbLf & tpl, vbExclamation: Exit Sub
    End If

    msg = WriteDwg(dwg, rows, tpl)
    Application.StatusBar = False
    If Left(msg, 1) = "!" Then
        KSay Mid(msg, 2), vbExclamation
    Else
        KJ_Mark wb, dwg, rows
        If kjErr > 0 Then kjNote = kjNote & vbLf & "Ячеек Excel с ошибкой (перенесены как есть): " & kjErr & "."
        KSay msg & kjNote
    End If
End Sub

' Отметка для программы «Объект» (контроль актуальности КЖ): <книга>.kj.txt рядом с книгой, UTF-8.
' Строки «ключ<Tab>значение»: DWG, AT (время записи), BOOK (дата файла книги), SAVED (1/0),
' ROW — графы 1–8 (по проекту) в том виде, в каком ушли в DWG. Программа сравнит их с листом по маркировке.
Private Sub KJ_Mark(wb As Workbook, dwg As String, rows As Collection)
    Dim st As Object, a As Variant, i As Long, s As String
    On Error GoTo fail
    If wb.Path = "" Then Exit Sub
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText ";; odn kj v1 - KJ stamp for program Objekt (6_Programma)" & vbCrLf
    st.WriteText "DWG" & vbTab & dwg & vbCrLf
    st.WriteText "AT" & vbTab & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    st.WriteText "BOOK" & vbTab & Format(FileDateTime(wb.FullName), "yyyy-mm-dd hh:nn:ss") & vbCrLf
    st.WriteText "SAVED" & vbTab & IIf(wb.Saved, "1", "0") & vbCrLf
    For Each a In rows
        s = "ROW"
        For i = 0 To MAN0 - 1
            s = s & vbTab & Replace(Replace(Replace(a(i), vbTab, " "), vbCr, " "), vbLf, " ")
        Next i
        st.WriteText s & vbCrLf
    Next a
    st.SaveToFile Fso.BuildPath(wb.Path, Fso.GetBaseName(wb.Name) & ".kj.txt"), 2
    st.Close
    Exit Sub
fail:
    kjNote = kjNote & vbLf & "Отметка для программы «Объект» не записана: " & Err.Description
End Sub

' Строки P:AB с 5-й: маркировка не пустая. Значение — как видно в ячейке.
Private Function ReadRows(ws As Worksheet) As Collection
    Dim res As New Collection, last As Long, v As Variant, r As Long, c As Long, a() As String
    Dim seen As Object, mk As String
    Set seen = CreateObject("Scripting.Dictionary")
    last = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
    If last < 5 Then Set ReadRows = res: Exit Function
    v = ws.Range("P5:AB" & last).Value
    For r = 1 To UBound(v, 1)
        If Not IsError(v(r, 1)) Then
            mk = Trim(CStr(v(r, 1)))
            If mk <> "" And Not (IsNumeric(v(r, 1)) And mk = "0") Then
                ReDim a(0 To NC - 1)
                For c = 1 To NC
                    a(c - 1) = CellText(v(r, c), ws.Cells(r + 4, 15 + c))
                Next c
                a(0) = mk
                If seen.Exists(mk) Then
                    kjNote = kjNote & vbLf & "Маркировка «" & mk & "» повторяется (строка " & (r + 4) & ")."
                Else
                    seen.Add mk, True
                End If
                res.Add a
            End If
        End If
    Next r
    Set ReadRows = res
End Function

Private Function CellText(v As Variant, c As Range) As String
    If IsError(v) Then
        kjErr = kjErr + 1
        CellText = c.Text
    ElseIf IsEmpty(v) Then
        CellText = ""
    ElseIf VarType(v) = vbString Then
        CellText = Trim(v)
    Else
        CellText = Trim(c.Text)
        If InStr(CellText, "#") > 0 Then CellText = CStr(v)   ' узкий столбец: ####
    End If
End Function

' Путь DWG: из имени книги SX_KJ_DWG (относительно папки книги) или спросить.
Private Function DwgPath(wb As Workbook, ask As Boolean) As String
    Dim p As String, nm As Name, f As Variant, rel As String, def As String
    On Error Resume Next
    Set nm = wb.Names(KJ_NAME)
    On Error GoTo 0
    If Not nm Is Nothing Then
        p = Replace(Mid(nm.RefersTo, 3, Len(nm.RefersTo) - 3), """""", """")
        If InStr(p, ":") = 0 And Left(p, 2) <> "\\" And wb.Path <> "" Then p = wb.Path & "\" & p
    End If
    If Not ask And p <> "" Then
        If Fso.FolderExists(Fso.GetParentFolderName(p)) Then DwgPath = p: Exit Function
    End If
    If p <> "" Then
        def = p
    ElseIf wb.Path <> "" Then
        def = wb.Path & "\КЖ_" & ShieldName(wb) & ".dwg"
    Else
        def = "КЖ_" & ShieldName(wb) & ".dwg"
    End If
    f = Application.GetSaveAsFilename(def, "Чертёж AutoCAD (*.dwg),*.dwg", 1, _
        "Кабельный журнал: новый DWG или уже готовый (он обновится)")
    If VarType(f) = vbBoolean Then Exit Function
    p = CStr(f)
    rel = p
    If wb.Path <> "" Then
        If LCase(Left(p, Len(wb.Path) + 1)) = LCase(wb.Path & "\") Then rel = Mid(p, Len(wb.Path) + 2)
    End If
    wb.Names.Add Name:=KJ_NAME, RefersTo:="=""" & Replace(rel, """", """""") & """", Visible:=False
    DwgPath = p
End Function

' Файл штампа: из имени книги SX_KJ_STAMP или выбрать (ask). "" — не задан (ссылку в DWG не трогаем).
Private Function StampPath(wb As Workbook, ask As Boolean) As String
    Dim p As String, nm As Name, f As Variant
    On Error Resume Next
    Set nm = wb.Names(KJ_STAMP)
    On Error GoTo 0
    If Not nm Is Nothing Then
        p = Replace(Mid(nm.RefersTo, 3, Len(nm.RefersTo) - 3), """""", """")
        If InStr(p, ":") = 0 And Left(p, 2) <> "\\" And wb.Path <> "" Then p = Fso.GetAbsolutePathName(wb.Path & "\" & p)
    End If
    If Not ask Then StampPath = p: Exit Function
    If p <> "" Then
        On Error Resume Next
        ChDrive Left(p, 1): ChDir Fso.GetParentFolderName(p)
        On Error GoTo 0
    End If
    f = Application.GetOpenFilename("Чертёж AutoCAD (*.dwg),*.dwg", 1, _
        "Штамп для кабельного журнала: DWG, на который ссылается рамка (Отмена — оставить как есть)")
    If VarType(f) = vbBoolean Then StampPath = p: Exit Function
    p = CStr(f)
    wb.Names.Add Name:=KJ_STAMP, RefersTo:="=""" & Replace(RelPath(wb.Path, p), """", """""") & """", Visible:=False
    StampPath = p
End Function

' Внешняя ссылка «Штамп»: путь — относительно папки DWG (на другом диске — полный). Возвращает замечание.
Private Function SetStamp(db As Object, dwg As String) As String
    Dim b As Object, x As Object, cur As String, full As String, folder As String
    For Each b In db.Blocks
        If b.IsXRef Then
            If LCase(b.Name) = LCase(STAMP_XREF) Then Set x = b: Exit For
            If x Is Nothing And InStr(LCase(b.Name), "штамп") > 0 Then Set x = b
        End If
    Next b
    If x Is Nothing Then
        If kjStamp <> "" Then SetStamp = vbLf & "В DWG нет внешней ссылки «" & STAMP_XREF & "» — штамп не подключён."
        Exit Function
    End If
    folder = Fso.GetParentFolderName(dwg)
    If kjStamp <> "" Then x.Path = RelPath(folder, kjStamp)
    cur = x.Path
    If InStr(cur, ":") > 0 Or Left(cur, 2) = "\\" Then full = cur Else full = Fso.GetAbsolutePathName(folder & "\" & cur)
    If Fso.FileExists(full) Then
        SetStamp = vbLf & "Штамп: " & cur
    Else
        SetStamp = vbLf & "Штамп не найден: " & cur & vbLf & "Укажите файл: кнопка «Штамп КЖ»."
    End If
End Function

' Относительный путь от папки до файла; другой диск — полный путь.
Private Function RelPath(ByVal fromDir As String, ByVal target As String) As String
    Dim a() As String, b() As String, i As Long, n As Long, r As String
    If fromDir = "" Then RelPath = target: Exit Function
    a = Split(Fso.GetAbsolutePathName(fromDir), "\")
    b = Split(Fso.GetAbsolutePathName(target), "\")
    If LCase(a(0)) <> LCase(b(0)) Then RelPath = target: Exit Function
    Do While n <= UBound(a) And n < UBound(b)
        If LCase(a(n)) <> LCase(b(n)) Then Exit Do
        n = n + 1
    Loop
    For i = n To UBound(a)
        If a(i) <> "" Then r = r & "..\"
    Next i
    For i = n To UBound(b)
        r = r & b(i)
        If i < UBound(b) Then r = r & "\"
    Next i
    If Left(r, 3) <> "..\" Then r = ".\" & r
    RelPath = r
End Function

Private Function ShieldName(wb As Workbook) As String
    Dim s As String
    s = wb.Name
    If InStrRev(s, ".") > 0 Then s = Left(s, InStrRev(s, ".") - 1)
    If LCase(Left(s, 11)) = "однолинейка" Then s = Mid(s, 12)
    Do While Len(s) > 0 And InStr(" _-", Left(s, 1)) > 0
        s = Mid(s, 2)
    Loop
    If s = "" Then s = "щит"
    ShieldName = s
End Function

' ---------------------------------------------------------------- AutoCAD
Private Function AcadApp(started As Boolean) As Object
    Dim app As Object, t As Single, v As Variant
    On Error Resume Next
    Set app = GetObject(, "AutoCAD.Application")
    If app Is Nothing Then
        Application.StatusBar = "Кабельный журнал: запускаю AutoCAD в фоне..."
        Set app = CreateObject("AutoCAD.Application")
        started = Not app Is Nothing
    End If
    If app Is Nothing Then Exit Function
    t = Timer
    Do
        Err.Clear
        v = app.Version
        If Err.Number = 0 Then Exit Do
        DoEvents
    Loop While Timer - t < 120
    On Error GoTo 0
    Set AcadApp = app
End Function

' Возвращает итоговое сообщение; ошибка — с "!" в начале. Старый DWG -> *.bak рядом.
Private Function WriteDwg(dwg As String, rows As Collection, tpl As String) As String
    Dim app As Object, started As Boolean, db As Object, doc As Object, i As Long
    Dim src As String, tmp As String, info As String, stage As String, bak As String
    Application.StatusBar = "Кабельный журнал: подключаюсь к AutoCAD..."
    Set app = AcadApp(started)
    If app Is Nothing Then WriteDwg = "!AutoCAD не найден. Нужен установленный AutoCAD.": Exit Function

    On Error GoTo fail
    stage = "открыть DWG"
    For i = 0 To app.Documents.Count - 1
        If LCase(app.Documents.Item(i).FullName) = LCase(dwg) Then Set doc = app.Documents.Item(i)
    Next i
    If doc Is Nothing Then
        If Fso.FileExists(dwg) Then src = dwg Else src = tpl
        Set db = app.GetInterfaceObject("ObjectDBX.AxDbDocument." & Split(app.Version, ".")(0))
        db.Open src
    Else
        If doc.ReadOnly Then Err.Raise vbObjectError + 515, , "DWG открыт в AutoCAD только для чтения. Закройте его и нажмите ещё раз."
        Set db = doc
    End If

    stage = "заполнить таблицу"
    Set kjDrop = New Collection
    Application.StatusBar = "Кабельный журнал: заполняю " & rows.Count & " строк..."
    info = BuildKJ(db, rows, tpl, app, src)

    stage = "подключить штамп"
    info = info & SetStamp(db, dwg)
    stage = "сохранить DWG"
    Application.StatusBar = "Кабельный журнал: сохраняю..."
    If Not doc Is Nothing Then
        If kjDrop.Count > 0 Then doc.Activate: DropLayouts doc
        doc.Save
    Else
        tmp = Fso.GetSpecialFolder(2) & "\sx_kj_" & Format(Now, "yymmdd_hhnnss") & ".dwg"
        db.SaveAs tmp
        Set db = Nothing
        bak = Left(dwg, Len(dwg) - 4) & ".bak"
        If Fso.FileExists(dwg) Then
            stage = "заменить DWG (файл занят? закройте его у себя и у коллег)"
            Fso.CopyFile dwg, tmp & ".old", True
        End If
        Fso.CopyFile tmp, dwg, True
        Fso.DeleteFile tmp
        If kjDrop.Count > 0 Then
            stage = "удалить лишние листы (AutoCAD откроет файл на секунду)"
            Set doc = app.Documents.Open(dwg)
            DropLayouts doc
            doc.Close True
        End If
        If Fso.FileExists(tmp & ".old") Then                ' .bak — прежняя версия (не промежуточная от AutoCAD)
            Fso.CopyFile tmp & ".old", bak, True
            Fso.DeleteFile tmp & ".old"
        End If
        tmp = ""
    End If
    If started Then app.Quit
    WriteDwg = "Кабельный журнал обновлён." & vbLf & dwg & vbLf & info
    Exit Function
fail:
    WriteDwg = "!Не удалось " & stage & ":" & vbLf & Err.Description & " (стр. " & Erl & ")" & vbLf & dwg
    Set db = Nothing
    On Error Resume Next
    If tmp <> "" Then
        If Fso.FileExists(tmp) Then Fso.DeleteFile tmp
        If Fso.FileExists(tmp & ".old") Then Fso.CopyFile tmp & ".old", bak, True: Fso.DeleteFile tmp & ".old"
    End If
    If started Then app.Quit
End Function

' ---------------------------------------------------------------- построение
Private Function BuildKJ(db As Object, rows As Collection, tpl As String, app As Object, src As String) As String
    Dim ms As Object, e As Object, bodies As New Collection, p As Object, t As Object
    Dim man As Object, lost As String, nLost As Long, kept As Long
    Dim i As Long, j As Long, k As Long, a As Variant, b As Variant, mk As String, raw(0 To 4) As String, has As Boolean
    Dim ip As Variant, sty As String, lay As String, w(0 To NC - 1) As Double, al(0 To NC - 1) As Long
    Dim tc As Object, lw As Long, cc(0 To NC - 1) As Object, th(0 To NC - 1) As Double, hm As Double, vm As Double, rowH As Double, txt As String
    Dim brk As Boolean, sp As Double, bh0 As Double, bh1 As Double, rTop As Boolean, manH As Boolean

    On Error GoTo bad
    Set ms = db.ModelSpace
    For Each e In ms
        If e.ObjectName = "AcDbTable" Then
            If e.Columns = NC Then
                If Clean(e.GetText(0, 0)) <> KJ_HEAD Then bodies.Add e
            End If
        End If
    Next e
    If bodies.Count = 0 Then Err.Raise vbObjectError + 513, , _
        "В DWG нет таблицы кабельного журнала (13 граф, без шапки). Образец — КЖ_шаблон.dwg."
    Set p = bodies(1)
    For Each e In bodies
        ip = e.InsertionPoint: a = p.InsertionPoint
        If ip(0) < a(0) - 0.01 Then Set p = e
    Next e

    ' ручные графы из DWG
    Set man = CreateObject("Scripting.Dictionary")
    For Each e In bodies
        For i = 0 To e.Rows - 1
            mk = Clean(e.GetText(i, 0))
            If mk <> "" And Not man.Exists(mk) Then
                has = False
                For j = 0 To 4
                    raw(j) = e.GetText(i, MAN0 + j)
                    If Clean(raw(j)) <> "" Then has = True
                Next j
                If has Then man.Add mk, raw
            End If
        Next i
    Next e

    ' вид таблицы — с образца
    ip = p.InsertionPoint: sty = p.StyleName: lay = p.Layer
    Set tc = p.TrueColor: lw = p.Lineweight                 ' в фоне новая таблица иначе пурпурная
    hm = p.HorzCellMargin: vm = p.VertCellMargin
    rowH = p.GetRowHeight(0)
    For i = 1 To p.Rows - 1
        If p.GetRowHeight(i) < rowH Then rowH = p.GetRowHeight(i)
    Next i
    For j = 0 To NC - 1
        w(j) = p.GetColumnWidth(j): al(j) = p.GetCellAlignment(0, j): th(j) = p.GetCellTextHeight(0, j)
        Set cc(j) = p.GetCellContentColor(0, j)            ' цвет текста (в образце — ПоБлоку)
    Next j
    brk = p.BreaksEnabled
    If brk Then
        sp = p.BreakSpacing: rTop = p.RepeatTopLabels: manH = p.AllowManualHeights
        bh0 = p.GetBreakHeight(0)
        On Error Resume Next
        bh1 = bh0: bh1 = p.GetBreakHeight(1)
        On Error GoTo bad
    Else
        sp = 25: rTop = True: manH = True: bh0 = 190.83: bh1 = 231.31   ' как в КЖ_шаблон.dwg
    End If
    For Each e In bodies
        e.Delete
    Next e
    DropDataLinks db

    ' новая таблица: проектные графы из Excel, ручные — из Excel, если там заполнены, иначе из DWG
    Set t = ms.AddTable(ip, rows.Count, NC, rowH, 30#)
    t.StyleName = sty
    t.Layer = lay
    t.TrueColor = tc
    t.Lineweight = lw
    t.TitleSuppressed = True
    t.HeaderSuppressed = True
    On Error Resume Next
    t.UnmergeCells 0, 0, 0, NC - 1
    On Error GoTo bad
    For j = 0 To NC - 1
        t.SetColumnWidth j, w(j)
    Next j
    t.HorzCellMargin = hm: t.VertCellMargin = vm
    t.RegenerateTableSuppressed = True
    For i = 0 To rows.Count - 1
        a = rows(i + 1)
        has = False
        For j = MAN0 To NC - 1
            If a(j) <> "" Then has = True
        Next j
        For j = 0 To NC - 1
            t.SetCellTextHeight i, j, th(j)
            t.SetCellAlignment i, j, al(j)
            t.SetCellContentColor i, j, cc(j)
            If j >= MAN0 And Not has And man.Exists(a(0)) Then
                b = man(a(0))
                txt = b(j - MAN0)
            Else
                txt = Mt(CStr(a(j)))
            End If
            If Len(txt) > 0 Then t.SetText i, j, txt          ' пустую строку AutoCAD не принимает
        Next j
        If Not has And man.Exists(a(0)) Then kept = kept + 1
        t.SetRowHeight i, rowH
        If i Mod 20 = 0 Then Application.StatusBar = "Кабельный журнал: строка " & (i + 1) & " из " & rows.Count
    Next i
    t.RegenerateTableSuppressed = False

    ' кабели с ручными данными, которых больше нет в Excel
    For Each a In man.Keys
        has = False
        For i = 1 To rows.Count
            b = rows(i)
            If b(0) = a Then has = True: Exit For
        Next i
        If Not has Then
            nLost = nLost + 1
            If nLost <= 10 Then lost = lost & " " & a
        End If
    Next a

    t.BreaksEnabled = True
    t.BreakSpacing = sp
    t.RepeatTopLabels = rTop
    t.AllowManualHeights = manH
    t.SetBreakHeight 0, bh0
    t.SetBreakHeight 1, bh1

    ' листов: строки не делятся, первый фрагмент bh0, остальные bh1
    Dim pages As Long, used As Double, cap As Double, h As Double
    pages = 1: cap = bh0
    For i = 0 To t.Rows - 1
        h = t.GetRowHeight(i)
        If used > 0 And used + h > cap + 0.01 Then pages = pages + 1: used = 0: cap = bh1
        used = used + h
    Next i
    k = SetPages(db, pages, t.Width + sp, tpl, app, src)

    BuildKJ = "Кабелей: " & rows.Count & ", листов: " & pages & "."
    If kept > 0 Then BuildKJ = BuildKJ & vbLf & "Ручные графы из DWG сохранены: " & kept & "."
    If k > 0 Then BuildKJ = BuildKJ & vbLf & "Добавлено листов: " & k & "."
    If k < 0 Then BuildKJ = BuildKJ & vbLf & "Удалено листов: " & -k & "."
    If nLost > 0 Then BuildKJ = BuildKJ & vbLf & "Нет в Excel, ручные данные остались в .bak:" & lost & IIf(nLost > 10, " ...", "")
    Exit Function
bad:
    Err.Raise Err.Number, , Err.Description & " [стр. " & Erl & ", " & i & ":" & j & "]"
End Function

' Вкладки листов: столько, сколько фрагментов таблицы. Новые — копия 2-го листа (штамп ф. 2а),
' видовой экран сдвигается на шаг фрагмента dx. Возвращает + добавлено / - удалено.
Private Function SetPages(db As Object, pages As Long, dx As Double, tpl As String, app As Object, src As String) As Long
    Dim lays() As Object, n As Long, l As Object, i As Long, k As Long
    Dim pat As Object, patDb As Object, patPage As Long, tdb As Object
    ReDim lays(1 To db.Layouts.Count)
    For Each l In db.Layouts
        If Not l.ModelType Then n = n + 1: Set lays(n) = l
    Next l
    If n = 0 Then Err.Raise vbObjectError + 514, , "В DWG нет вкладки листа с рамкой."
    For i = 1 To n - 1                                      ' по порядку вкладок
        For k = i + 1 To n
            If lays(k).TabOrder < lays(i).TabOrder Then Set l = lays(i): Set lays(i) = lays(k): Set lays(k) = l
        Next k
    Next i

    If pages > n Then
        If n >= 2 Then
            Set pat = lays(2): Set patDb = db: patPage = 2
        ElseIf Fso.FileExists(tpl) And LCase(tpl) <> LCase(src) Then
            Set tdb = app.GetInterfaceObject("ObjectDBX.AxDbDocument." & Split(app.Version, ".")(0))
            tdb.Open tpl
            For Each l In tdb.Layouts
                If l.TabOrder = 2 Then Set pat = l
            Next l
            If Not pat Is Nothing Then Set patDb = tdb: patPage = 2
        End If
        If pat Is Nothing Then Set pat = lays(1): Set patDb = db: patPage = 1
        Set l = lays(n)
        For k = n + 1 To pages
            Set l = AddPage(db, pat, patDb, (k - patPage) * dx, l, lays(1).Name & " (" & k & ")")
        Next k
        SetPages = pages - n
    ElseIf pages < n Then
        For k = n To pages + 1 Step -1
            kjDrop.Add lays(k).Name
        Next k
        SetPages = pages - n
    End If
    Set tdb = Nothing
End Function

Private Function AddPage(db As Object, pat As Object, patDb As Object, off As Double, prev As Object, nm As String) As Object
    Dim nl As Object, objs() As Object, e As Object, n As Long, res As Variant, i As Long, firstVp As Boolean
    Dim tg As Variant, lk As Boolean, l As Object, ok As Boolean, base As String, s As Long
    base = nm: s = 1
    Do
        ok = True
        For Each l In db.Layouts
            If LCase(l.Name) = LCase(nm) Then ok = False
        Next l
        If ok Then Exit Do
        s = s + 1: nm = base & "-" & s
    Loop
    Set nl = db.Layouts.Add(nm)
    On Error Resume Next
    nl.CopyFrom pat                                         ' параметры печати
    On Error GoTo 0
    ReDim objs(0 To pat.Block.Count - 1)
    For Each e In pat.Block
        Set objs(n) = e: n = n + 1
    Next e
    If n = 0 Then Set AddPage = nl: Exit Function
    If n - 1 < UBound(objs) Then ReDim Preserve objs(0 To n - 1)
    res = patDb.CopyObjects(objs, nl.Block)
    firstVp = True
    For i = LBound(res) To UBound(res)
        Set e = res(i)
        If e.ObjectName = "AcDbViewport" Then
            If firstVp Then
                firstVp = False                             ' общий видовой экран листа
            Else
                lk = e.DisplayLocked
                e.DisplayLocked = False
                tg = e.Target
                tg(0) = tg(0) + off
                e.Target = tg
                e.DisplayLocked = lk
            End If
        ElseIf e.ObjectName = "AcDbBlockReference" Then
            NextNumbers e, prev
        End If
    Next i
    Set AddPage = nl
End Function

' Номер листа и сквозной номер: у предыдущего листа + 1 (если там число).
Private Sub NextNumbers(br As Object, prev As Object)
    Dim at As Variant, pa As Variant, e As Object, i As Long, j As Long, tg As String
    If Not br.HasAttributes Then Exit Sub
    at = br.GetAttributes
    For Each e In prev.Block
        If e.ObjectName = "AcDbBlockReference" Then
            If e.HasAttributes Then pa = e.GetAttributes: Exit For
        End If
    Next e
    If IsEmpty(pa) Then Exit Sub
    For i = LBound(at) To UBound(at)
        tg = UCase(at(i).TagString)
        If tg = "ЛИСТ_№" Or tg = "ЛИСТ№" Or tg = "СКВОЗНАЯ_НУМЕРЦИЯ" Or tg = "СКВОЗНАЯ_НУМЕРАЦИЯ" Then
            For j = LBound(pa) To UBound(pa)
                If UCase(pa(j).TagString) = tg Then
                    If IsNumeric(pa(j).TextString) And Trim(pa(j).TextString) <> "" Then
                        at(i).TextString = CStr(CLng(pa(j).TextString) + 1)
                    End If
                End If
            Next j
        End If
    Next i
End Sub

Private Sub DropLayouts(doc As Object)
    Dim nm As Variant
    For Each nm In kjDrop
        doc.Layouts.Item(CStr(nm)).Delete
    Next nm
End Sub

' Старые связи таблиц с Excel (DATALINK) — удалить: таблица КЖ теперь пишется программой.
Private Sub DropDataLinks(db As Object)
    Dim d As Object, x As Object, c As New Collection
    On Error Resume Next
    Set d = db.Dictionaries.Item("ACAD_DATALINK")
    If d Is Nothing Then Exit Sub
    For Each x In d
        c.Add x
    Next x
    For Each x In c
        x.Delete
    Next x
End Sub

' ---------------------------------------------------------------- текст
Private Function Mt(s As String) As String
    If s = "" Then Exit Function
    s = Replace(s, "\", "\\")
    s = Replace(s, "{", "\{")
    s = Replace(s, "}", "\}")
    s = Replace(s, vbCrLf, "\P")
    s = Replace(s, vbLf, "\P")
    Mt = FONT_PFX & s & "}"
End Function

Private Function Clean(ByVal s As String) As String
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.Pattern = "\\[fFHhCcTtQqWwAap][^;]*;|\\[LlOoKk]|[{}]"
    End If
    s = Replace(s, "\P", " ")
    s = Replace(s, "\\", Chr(1))
    s = Replace(s, "\{", Chr(2))
    s = Replace(s, "\}", Chr(3))
    s = re.Replace(s, "")
    s = Replace(Replace(Replace(s, Chr(1), "\"), Chr(2), "{"), Chr(3), "}")
    Clean = Trim(s)
End Function
