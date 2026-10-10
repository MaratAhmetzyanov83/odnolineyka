# ------------------------------------------------------------------------------------------------
# v5: значения по умолчанию в «Кодах» для шаблона и примера (--defaults): марка, кВт на шт, управление
V5_CODE_DEFAULTS = {
    "R": {"kw": 0.06}, "L": {}, "LED": {"brand": "мс"}, "LD": {"brand": "мс"}, "SH": {"brand": "мс", "ctl": "шторы"},
    "ТП": {"brand": "мс"}, "K": {"brand": "мс"}, "К": {"brand": "мс"}, "ПВ": {"brand": "мс"}, "В": {"brand": "мс"},
    "УВ": {"brand": "мс"}, "ПН": {"brand": "мс"}, "КМН": {"brand": "мс"}, "ПУ": {"brand": "мс"}, "ЭK": {"brand": "мс", "ctl": "клапан"},
    "ЭК": {"brand": "мс", "ctl": "клапан"}, "R.ШТК": {"brand": "мс"}, "R.ЩОПС": {"brand": "мс"}, "R.GSM": {"brand": "мс"},
    "АСУ": {"brand": "мс"}, "PS-KNX": {"brand": "мс"}, "АПС": {"brand": "мс"}, "DALI": {"brand": "мс", "ctl": "DALI"},
    "КУП": {"brand": "ре"}, "PE": {"brand": "ре"}, "НВК": {"brand": "мс"},
}
CODE_HEADS = ["Код", "Наименование", "Опуск, м", "1 - этаж/2 - шлейф", "Пуск: 1 нет/2 двигатель/3 свет", "Марка кабеля",
              "Жил", "Напряжение", "Фаз", "Сечение", "Вид нагрузки", "Группа Кс", "Пожарная",
              "кВт на шт", "Характеристика (B/C/D)", "Вид аппарата утечки", "Управление", "Аппараты",
              "УГО", "Дифзащита, мА (да = 30)", "Мин. автомат, А", "3L клемм на линию"]


def prepare_rows(data):
    """строки таблиц из JSON extract_v3 (v5: маркер группы «гр», один столбец аппаратов, листы вводного узла)."""
    lines = {(ln["file"], ln.get("board"), ln["line"]): ln for ln in data["lines"]}
    board_file = {}
    for p in data["places"]:
        board_file.setdefault(p["board"], p["file"])
    for b in data["boards"]:
        board_file.setdefault(b["name"], b["file"])
    file_board = {}
    for b, f in board_file.items():
        file_board.setdefault(f, b)
    parent, feeds = {}, {}
    notes = []
    for f in data["feeders"]:
        child = file_board.get(f.get("file"))
        if not child:
            notes.append(f"фидер {f['line']}: нет файла {f.get('file')}")
            continue
        if child in parent:
            notes.append(f"фидер {f['line']}: щит {child} уже питается линией {feeds[child][1]} — второй фидер считается обычной линией")
            continue
        parent[child] = file_board.get(f["from_file"])
        feeds[child] = (f["from_file"], f["line"])
    names = set(board_file)
    for p in data["places"]:
        ln = p.get("line")
        if ln in names and ln not in parent and board_file[ln] != p["file"]:
            parent[ln] = p["board"]
            feeds[ln] = (p["file"], ln)
            notes.append(f"щит {ln} подключён к {p['board']} по номеру линии")
    feed_of = {v: k for k, v in feeds.items()}
    order = sorted(board_file, key=lambda b: (0 if not parent.get(b) else 1, b))
    if not order:
        order = ["ЩР"]
    per_sheet = int((data.get("settings") or {}).get("places_per_sheet") or 17)
    bmeta = {b["name"]: b for b in data["boards"]}
    rows, boards = [], []
    for b in order:
        m = bmeta.get(b, {})
        pl = [p for p in data["places"] if p["board"] == b]
        pl.sort(key=lambda p: (p.get("pos") is None, p.get("pos") or 0))
        drawn = [p.get("pos") for p in pl if p.get("mode") != "вручную" and isinstance(p.get("pos"), (int, float))
                 and not str(p.get("line") or "").startswith(("Ввод", "На ИБП", "От ИБП"))]     # вводы и ИБП — часть узла
        nin_old = int(m.get("nin", int(min(drawn)) - 1 if drawn else per_sheet))     # мест под вводной узел
        in_grp = False
        for p in pl:
            line = p.get("line")
            r = {"board": b, "line": line}
            manual = p.get("mode") == "вручную"
            pos = p.get("pos")
            if p.get("unplaced"):
                r["pos"] = "-"
            elif isinstance(pos, (int, float)):
                r["pos"] = int(pos)
            ln = lines.get((p["file"], b, line)) or lines.get((p["file"], None, line))
            qf_man = ln.get("qf") if ln else None
            qt = p.get("qf_txt")
            if qt not in (None, ""):                         # свой автомат (первая строка группы)
                in_grp = not manual
                if str(qt).strip() not in ("авто", "+"):
                    r["qf_in"] = qt
                elif qf_man not in (None, ""):
                    r["qf_in"] = qf_man
            elif line:
                if manual or p.get("unplaced") or not in_grp:
                    if qf_man not in (None, ""):
                        r["qf_in"] = qf_man
                else:
                    r["grp_mark"] = "гр"                     # на автомате строки выше
            if p.get("grp") == "узо":
                r["grp_mark"] = "узо"                         # свой автомат на шине УЗО строки выше
            if manual:
                in_grp = False
            if p.get("rcd") is not None:
                r["rcd_in"] = p["rcd"]
            devs = [str(p.get(k)).strip() for k in ("dev1", "dev2", "dev3") if p.get(k) not in (None, "")]
            if devs:
                r["dev_in"] = "; ".join(devs)
            for k_src, k_dst in (("bus", "bus"), ("module", "module"), ("ch", "ch_in"), ("xt", "xt_in"), ("mode", "mode")):
                if p.get(k_src) is not None:
                    r[k_dst] = p[k_src]
            if p.get("phase") is not None:                  # только ручная фаза; остальное — автобалансировка
                r["phase_in"] = p["phase"]
            if ln:
                for k_src, k_dst in (("name", "name_in"), ("p", "p_in"), ("u", "u_in"), ("n", "n"), ("ph", "ph_in"), ("l_plan", "lplan_in"),
                                     ("l_tray", "tray"), ("l_pvc", "pvc_in"), ("l_pnd", "pnd_in"), ("cable", "brand_in"), ("cores", "cores_in"),
                                     ("s", "s_in"), ("floor", "floor"), ("dwg_pvc", "dwg_pvc"), ("dwg_pnd", "dwg_pnd"),
                                     ("dwg_handles", "dwg_h"), ("dwg_note", "dwg_note"), ("dwg_label", "dwg_label"), ("kc_group", "grp_in"),
                                     ("kc", "kc_in"), ("section", "sec_in"), ("cat_pz", "catpz_in"), ("k_lay", "klay_in"), ("cat", "cat_in"),
                                     ("fire", "fire_in"), ("cos", "cos_in"), ("note", "note")):
                    if ln.get(k_src) is not None:
                        r[k_dst] = ln[k_src]
                if isinstance(r.get("p_in"), str):               # «кВт» и прочий текст вместо числа — не данные
                    try:
                        r["p_in"] = float(r["p_in"].replace(",", ".").strip())
                    except ValueError:
                        r.pop("p_in")
                if str(line).startswith(("Ввод", "На ИБП", "От ИБП")) and r.get("p_in") in (0, 0.0):
                    r.pop("p_in")                                # ручной ноль у ввода — не данные: серое = нагрузка ввода
                r["_ref"] = ln.get("_ref")
            child = feed_of.get((p["file"], line)) or p.get("feeds")
            if child and child != line:
                r["feeds"] = child
            rows.append(r)
        sch, man = m.get("scheme", {}), m.get("manual", {})
        boards.append({"Щит": b, "Питается от": None, "Мест под вводной узел": nin_old, "Вводов": sch.get("inputs"), "АВР": sch.get("avr"),
                       "ИБП от": sch.get("ups_from"),
                       "Питание: от (ТП / сеть)": man.get("from") if not parent.get(b) else None,
                       "ΔU до питающего, %": man.get("du_up"), "Iкз min от сети, А": man.get("ikz_min_up"),
                       "Iкз max от сети, кА": man.get("ikz_max_up"), "Sтр, кВА": man.get("tr_s"), "Схема тр-ра": man.get("tr_scheme"),
                       "Питающий кабель: S, мм²": man.get("s"), "Жил": man.get("cores"), "Кабелей параллельно": man.get("n_par"),
                       "Длина, м": man.get("l"), "Материал": man.get("mat"), "In питающего аппарата, А": man.get("qf_in"),
                       "Хар-ка аппарата": man.get("qf_char"), "_ref": m.get("_ref")})
    return rows, boards, notes


def as_text(v):
    """строка из цифр («00», «0») должна остаться текстом — иначе Excel превратит её в число."""
    if isinstance(v, str) and re.fullmatch(r"\s*[+-]?\d+([.,]\d+)?\s*", v):
        return "'" + v
    return v


def make_table(ws, name, heads, rows_values, first_row=2):
    n = max(1, len(rows_values))
    ncol = len(heads)
    last = letters(ncol)
    ws.Range(f"A{first_row}:{last}{first_row}").Value = [heads]
    if rows_values:
        ws.Range(f"A{first_row + 1}:{last}{first_row + n}").Value = rows_values
    lo = ws.ListObjects.Add(1, ws.Range(f"A{first_row}:{last}{first_row + n}"), None, 1)
    lo.Name = name
    lo.TableStyle = "TableStyleLight1"
    return lo


def set_names(wb, names):
    for nm, ref in names.items():
        try:
            wb.Names(nm).Delete()
        except Exception:
            pass
        wb.Names.Add(nm, ref)


def validation(rng, items, sep):
    try:
        v = rng.Validation
        v.Delete()
        v.Add(3, 1, 1, items if items.startswith("=") else sep.join(items.split("|")))
        v.ShowError = False
        v.IgnoreBlank = True
    except Exception as e:
        log("список не задан:", rng.Address, e)


# ------------------------------------------------------------------------------------------------
# макросы книги (ASCII; русский текст — ^XXXX, расшифровка U())
def enc(s):
    return "".join(ch if ord(ch) < 128 and ch not in "^~" else "^%04X" % ord(ch) for ch in s)


def uq(s):
    return 'U("%s")' % enc(s)


VBA_STD = r'''Option Explicit
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
    Set SxTable = ThisWorkbook.Worksheets(@LINES@).ListObjects(@LINES@)
End Function

Public Function SxTpl() As Worksheet
    On Error Resume Next
    Set SxTpl = ThisWorkbook.Worksheets(@TPL@)
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
        MsgBox @REPAIR_BAD@ & bad, vbExclamation
    Else
        Application.StatusBar = @REPAIR_OK@
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
    Set wsS = wb.Worksheets(@SH_BOARD@)
    Set wsJ = wb.Worksheets(@SH_JOURNAL@)
    Set wsP = wb.Worksheets(@SH_SPEC@)
    Set lo = wb.Worksheets(@SHIELDS@).ListObjects(@SHIELDS@)
    p = wb.Path & Application.PathSeparator & Left$(wb.Name, InStrRev(wb.Name, ".") - 1) & @EXP_SUFFIX@
    oS = wsS.Range("C3").Value
    oJ = wsJ.Range("E1").Value
    oP = wsP.Range("D1").Value
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set ex = Workbooks.Add(-4167)
    For Each c In lo.ListColumns(@COL_BOARD@).DataBodyRange.Cells
        b = Trim$(CStr(c.Value))
        If Len(b) > 0 Then
            wsS.Range("C3").Value = b
            wsJ.Range("E1").Value = b
            wsP.Range("D1").Value = b
            Application.Calculate
            Set sp = SxSpill(wsS.Range("B7"))
            SxCopyBlock wsS.Range(wsS.Range("B6"), wsS.Cells(sp.Row + sp.Rows.Count - 1, 10)), ex, b & @SFX_SCHEME@
            Set sp = SxSpill(wsS.Range("M6"))
            SxCopyBlock wsS.Range(wsS.Range("L6"), wsS.Cells(15, sp.Column + sp.Columns.Count - 1)), ex, b & @SFX_TOTALS@
            Set sp = SxSpill(wsJ.Range("A5"))
            SxCopyBlock wsJ.Range(wsJ.Range("A2"), wsJ.Cells(sp.Row + sp.Rows.Count - 1, 13)), ex, b & @SFX_JOURNAL@
            Set sp = SxSpill(wsP.Range("B3"))
            SxCopyBlock wsP.Range(wsP.Range("B2"), wsP.Cells(sp.Row + sp.Rows.Count - 1, 10)), ex, b & @SFX_SPEC@
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
        MsgBox @EXP_DONE@ & n & vbCrLf & p, vbInformation
    Else
        Application.StatusBar = @EXP_DONE@ & n & " | " & p
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
        MsgBox @EXP_FAIL@ & msg, vbExclamation
    Else
        Application.StatusBar = @EXP_FAIL@ & msg
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
                b = CStr(lo.DataBodyRange.Cells(r, SxCol(lo, @C_BOARD@)).Value)
            End If
        End If
    End If
    On Error GoTo 0
    If b = "" And Application.Visible Then b = InputBox(@ASK_BOARD@, @TITLE@)
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
    ElseIf c = LCase$(@L_SHTORY@) Then
        SxCtlOrder = "4"
    ElseIf c = LCase$(@L_KLAPAN@) Then
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
    cB = SxCol(lo, @C_BOARD@): cL = SxCol(lo, @C_LINE@): cMode = SxCol(lo, @C_MODE@): cSt = SxCol(lo, @C_START@)
    cUz = SxCol(lo, @C_UZ@): cMst = SxCol(lo, @C_MST@): cZone = SxCol(lo, @C_ZONE@): cCtl = SxCol(lo, @C_CTL@)
    cCode = SxCol(lo, @C_CODE@): cRoom = SxCol(lo, @C_ROOM@): cKey = SxCol(lo, @C_SORTKEY@)
    If cKey = 0 Or cB = 0 Then Exit Sub
    v = lo.DataBodyRange.Value
    n = UBound(v, 1)
    ReDim unitOf(1 To n): ReDim uKey(1 To n): ReDim posIn(1 To n)
    For i = 1 To n
        If CStr(v(i, cB)) = board Then
            If first = 0 Then first = i
            ln = Trim$(CStr(v(i, cL)))
            isNode = (CStr(v(i, cMode)) = @MANUAL@) Or (CStr(v(i, cZone)) = @NODEZONE@)
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
    Application.StatusBar = @SORT_DONE@ & board
End Sub

' gray (auto) phases of the board become blue (manual): they no longer change after edits
Public Sub SX_FreezePhases(Optional ByVal board As String = "")
    Dim lo As ListObject, cB As Long, cP As Long, i As Long, c As Range, n As Long
    Set lo = SxTable()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If board = "" Then board = SxBoardOf(lo)
    If board = "" Then Exit Sub
    cB = SxCol(lo, @C_BOARD@): cP = SxCol(lo, @C_PHASE@)
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
    Application.StatusBar = @PH_DONE@ & n
End Sub

Private Function SxSnapSheet() As Worksheet
    On Error Resume Next
    Set SxSnapSheet = ThisWorkbook.Worksheets(@SNAP_SHEET@)
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
            rev = InputBox(@ASK_REV@, @TITLE@, IIf(CStr(ws.Range("B1").Value) = "", "0", CStr(Val(ws.Range("B1").Value) + 1)))
        Else
            rev = IIf(CStr(ws.Range("B1").Value) = "", "0", CStr(Val(ws.Range("B1").Value) + 1))
        End If
    End If
    If Trim$(rev) = "" Then Exit Sub
    If Application.Visible Then
        If MsgBox(@ASK_FREEZE@, vbYesNo + vbQuestion, @TITLE@) = vbYes Then SX_FreezePhases "*"
    End If
    cols = Array(@F_BOARD@, @F_LINE@, @F_NAME@, @F_P@, @F_QF@, @F_S@, @F_LEN@, @F_PHASE@, @F_CABLE@, @F_MOD@, @F_CH@, @F_XT@, @F_PLACE@, @F_SHEET@)
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
    If Application.Visible Then MsgBox @SNAP_DONE@ & n, vbInformation
End Sub

Private Function SxModWidth(ByVal model As String, ByVal dflt As Double) As Double
    Dim r As Variant, w As Variant
    SxModWidth = dflt
    If Trim$(model) = "" Then Exit Function
    On Error Resume Next
    r = Application.Match(model, ThisWorkbook.Worksheets(@CAT_SHEET@).Range("C4:C500"), 0)
    If Not IsError(r) Then
        w = ThisWorkbook.Worksheets(@CAT_SHEET@).Range("L4:L500").Cells(r, 1).Value
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
    If board = "" Then board = Trim$(CStr(ThisWorkbook.Worksheets(@SH_BOARD@).Range("C3").Value))
    If board = "" Then Exit Sub
    cB = SxCol(lo, @C_BOARD@): cL = SxCol(lo, @C_LINE@): cPos = SxCol(lo, @C_POS@): cSt = SxCol(lo, @C_START@)
    cQF = SxCol(lo, @C_QF@): cH = SxCol(lo, @C_H@): cMS = SxCol(lo, @C_MODSTART@): cMM = SxCol(lo, @C_MODEL@)
    cNm = SxCol(lo, @C_NAME@): cMode = SxCol(lo, @C_MODE@)
    cD(1) = SxCol(lo, @C_DEV1@): cD(2) = SxCol(lo, @C_DEV2@): cD(3) = SxCol(lo, @C_DEV3@)
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
                            If Left$(UCase$(a), 3) = @L_UZO@ Then items(1, ni) = SxPoles(a) Else items(1, ni) = SxModWidth(a, 1)
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
    Set ws = ThisWorkbook.Worksheets(@LBL_SHEET@)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    perRail = Val(CStr(ws.Range("C2").Value))
    If perRail <= 0 Then perRail = 24
    ws.Range(ws.Cells(3, 1), ws.Cells(400, 80)).UnMerge
    ws.Range(ws.Cells(3, 1), ws.Cells(400, 80)).Clear
    ws.Range("A1").Value = @LBL_TITLE@ & board
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
    SxLblCap ws, row, @LBL_RAIL@ & rail
    row = row + 1
    For k = 1 To ni
        w = items(1, k)
        If w < 1 Then w = 1
        If used + w > perRail And used > 0 Then
            rail = rail + 1: used = 0: lu = 0: col = 1
            SxLblCap ws, row + 1, @LBL_RAIL@ & rail
            row = row + 2
        ElseIf lu + w > perLine And lu > 0 Then
            lu = 0: col = 1
            SxLblCap ws, row + 1, @LBL_RAIL@ & rail & @LBL_CONT@
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
    Application.StatusBar = @LBL_DONE@ & ni
End Sub
'''

VBA_SHEET = r'''Option Explicit
' table "Lines": gray "auto" cells come back when emptied; "gr" typed into "Breaker" moves to "Group";
' whole rows inserted / deleted -> formulas are repaired
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lo As ListObject, rng As Range, c As Range, h As String, v As String, i As Long
    On Error GoTo done
    Set lo = SxTable()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set rng = Application.Intersect(Target, lo.DataBodyRange)
    If rng Is Nothing Then Exit Sub
    If Target.Columns.Count >= lo.ListColumns.Count Then
        SX_RepairFormulas
        Exit Sub
    End If
    If rng.Cells.Count > 3000 Then Exit Sub
    Application.EnableEvents = False
    For Each c In rng.Cells
        If Not c.HasFormula Then
            h = CStr(lo.HeaderRowRange.Cells(1, c.Column - lo.Range.Column + 1).Value)
            If h = @QF@ Then
                v = Trim$(CStr(c.Formula))
                If v = @GR@ Or v = "=" Or v = "^" Or v = ChrW(8593) Then
                    i = c.Row - lo.DataBodyRange.Row + 1
                    lo.ListColumns(@GRP@).DataBodyRange.Cells(i, 1).Value = @GR@
                    c.ClearContents
                End If
            End If
        End If
    Next
    SX_FillAuto rng
done:
    Application.EnableEvents = True
End Sub
'''

VBA_SHIELDS = r'''Option Explicit
' table "Shields": gray "auto" cells come back when emptied ("Fed from" = by the feeder line)
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lo As ListObject, rng As Range, c As Range, h As String
    On Error GoTo done
    Set lo = Me.ListObjects(@SHIELDS@)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set rng = Application.Intersect(Target, lo.DataBodyRange)
    If rng Is Nothing Then Exit Sub
    If rng.Cells.Count > 3000 Then Exit Sub
    Application.EnableEvents = False
    For Each c In rng.Cells
        If Not c.HasFormula Then
            If Len(Trim$(CStr(c.Formula))) = 0 Then
                h = CStr(lo.HeaderRowRange.Cells(1, c.Column - lo.Range.Column + 1).Value)
                If h = @FEED@ Then c.Formula2 = "=[@[" & @FEEDAUTO@ & "]]"
            End If
        End If
    Next
done:
    Application.EnableEvents = True
End Sub
'''

VBA_WB = r'''Option Explicit
Private Sub Workbook_Open()
    SX_AutoFillOff
End Sub

Private Sub Workbook_Activate()
    SX_AutoFillOff
End Sub

Private Sub Workbook_Deactivate()
    SX_AutoFillRestore
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    SX_AutoFillRestore
End Sub
'''

VBA_REPL = {"@C_BOARD@": uq("Щит"), "@C_LINE@": uq("Номер линии"), "@C_MODE@": uq("Режим"), "@C_START@": uq("_свой автомат"),
            "@C_UZ@": uq("_на шине УЗО"), "@C_MST@": uq("_блок модуля"), "@C_ZONE@": uq("_зона"), "@C_CTL@": uq("_управление (итог)"),
            "@C_CODE@": uq("Код"), "@C_ROOM@": uq("Помещение"), "@C_SORTKEY@": uq("_ключ упорядочивания"), "@C_PHASE@": uq("Фаза"),
            "@MANUAL@": uq("вручную"), "@NODEZONE@": uq("вводной узел"), "@L_SHTORY@": uq("шторы"), "@L_KLAPAN@": uq("клапан"),
            "@L_UZO@": uq("УЗО"), "@ASK_BOARD@": uq("Щит (имя, как в «Линиях»):"), "@TITLE@": uq("Однолинейка v5"),
            "@SORT_DONE@": uq("Щит упорядочен: "), "@PH_DONE@": uq("Фазы закреплены, строк: "), "@SNAP_SHEET@": uq("_выдача"),
            "@ASK_REV@": uq("Номер изменения (0 — первая выдача):"), "@SNAP_DONE@": uq("Выдача зафиксирована, линий: "),
            "@F_BOARD@": uq("Щит"), "@F_LINE@": uq("Номер линии"), "@F_NAME@": uq("Наименование"), "@F_P@": uq("Р, кВт"),
            "@F_QF@": uq("Автомат"), "@F_S@": uq("Сечение, мм²"), "@F_LEN@": uq("Длина, м"), "@F_PHASE@": uq("Фаза"), "@F_CABLE@": uq("Кабель"),
            "@F_MOD@": uq("_Модуль (схема)"), "@F_CH@": uq("Канал (итог)"), "@F_XT@": uq("Клемма"), "@F_PLACE@": uq("Лист / место"),
            "@F_SHEET@": uq("Лист схемы"), "@CAT_SHEET@": uq("Каталог"), "@C_POS@": uq("_№ места (итог)"), "@C_QF@": uq("_QF гр."),
            "@C_H@": uq("Автомат (схема)"), "@C_MODSTART@": uq("_модуль нач."), "@C_MODEL@": uq("_модель гр."), "@C_NAME@": uq("Наименование (итог)"),
            "@C_DEV1@": uq("Аппарат 1"), "@C_DEV2@": uq("Аппарат 2"), "@C_DEV3@": uq("Аппарат 3"), "@LBL_SHEET@": uq("Наклейки"),
            "@LBL_TITLE@": uq("НАКЛЕЙКИ щита "), "@LBL_RAIL@": uq("Рейка "), "@LBL_CONT@": uq(", продолжение"), "@ASK_FREEZE@": uq("Закрепить фазы всех щитов? Тогда правки после выдачи не переставят фазы."), "@LBL_DONE@": uq("Наклеек: "),
            "@SH_BOARD@": uq("Щит"), "@SH_JOURNAL@": uq("Кабельный_журнал"), "@SH_SPEC@": uq("Спецификация"), "@COL_BOARD@": uq("Щит"),
            "@EXP_SUFFIX@": uq("_AutoCAD.xlsx"), "@SFX_SCHEME@": uq(" схема"), "@SFX_TOTALS@": uq(" итоги"), "@SFX_JOURNAL@": uq(" журнал"),
            "@SFX_SPEC@": uq(" спец"), "@EXP_DONE@": uq("Выгружено щитов: "), "@EXP_FAIL@": uq("Выгрузка не удалась: "),
            "@SHIELDS@": uq("Щиты"), "@FEED@": uq("Питается от"), "@FEEDAUTO@": uq("_авто питается от"), "@LINES@": uq("Линии"), "@TPL@": uq("_формулы"), "@QF@": uq("Автомат"), "@GRP@": uq("Группа"), "@GR@": uq("гр"),
            "@REPAIR_BAD@": uq("Не удалось записать формулы столбцов: "),
            "@REPAIR_OK@": uq("Формулы таблицы «Линии» восстановлены")}


def vba_text(s):
    for k, v in VBA_REPL.items():
        s = s.replace(k, v)
    assert all(ord(ch) < 128 for ch in s), "в модуле VBA не-ASCII символы"
    return s


def add_vba(wb, wsl):
    vbp = wb.VBProject
    std = vbp.VBComponents.Add(1)
    std.Name = "SX_v5"
    std.CodeModule.AddFromString(vba_text(VBA_STD))
    vbp.VBComponents(wsl.CodeName).CodeModule.AddFromString(vba_text(VBA_SHEET))
    vbp.VBComponents(wb.CodeName).CodeModule.AddFromString(vba_text(VBA_WB))
    vbp.VBComponents(wb.Worksheets("Щиты").CodeName).CodeModule.AddFromString(vba_text(VBA_SHIELDS))
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "vba")
    os.makedirs(out_dir, exist_ok=True)
    for nm, txt in (("SX_v5.bas", 'Attribute VB_Name = "SX_v5"\r\n' + vba_text(VBA_STD)), ("Линии_лист.cls.txt", vba_text(VBA_SHEET)),
                    ("ЭтаКнига.cls.txt", vba_text(VBA_WB)), ("Щиты_лист.cls.txt", vba_text(VBA_SHIELDS))):
        with open(os.path.join(out_dir, nm), "w", encoding="ascii", newline="\r\n") as fh:
            fh.write(txt)


def add_button(ws, left, top, width, height, text, macro):
    try:
        shp = ws.Shapes.AddFormControl(0, left, top, width, height)      # 0 = кнопка (элемент управления формы)
        shp.OnAction = macro
        try:
            shp.OLEFormat.Object.Caption = text
        except Exception:
            shp.TextFrame.Characters().Text = text
        shp.Placement = 3                                                  # не двигать с ячейками
        return shp
    except Exception as e:
        log("кнопка не создана:", text, e)


CARD = ('=LET(v,FILTER(Линии[Щит],(Линии[_видна]=1)*(Линии[Щит]<>""),""),u,UNIQUE(v),'
        'IF(OR(ROWS(u)<>1,INDEX(u,1)=""),"Объект: щитов "&ROWS(Щиты[Щит])&", линий "&SUM(--(Линии[Номер линии]<>""))&'
        '", замечаний "&COUNTIF(Линии[Замечания],"?*")&CHAR(10)&"Кнопка щита справа — показать один щит и его итоги",'
        'LET(щ,INDEX(u,1),X,LAMBDA(h,XLOOKUP(щ,Щиты[Щит],h,"")),'
        'щ&IF(X(Щиты[Питается от (итог)])<>"","  ← от "&X(Щиты[Питается от (итог)]),"")&"   ·   листов схемы: "&MAXIFS(Линии[Лист схемы],Линии[Щит],щ)&'
        '"   ·   замечаний: "&COUNTIFS(Линии[Щит],щ,Линии[Замечания],"?*")&CHAR(10)&'
        '"Руст "&FIXED(X(Щиты[Руст, кВт]),1)&" · Рр "&FIXED(X(Щиты[Рр, кВт]),1)&" кВт · Iр "&FIXED(X(Щиты[Iр, А]),1)&" А · cos φ "&'
        'FIXED(X(Щиты[cos φ]),2)&" · ΔU до шин "&IFERROR(FIXED(X(Щиты[ΔU до шин, %]),2),"—")&" %")))')


def add_slicer(wb, wsl, lol, n_boards):
    import pythoncom
    try:
        sc = wb.SlicerCaches.Add2(lol, "Щит")
        ncol = max(1, min(n_boards, 12))
        nrow = -(-max(1, n_boards) // ncol)
        h = max(30, 4 + 21 * nrow)
        wsl.Rows(1).RowHeight = max(48, h + 4)
        left = wsl.Range("D1").Left + 3
        width = 8 + 74 * ncol
        sl = sc.Slicers.Add(wsl, pythoncom.Missing, "Кнопки_щитов", "Щит", 2, left, width, h)
        sl.DisplayHeader = False
        sl.NumberOfColumns = ncol
        sl.Top, sl.Left, sl.Width, sl.Height = 2, left, width, h          # при создании координаты не применяются
        try:
            sl.Style = "SlicerStyleLight6"
        except Exception:
            pass
        try:
            sl.ColumnWidth = 70
            sl.RowHeight = 18
        except Exception:
            pass
        try:
            sc.SortItems = 1
            sc.CrossFilterType = 1
        except Exception:
            pass
        return sl
    except Exception as e:
        log("срез «Щит» не создан:", e)


# ------------------------------------------------------------------------------------------------
# «Щит»: то, что уходит в AutoCAD по выбранному щиту — таблица под схемой (строка на место) и итоги по каждому вводу
BOARD_TABLE = ('=LET(щ,$C$3,m,(Линии[Щит]=щ)*ISNUMBER(Линии[_№ места (итог)]),IF(SUM(m)=0,"в щите нет строк на схеме",'
               'LET(pos,FILTER(Линии[_№ места (итог)],m),n,MAX(pos),k,SEQUENCE(n),'
               'f,LAMBDA(c_,LET(v,XLOOKUP(k,pos,FILTER(c_,m),""),IF(v="","",v))),'
               'HSTACK(k,INT((k-1)/Столбцов_на_лист)+1,f(Линии[Номер линии]),f(Линии[_фаза итог]),f(Линии[_Р схемы]),f(Линии[_Ток схемы]),'
               'f(Линии[Наименование (итог)]),f(Линии[Кабель (строкой)]),f(Линии[_автомат гр.])))))')
BOARD_INPUTS = ('=LET(РАСЧЁТ,{R},щ,$C$3,inp,(Линии[Щит]=щ)*(Линии[_ввод щита]=1),k_,SUM(inp),'
                'nm,IF(k_=0,"Щит "&щ,FILTER(Линии[Номер линии],inp,"")),srcs,IF(k_=0,"*",FILTER(Линии[Источник для режимов],inp,"")),'
                'cats,IF(k_=0,XLOOKUP(щ,Щиты[Щит],Щиты[Категория шины «Ввод»],""),FILTER(Линии[Категория ввода (итог)],inp,"")),'
                'own,(Линии[Щит]=щ)*(Линии[Группа Кс (итог)]<>"Нет")*(Линии[Номер линии]<>""),ph,Линии[_фаза итог],pw,IFERROR(--Линии[_Р линии, кВт],0),'
                'ибп,IF(SUM(inp*(Линии[Источник для режимов]="ИБП"))>0,"",LET(u_,XLOOKUP(щ,Щиты[Щит],Щиты[ИБП от],"")&"",IF(u_="","Ввод 1",u_))),'
                'col,LAMBDA(j,LET(s,INDEX(srcs,j),m,AA3#*IF(s="*",1,((AC3#=s)+(AC3#="ИБП")*(ибп=s))>0),x,РАСЧЁТ(m,0,0),'
                'o,own*IF(s="*",1,((Линии[Источник для режимов]=s)+(Линии[Источник для режимов]="ИБП")*(ибп=s))>0),ф1,SUM(o*pw*((ph="L1,L2,L3")/3+(ph="L1"))),'
                'ф2,SUM(o*pw*((ph="L1,L2,L3")/3+(ph="L2"))),ф3,SUM(o*pw*((ph="L1,L2,L3")/3+(ph="L3"))),фс,ф1+ф2+ф3,рр,INDEX(x,5),'
                'кт,INDEX(cats,j),VSTACK(INDEX(nm,j),IF(кт="I особая","I особой категории",IF(кт="","",кт&" категории")),'
                'INDEX(x,1),рр,INDEX(x,8),INDEX(x,3),INDEX(x,2),IF(фс=0,0,рр*ф1/фс),IF(фс=0,0,рр*ф2/фс),IF(фс=0,0,рр*ф3/фс)))),'
                'IF(ROWS(srcs)=1,col(1),REDUCE(col(1),SEQUENCE(ROWS(srcs)-1,1,2),LAMBDA(acc,j,HSTACK(acc,col(j))))))')


def build_board(wb, root, sep):
    ws = wb.Worksheets.Add(After=wb.Worksheets("Линии"))
    ws.Name = "Щит"
    ws.Range("B1").Value = "ЩИТ: данные под однолинейку — таблица под схемой и итоги по каждому вводу (то, что уходит в AutoCAD)"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    ws.Range("B3").Value = "Щит:"
    ws.Range("B3").Font.Bold = True
    ws.Range("C3").Value = root
    ws.Range("C3").Interior.Color = C_IN
    ws.Range("C3").Font.Bold = True
    validation(ws.Range("C3"), "=Список_щитов", sep)
    F2(ws.Range("E3"), '=LET(щ,$C$3,X,LAMBDA(h,XLOOKUP(щ,Щиты[Щит],h,"")),"Питается от: "&IF(X(Щиты[Питается от (итог)])="","—",X(Щиты[Питается от (итог)]))&'
                       '"   ·   вводов: "&X(Щиты[Вводов])&"   ·   АВР: "&X(Щиты[АВР])&"   ·   мест под вводной узел: "&'
                       'IF(X(Щиты[Мест под вводной узел])="",Столбцов_на_лист,X(Щиты[Мест под вводной узел]))&"   ·   замечаний: "&'
                       'COUNTIFS(Линии[Щит],щ,Линии[Замечания],"?*"))')
    ws.Range("E3").Font.Color = 8421504
    # служебные массивы (скрыты): линии щита и нижних щитов, источник (ввод) в выбранном щите
    ws.Range("AA2:AC2").Value = [["в составе", "нижний щит", "ввод в щите"]]
    F2(ws.Range("AA3"), '=--(ISNUMBER(SEARCH("|"&$C$3&"|",Линии[_путь щита]))*(Линии[Номер линии]<>""))')
    F2(ws.Range("AB3"), '=IFERROR(TEXTBEFORE(TEXTAFTER(Линии[_путь щита],"|"&$C$3&"|"),"|"),"")')
    F2(ws.Range("AC3"), '=IF(AB3#="",Линии[Источник для режимов],XLOOKUP(AB3#,Щиты[Щит],Щиты[Источник фидера],"Ввод 1"))')
    ws.Columns("AA:AC").Hidden = True
    # таблица под схемой
    ws.Range("B5").Value = "ТАБЛИЦА ПОД СХЕМОЙ (строка = место на схеме, шаг 20 мм; в AutoCAD таблица поворачивается)"
    ws.Range("B5").Font.Bold = True
    heads = ["Место", "Лист", "Обозначение на плане", "Фаза", "Рн, кВт", "Iн, А", "Наименование электроприёмника", "Кабель", "Автомат"]
    ws.Range("B6:J6").Value = [heads]
    ws.Range("B6:J6").Font.Bold = True
    ws.Range("B6:J6").Interior.Color = C_HEAD
    ws.Range("B6:J6").WrapText = True
    F2(ws.Range("B7"), BOARD_TABLE)
    ws.Range("F7:G400").NumberFormatLocal = "0,00"
    # итоги по вводам
    ws.Range("L5").Value = "ИТОГИ ПО ВВОДАМ (под схемой)"
    ws.Range("L5").Font.Bold = True
    labels = ["Ввод", "Категория", "Руст, кВт", "Рр, кВт", "Iр, А", "cos φ", "Кс", "L1р, кВт", "L2р, кВт", "L3р, кВт"]
    ws.Range("L6:L15").Value = [[x] for x in labels]
    ws.Range("L6:L15").Font.Bold = True
    ws.Range("L6:L15").Interior.Color = C_HEAD
    F2(ws.Range("M6"), BOARD_INPUTS.replace("{R}", RASCHET))
    ws.Range("M8:T15").NumberFormatLocal = "0,00"
    ws.Range("M6:T7").Font.Bold = True
    ws.Range("L17").Value = ("Ввод = строки с кодом «Ввод» этого щита. Линии ввода — по шине и секции (столбцы «Шина», «Секция» в «Линиях»); "
                             "нижние щиты — по своему фидеру; шина ИБП — к вводу из «ИБП от» на листе «Щиты» (если у щита нет отдельного ввода с шины ИБП). "
                             "Итоги — нормальный режим. При АВР у строки «Ввод» в таблице слева — вся нагрузка щита (один ввод тянет всё): по ней подобраны автомат и кабель ввода. "
                             "Категория — у строки «Ввод» (столбец «Категория ввода» в «Ещё») или по схеме щита.")
    ws.Range("L17:P22").Merge()                          # заметка — компактным блоком под итогами
    ws.Range("L17").WrapText = True
    ws.Range("L17").VerticalAlignment = -4160
    ws.Range("L17").Font.Size = 9
    ws.Range("L17").Font.Color = 8421504
    for c_, w_ in (("A", 2), ("B", 7), ("C", 5), ("D", 15), ("E", 10), ("F", 8), ("G", 8), ("H", 44), ("I", 34), ("J", 18), ("K", 3), ("L", 14)):
        ws.Columns(c_).ColumnWidth = w_
    ws.Range("M:T").ColumnWidth = 15
    return ws


# ------------------------------------------------------------------------------------------------
def build(data, base, out, mode="нормы 2026", defaults=False):
    rows, boards, notes = prepare_rows(data)
    for n_ in notes:
        log("  ", n_)
    tmp = os.path.splitext(out)[0] + "_tmp.xlsx"
    shutil.copyfile(base, tmp)
    with ExcelSession() as xl:
        t0 = time.time()
        wb = xl.Workbooks.Open(os.path.abspath(tmp), 0, False)
        xl.Calculation = -4135
        xl.ScreenUpdating = False
        for ws in wb.Worksheets:
            try:
                ws.Unprotect(PW)
            except Exception:
                pass
        for nm in DELETE:
            try:
                wb.Worksheets(nm).Delete()
            except Exception:
                pass
        for src in (wb.LinkSources(1) or []):
            try:
                wb.BreakLink(src, 1)
            except Exception:
                pass
        for i in range(wb.Names.Count, 0, -1):
            nmo = wb.Names(i)
            try:
                if "#REF" in nmo.RefersTo:
                    nmo.Delete()
            except Exception:
                pass
        log("основа подготовлена: %.0f с" % (time.time() - t0))
        sep = xl.International[4]

        # --- «Объект»
        wo = wb.Worksheets.Add(Before=wb.Worksheets(1))
        wo.Name = "Объект"
        obj = [("Объект", "", "название и адрес"),
               ("ρ меди для Iкз min, Ом·мм²/м", data["kz"].get("rho_cu") or 0.0263, "≈ 1,5·ρ20 (ГОСТ 28249 п. 2.4.2)"),
               ("ρ алюминия для Iкз min, Ом·мм²/м", data["kz"].get("rho_al") or 0.042, "≈ 1,5·ρ20"),
               ("Переходные сопротивления на щит, Ом", data["kz"].get("r_contact") or 0.015, "контакты и аппараты на каждый щит"),
               ("c_min для Iкз min", data["kz"].get("c_min") or 0.95, "МЭК 60909-0"),
               ("Методика расчёта", mode, "v3.12 — как старая книга (для сверки); нормы 2026 — СП 256 с Изм. 9 (см. «Методика»)"),
               ("kп.к — региональный коэффициент", 1, "СП 256 табл. 7.5а (ЦФО 0,81, СЗФО 0,91…); для частного дома по умолчанию 1"),
               ("Наименование линии", "по коду и № помещения", "«по коду и № помещения» — «Розеточная группа, пом. 105»; «по экспликации» — название из листа «Экспликация»"),
               ("Писать «пом.» в наименовании", "да", "да — «…, пом. 105»; нет — «… 105»"),
               ("Марка кабеля для пожарных линий", "мфр", "обозначение из «Коды» → «Кабели» (мфр = ВВГнг(А)-FRLS)"),
               ("Запас мощности БП", 1.3, "линии до 50 В: мощность БП ≥ запас × Σ (Uвых·I) — как в листе «Расчет_БП» v1.05"),
               ("Допустимо ΔU в линиях до 50 В, %", 5, "от выхода БП до конца линии (LED 24 В, 12 В)")]
        wo.Range("B1").Value = "ОБЪЕКТ: настройки"
        wo.Range("B1").Font.Bold = True
        wo.Range("B1").Font.Size = 14
        for i, (a, b_, c_) in enumerate(obj):
            r = 3 + i
            wo.Range(f"B{r}").Value = a
            wo.Range(f"C{r}").Value = b_
            wo.Range(f"D{r}").Value = c_
            wo.Range(f"C{r}").Interior.Color = C_IN
        wo.Columns("B").ColumnWidth = 42
        wo.Columns("C").ColumnWidth = 22
        wo.Columns("D").ColumnWidth = 60
        validation(wo.Range("C8"), "v3.12|нормы 2026", sep)
        validation(wo.Range("C10"), "по коду и № помещения|по экспликации", sep)
        validation(wo.Range("C11"), "да|нет", sep)
        set_names(wb, {"ρкз_Cu": "=Объект!$C$4", "ρкз_Al": "=Объект!$C$5", "R_пер": "=Объект!$C$6", "c_min": "=Объект!$C$7",
                       "Методика": "=Объект!$C$8", "k_пк": "=Объект!$C$9", "Имя_режим": "=Объект!$C$10", "Имя_пом": "=Объект!$C$11",
                       "Марка_пож": "=Объект!$C$12", "Запас_БП": "=Объект!$C$13", "dU_LED": "=Объект!$C$14"})
        # модули по умолчанию для управления строки (как в проекте «Жуковка»): шторы — 2 выхода на линию; DALI — шлюз, линии на одной шине
        wo.Range("I1").Value = "МОДУЛИ ПО УМОЛЧАНИЮ (для «Управления» линии; модель — из «Каталога»)"
        wo.Range("I1").Font.Bold = True
        wo.Range("I2:K2").Value = [["Управление", "Модель", "Каналов на линию"]]
        wo.Range("I3:K7").Value = [["Relay", "ZIOMB24V2", 1], ["DIM", "ZDINDX4", 1], ["DALI", "ZDID64X2", 0], ["шторы", "ZIOMB24V2", 2],
                                   ["клапан", "ZCL-8HT230", 1]]
        lmd = wo.ListObjects.Add(1, wo.Range("I2:K7"), None, 1)
        lmd.Name = "Модули_умолч"
        lmd.TableStyle = "TableStyleLight9"
        wo.Columns("I").ColumnWidth = 14
        wo.Columns("J").ColumnWidth = 16
        wo.Columns("K").ColumnWidth = 18
        validation(lmd.ListColumns("Модель").DataBodyRange, "=Каталог!$AB$4:$AB$100", sep)
        try:
            wo.Range("K2").AddComment("0 — линии не занимают каналы (DALI: группы на одной шине шлюза); 2 — шторы (вверх / вниз). Пусто в «Модели» — модуль не подставляется.")
        except Exception:
            pass
        wo.Range("F2:G2").Value = [["Число силовых ЭП", "Кс (СП 256 табл. 7.5)"]]
        k75 = [(1, 1), (2, 1), (3, 0.9), (5, 0.8), (8, 0.75), (10, 0.7), (15, 0.65), (20, 0.65), (30, 0.6), (50, 0.55), (100, 0.55), (200, 0.5)]
        wo.Range("F3:G%d" % (2 + len(k75))).Value = k75
        wo.Range("F3:G%d" % (2 + len(k75))).Interior.Color = C_IN
        wo.Range("F2:G2").Font.Bold = True
        wo.Columns("F").ColumnWidth = 16
        wo.Columns("G").ColumnWidth = 20
        set_names(wb, {"Кс75_n": "=Объект!$F$3:$F$%d" % (2 + len(k75)), "Кс75_k": "=Объект!$G$3:$G$%d" % (2 + len(k75))})

        # --- «Коды»: типовые потребители (код задаёт всё по умолчанию) + УГО, этажи, кабели
        wk = wb.Worksheets.Add(After=wo)
        wk.Name = "Коды"
        wk.Range("A1").Value = ("КОДЫ = ТИПОВЫЕ ПОТРЕБИТЕЛИ (одна строка на код: марка, жилы, мин. сечение, кВт на шт, характеристика, утечка, "
                                "управление, аппараты, УГО, мин. автомат). Справа — этажи, кабели, экспликация (по желанию). Новая строка — сразу под таблицей.")
        wk.Range("A1").Font.Bold = True
        ref = data.get("ref", {})
        codes = {}
        for k_, v_ in ref.get("codes", {}).items():
            v_ = list(v_) + [None] * (13 - len(v_))
            ex = [None] * 5
            if defaults:
                d_ = V5_CODE_DEFAULTS.get(k_, {})
                if d_.get("brand") and not v_[5]:
                    v_[5] = d_["brand"]
                ex = [d_.get("kw"), None, None, d_.get("ctl"), None]
            codes[k_] = v_[:13] + ex + [None] * 4
        if "УЗО" not in codes:                               # строка УЗО на группу автоматов: без кабеля и нагрузки, УЗО в «Аппаратах»
            codes["УЗО"] = ["УЗО", "Шина УЗО (линии со своими автоматами, «узо» в «Группе»)", 0, None, 1, None, None, 230, 1, None, None, "Нет", "нет",
                            None, None, None, None, "УЗО 2P 40А 30мА", None, None, None, None]
        for k_, u_ in ref.get("ugo", {}).items():           # УГО, ток утечки, мин. автомат, 3L клеммы — в ту же строку кода
            u_ = (list(u_) + [None] * 5)[1:5]
            if k_ not in codes:
                codes[k_] = [k_] + [None] * 17 + [None] * 4
            codes[k_][18:22] = u_
        expl = {i: list(e) for i, e in enumerate(data.get("explication", []))}
        blocks = [("Коды", "A", CODE_HEADS, codes),
                  ("Этажи", "Y", ["Этаж", "Название этажа", "Подъём, м", "По другим этажам, м"], ref.get("floors", {})),
                  ("Кабели", "AD", ["Обозначение", "Кабель", "Сечение в марке (1 да / 2 нет)"], ref.get("cables", {})),
                  ("Экспликация", "AH", ["Этаж", "Помещение", "Название"], expl)]
        for tname, c0, heads, vals in blocks:
            i0 = wk.Range(c0 + "1").Column
            rng_h = wk.Range(wk.Cells(3, i0), wk.Cells(3, i0 + len(heads) - 1))
            rng_h.Value = [heads]
            body = [[as_text(x) for x in list(v) + [None] * (len(heads) - len(v))] for v in vals.values()] or [[None] * len(heads)]
            wk.Range(wk.Cells(4, i0), wk.Cells(3 + len(body), i0 + len(heads) - 1)).Value = body
            lo_ = wk.ListObjects.Add(1, wk.Range(wk.Cells(3, i0), wk.Cells(3 + len(body), i0 + len(heads) - 1)), None, 1)
            lo_.Name = tname
            lo_.TableStyle = "TableStyleLight9"
            lo_.HeaderRowRange.WrapText = True
        wk.Rows(3).RowHeight = 45
        wk.Columns.ColumnWidth = 11
        wk.Columns("B").ColumnWidth = 28
        lk = wk.ListObjects("Коды")
        validation(lk.ListColumns("Характеристика (B/C/D)").DataBodyRange, "B|C|D", sep)
        validation(lk.ListColumns("Вид аппарата утечки").DataBodyRange, "дифавтомат|УЗО на группу|нет", sep)
        validation(lk.ListColumns("Управление").DataBodyRange, "Relay|DIM|DALI|шторы|клапан", sep)
        set_names(wb, {"Потребители_обозначение": "=Коды[[Код]:[Сечение]]", "Этаж_обозначение": "=Этажи",
                       "Кабель_обозначение": "=Кабели", "Коды_код": "=Коды[Код]", "Коды_группа": "=Коды[Группа Кс]",
                       "Коды_пожарная": "=Коды[Пожарная]", "Коды_вид": "=Коды[Вид нагрузки]", "Коды_имя": "=Коды[Наименование]",
                       "Коды_режим": "=Коды[1 - этаж/2 - шлейф]", "Коды_квт": "=Коды[кВт на шт]", "Коды_хар": "=Коды[Характеристика (B/C/D)]",
                       "Коды_утечка": "=Коды[Вид аппарата утечки]", "Коды_упр": "=Коды[Управление]", "Коды_апп": "=Коды[Аппараты]",
                       "Коды_уго": "=Коды[УГО]", "Коды_мА": "=Коды[Дифзащита, мА (да = 30)]", "Коды_минавт": "=Коды[Мин. автомат, А]",
                       "Коды_3L": "=Коды[3L клемм на линию]"})
        try:
            wb.Names("УГО_спр").Delete()
        except Exception:
            pass
        log("коды: %d, этажей: %d, кабелей: %d, помещений: %d" % tuple(len(b[3]) for b in blocks))
        wk.ListObjects("Экспликация").HeaderRowRange.Cells(1, 1).AddComment("Экспликация (по желанию): этаж и помещение — как в номере линии (R.1.05.3 → этаж 1, помещение 05). Включается на «Объекте»: «Наименование линии» = «по экспликации».")

        # --- «Щиты», «Шины», «Линии»: шапки + данные
        wsb = wb.Worksheets.Add(After=wk)
        wsb.Name = "Щиты"
        wsl = wb.Worksheets.Add(After=wb.Worksheets("Щиты"))
        wsl.Name = "Линии"
        sheads = [c["head"] for c in S]
        bvals = [[b.get(h) if c["kind"] == "in" else None for c, h in zip(S, sheads)] for b in boards]
        lob = make_table(wsb, "Щиты", sheads, bvals)
        # таблица «Шины» — справа от таблицы щитов (нужна, только если тип шины не определяется по имени)
        c0 = len(S) + 2
        bus_rows = [[b.get("bus"), b.get("board"), b.get("type")] for b in data.get("buses", [])] or [[None, None, None]]
        wsb.Cells(1, c0).Value = "ШИНЫ — только если тип шины не определяется по имени (ИБП, ДГУ) или по секции"
        wsb.Cells(1, c0).Font.Bold = True
        wsb.Range(wsb.Cells(2, c0), wsb.Cells(2, c0 + 2)).Value = [["Шина", "Щит", "Тип"]]
        wsb.Range(wsb.Cells(3, c0), wsb.Cells(2 + len(bus_rows), c0 + 2)).Value = bus_rows
        lsh = wsb.ListObjects.Add(1, wsb.Range(wsb.Cells(2, c0), wsb.Cells(2 + len(bus_rows), c0 + 2)), None, 1)
        lsh.Name = "Шины"
        lsh.TableStyle = "TableStyleLight9"
        validation(lsh.ListColumns("Тип").DataBodyRange, "Ввод 1|Ввод 2|Ввод 3|ИБП|ДГУ|Одиночная", sep)
        IN_KINDS = ("in", "x")
        lheads = [c["head"] for c in L]
        lvals = [[as_text(r.get(c["key"])) if c["kind"] in IN_KINDS else None for c in L] for r in rows]
        lol = make_table(wsl, "Линии", lheads, lvals)
        log("таблицы: щитов %d, строк %d" % (len(boards), len(rows)))
        set_names(wb, {"Список_щитов": "=Щиты[Щит]", "Список_марок": "=Кабели[Обозначение]"})
        # формулы служебных и результатов
        for c in L:
            if c["f"]:
                try:
                    lol.ListColumns(c["head"]).DataBodyRange.Formula2 = expand(c["f"])
                except Exception as e:
                    log("ОШИБКА формулы Линии[%s]: %s\n   %s" % (c["head"], e, expand(c["f"])[:300]))
        for c in S:
            if c["f"]:
                try:
                    lob.ListColumns(c["head"]).DataBodyRange.Formula2 = c["f"]
                except Exception as e:
                    log("ОШИБКА формулы Щиты[%s]: %s\n   %s" % (c["head"], e, c["f"][:300]))
        # серые вводы: формула =Линии[@[_авто …]], поверх — ручные значения из данных
        n_man = 0
        for c in L:
            if c["kind"] in ("a", "xa"):
                lol.ListColumns(c["head"]).DataBodyRange.Formula2 = "=Линии[@[%s]]" % esc(HEAD[c["auto"]])
        # ручные значения поверх: автозаполнение формул в таблицах на время записи выключено,
        # иначе Excel по первой ячейке-формуле заполнит весь столбец и затрёт ручные значения
        af_prev = xl.AutoCorrect.AutoFillFormulasInLists
        xl.AutoCorrect.AutoFillFormulasInLists = False
        try:
            for c in L:
                if c["kind"] not in ("a", "xa"):
                    continue
                rng = lol.ListColumns(c["head"]).DataBodyRange
                for i_, r in enumerate(rows):
                    v = r.get(c["key"])
                    if v not in (None, ""):
                        rng.Cells(i_ + 1, 1).Value = as_text(v)
                        n_man += 1
        finally:
            xl.AutoCorrect.AutoFillFormulasInLists = af_prev
        log("серые вводы: ручных значений %d" % n_man)

        # оформление шапки: ввод — жёлтый, результат — серый
        C_RES = 15132390
        for i, c in enumerate(L):
            hc = lol.HeaderRowRange.Cells(1, i + 1)
            hc.WrapText = True
            hc.Interior.Color = C_IN if c["kind"] in ("in", "a", "x", "xa") else C_RES
            if c["kind"] in ("a", "xa"):
                hc.Font.Italic = True
            if c.get("note"):
                try:
                    hc.AddComment(c["note"])
                except Exception:
                    pass
        for i, c in enumerate(S):
            hc = lob.HeaderRowRange.Cells(1, i + 1)
            hc.WrapText = True
            hc.Interior.Color = C_IN if c["kind"] in ("in", "a") else C_CALC
            if c["kind"] == "a":
                hc.Font.Italic = True
            if c.get("note"):
                try:
                    hc.AddComment(c["note"])
                except Exception:
                    pass
        wsl.Rows(2).RowHeight = 48
        wsb.Rows(2).RowHeight = 60
        wsb.Columns.ColumnWidth = 11
        wsl.Columns.ColumnWidth = 10
        widths = {"board": 9, "line": 13, "name_in": 30, "n": 9, "p_in": 8, "grp_mark": 7, "qf_in": 16, "rcd_in": 8, "ctl_in": 12, "u_in": 12,
                  "dev_in": 14, "s_in": 8, "lplan_in": 9, "phase_in": 9, "xt_in": 8, "issues": 38, "i": 8, "l": 8, "du_vru": 8,
                  "cab_show": 22, "place_lbl": 10}
        for k_, w_ in widths.items():
            wsl.Columns(L.index(next(c for c in L if c["key"] == k_)) + 1).ColumnWidth = w_
        # группы столбцов: «Ещё» — свернуть кнопкой «+» (кнопка над последним видимым столбцом), служебные — скрыть
        try:
            wsl.Outline.SummaryColumn = -4131
        except Exception:
            pass
        for i, c in enumerate(L):
            if c["kind"] in ("x", "xa", "xc"):
                wsl.Columns(i + 1).OutlineLevel = 2
        for i, c in enumerate(S):
            if c["kind"] == "h":
                wsb.Columns(i + 1).OutlineLevel = 2
        # формат чисел у результатов
        for lo, spec in ((lob, S), (lol, L)):
            for c in spec:
                if c["kind"] not in ("in", "a", "x", "xa") and any(u in c["head"] for u in ("кВт", ", А", "%", "Ом", "кА", "квар", "кВА", "cos φ", "Кс", "Ко", "Iкз / In", "Длина")):
                    try:
                        lo.ListColumns(c["head"]).DataBodyRange.NumberFormatLocal = "0,00"
                    except Exception:
                        pass
        lol.ListColumns("Длина, м").DataBodyRange.NumberFormatLocal = "0"
        lol.ListColumns("Ток утечки, мА").DataBodyRange.NumberFormatLocal = "0,0"
        lol.ListColumns("Lмакс по КЗ, м").DataBodyRange.NumberFormatLocal = "0"
        lol.ListColumns("БП: нужно, Вт").DataBodyRange.NumberFormatLocal = "0"
        lol.ListColumns("Uвых БП, В").DataBodyRange.NumberFormatLocal = "0,0"
        # выпадающие списки (подсказки, не запрет)
        validation(lol.ListColumns("Щит").DataBodyRange, "=Список_щитов", sep)
        validation(lol.ListColumns("Группа").DataBodyRange, "гр|узо", sep)
        validation(lol.ListColumns("Утечка, мА").DataBodyRange, "10|30|100|300|нет", sep)
        validation(lol.ListColumns("Управление").DataBodyRange, "Relay|DIM|DALI|шторы|клапан", sep)
        validation(lol.ListColumns("Режим").DataBodyRange, "вручную|рисовать", sep)
        validation(lol.ListColumns("Марка кабеля").DataBodyRange, "=Список_марок", sep)
        # условное форматирование «Линий» (формулы — в записи русского Excel)
        lastr = R0 + max(1, len(rows)) + 300
        body = wsl.Range(f"A{R0}:{letters(len(L))}{lastr}")
        cf = body.FormatConditions
        cf.Delete()
        Lb, Lz, Ls, Ln = LETTER["bno"], LETTER["zone"], LETTER["sheetno"], LETTER["_n"]
        c1 = cf.Add(2, None, f'=${Lz}{R0}="вводной узел"')
        c1.Interior.Color = 16247773
        c2 = cf.Add(2, None, f'=НЕЧЁТ(${Lb}{R0})')
        c2.Interior.Color = 15921906
        c3 = cf.Add(2, None, f'=И(${Ln}{R0}>1;$A{R0}<>$A{R0 - 1})')
        c3.Borders(-4160).LineStyle = 1
        c3.Borders(-4160).Color = 0
        c4 = cf.Add(2, None, f'=И($A{R0}=$A{R0 - 1};${Ls}{R0}<>${Ls}{R0 - 1};${Ls}{R0}<>"")')
        c4.Borders(-4160).LineStyle = -4115
        c4.Borders(-4160).Color = 8421504
        for i, c in enumerate(L):
            col_ = letters(i + 1)
            rng = wsl.Range(f"{col_}{R0}:{col_}{lastr}")
            if c["kind"] in ("a", "xa"):
                g_ = rng.FormatConditions.Add(2, None, f'=ЕФОРМУЛА({col_}{R0})')
                g_.Font.Color = 8421504
                g_.Font.Italic = True
                b_ = rng.FormatConditions.Add(2, None, f'=И(НЕ(ЕФОРМУЛА({col_}{R0}));{col_}{R0}<>"")')
                b_.Font.Color = 12611584
                b_.Font.Bold = True
            if c["key"] == "issues":
                e_ = rng.FormatConditions.Add(2, None, f'={col_}{R0}<>""')
                e_.Interior.Color = 13421823
            if c["key"] == "grp_mark":
                m_ = rng.FormatConditions.Add(2, None, f'={col_}{R0}<>""')
                m_.Font.Color = 12611584
                m_.Font.Bold = True
        for i, c in enumerate(S):
            if c["kind"] == "a":
                col_ = letters(i + 1)
                rng = wsb.Range(f"{col_}3:{col_}{3 + len(boards) + 100}")
                g_ = rng.FormatConditions.Add(2, None, f'=ЕФОРМУЛА({col_}3)')
                g_.Font.Color = 8421504
                g_.Font.Italic = True
                b_ = rng.FormatConditions.Add(2, None, f'=И(НЕ(ЕФОРМУЛА({col_}3));{col_}3<>"")')
                b_.Font.Color = 12611584
                b_.Font.Bold = True
        # служебные столбцы — скрыть
        for i, c in enumerate(L):
            if c["kind"] == "h":
                wsl.Columns(i + 1).Hidden = True
        try:
            wsl.Outline.ShowLevels(0, 1)
            wsb.Outline.ShowLevels(0, 1)
        except Exception:
            pass

        # --- «Однолинейка»: интерфейс для SXDRAW (спилл столбцов «Линий»)
        w1 = wb.Worksheets("Однолинейка")
        w1.Range("A3:CA1000").ClearContents()
        for c_, k in ONE_MAP.items():
            w1.Range(f"{c_}3").Formula2 = '=LET(v,FILTER(Линии[%s],Линии[_№ места (итог)]<>""),IF(v="","",v))' % esc(HEAD[k])
        w1.Range("AE1").Formula2 = '="Ошибок: "&COUNTIF($BU$3:$BU$3000,2)&"   Предупреждений: "&COUNTIF($BU$3:$BU$3000,1)'

        root = next((b["Щит"] for b in boards if not b.get("Питается от")), boards[0]["Щит"])
        build_loads(wb, xl, root, sep)
        build_pz(wb, root, sep)
        extend_cable_diameters(wb)
        build_journal_v5(wb, sep)
        spec_by_board(wb, sep)
        build_spec_v5(wb, sep)
        build_board(wb, root, sep)
        build_tags(wb)
        single_line_view(wb)
        # --- «Проверка»
        wp = wb.Worksheets.Add(Before=wb.Worksheets(1))
        wp.Name = "Проверка"
        wp.Range("A1").Value = "ПРОВЕРКА ОБЪЕКТА"
        wp.Range("A1").Font.Bold = True
        wp.Range("A1").Font.Size = 14
        wp.Range("A2").Formula2 = '="Замечаний: щиты — "&COUNTIF(Щиты[Замечания щита],"?*")&", линии — "&COUNTIF(Линии[Замечания],"?*")&". Методика: "&Методика'
        wp.Range("A4:B4").Value = [["Щит", "Замечания щита"]]
        wp.Range("A5").Formula2 = '=FILTER(HSTACK(Щиты[Щит],Щиты[Замечания щита]),Щиты[Замечания щита]<>"","замечаний нет")'
        wp.Range("D4:H4").Value = [["Щит", "Номер линии", "Наименование", "Замечания", "Перейти"]]
        wp.Range("D5").Formula2 = ('=FILTER(HSTACK(Линии[Щит],Линии[Номер линии],Линии[Наименование (итог)],Линии[Замечания],'
                                   'HYPERLINK("#\'Линии\'!B"&(Линии[_n]+2),"→")),Линии[Замечания]<>"","замечаний нет")')
        for a in ("A4:B4", "D4:H4"):
            wp.Range(a).Font.Bold = True
            wp.Range(a).Interior.Color = C_HEAD
        for c_, w_ in (("A", 12), ("B", 70), ("D", 12), ("E", 16), ("F", 40), ("G", 70), ("H", 8)):
            wp.Columns(c_).ColumnWidth = w_
        wp.Columns("B").WrapText = True
        wp.Columns("G").WrapText = True

        build_changes_v5(wb, sep)
        build_labels_sheet(wb)

        # --- «Линии»: карточка щита, кнопки щитов, закрепление
        wsl.Range("A1:C1").Merge()
        wsl.Range("A1").Formula2 = CARD
        wsl.Range("A1").WrapText = True
        wsl.Range("A1").VerticalAlignment = -4108
        wsl.Range("A1").Font.Size = 10
        wsl.Range("A1:C1").Interior.Color = 15921906
        sl = add_slicer(wb, wsl, lol, len(boards))
        bl = (sl.Left + sl.Width + 8) if sl is not None else wsl.Range("D1").Left + 4
        add_button(wsl, bl, 4, 120, 22, "Починить формулы", "SX_RepairFormulas")
        add_button(wsl, bl + 125, 4, 120, 22, "Упорядочить щит", "SX_SortBoard")
        add_button(wsl, bl + 250, 4, 115, 22, "Закрепить фазы", "SX_FreezePhases")
        add_button(wsl, bl + 370, 4, 150, 22, "Зафиксировать выдачу", "SX_Snapshot")
        add_button(wp, wp.Range("D1").Left, 4, 130, 22, "Починить формулы", "SX_RepairFormulas")
        wsc = wb.Worksheets("Щит")
        add_button(wsc, wsc.Range("L1").Left, 4, 190, 24, "Выгрузить для AutoCAD (все щиты)", "SX_ExportAutoCAD")
        add_button(wsc, wsc.Range("L1").Left + 200, 4, 120, 24, "Наклейки (A4)", "SX_Labels")
        wch = wb.Worksheets("Изменения")
        add_button(wch, wch.Range("J1").Left, 4, 150, 22, "Зафиксировать выдачу", "SX_Snapshot")
        wsl.Activate()
        xl.ActiveWindow.SplitRow = 2
        xl.ActiveWindow.SplitColumn = 3
        xl.ActiveWindow.FreezePanes = True

        xl.Calculation = -4105
        t1 = time.time()
        xl.CalculateFull()
        log("пересчёт: %.1f с" % (time.time() - t1))
        for ws_ in (wsl, wsb, w1):
            try:
                e = ws_.UsedRange.SpecialCells(-4123, 16)
                log("ошибки формул на «%s»: %d (%s)" % (ws_.Name, e.Count, e.Address[:200]))
            except Exception:
                log("ошибки формул на «%s»: 0" % ws_.Name)
            try:
                cr = ws_.CircularReference
                if cr is not None:
                    log("ЦИКЛИЧЕСКАЯ ССЫЛКА на «%s»: %s" % (ws_.Name, cr.Address))
            except Exception:
                pass
        res = compare(wsl, lol, wsb, lob, rows, boards)

        # --- шаблоны формул для макроса «Починить формулы» (лист скрыт)
        wt = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        wt.Name = "_формулы"
        tpl = [["Столбец", "Формула (R1C1)", "Вид", "Серое из"]]
        for c in L:
            if c["kind"] in ("a", "xa"):
                tpl.append([c["head"], "'=Линии[@[%s]]" % esc(HEAD[c["auto"]]), c["kind"], HEAD[c["auto"]]])
            elif c["f"]:
                f1 = lol.ListColumns(c["head"]).DataBodyRange.Cells(1, 1).Formula2R1C1
                tpl.append([c["head"], "'" + f1, "c" if c["kind"] == "xc" else c["kind"], None])
        wt.Range("A1:D%d" % len(tpl)).Value = tpl
        wt.Visible = 2

        # гайд, методика, порядок листов
        import guide_v5
        import guide_v312
        import methodology_sheet
        guide_v312.G = guide_v5.G
        guide_v312.build(wb)
        methodology_sheet.render(wb)
        put_guides(wb)
        tidy_reference(wb)
        try:
            wb.Worksheets("Нагрузки").Name = "Режимы"
        except Exception:
            pass
        order = ["Гайд", "Проверка", "Объект", "Щиты", "Линии", "Коды", "Щит", "Кабельный_журнал", "Спецификация", "Бирки", "Изменения",
                 "Наклейки", "Каталог",
                 "Методика", "Таблицы_норм", "Режимы", "Расчёт_нагрузок", "Спецификация_авто", "Однолинейка", "Справочная_потребители"]
        for nm in reversed(order):
            try:
                wb.Worksheets(nm).Move(Before=wb.Worksheets(1))
            except Exception:
                pass
        for nm in ("Справочная_потребители", "Спецификация_авто", "Однолинейка", "_выдача"):     # служебные: читают программы и макросы
            try:
                wb.Worksheets(nm).Visible = 0
            except Exception:
                pass
        green, blue, grey, violet, yellow = 5287936, 12419407, 10921638, 10498160, 13431551
        for nm, color in (("Гайд", green), ("Проверка", 255), ("Объект", yellow), ("Щиты", yellow), ("Линии", yellow), ("Коды", yellow),
                          ("Щит", blue), ("Кабельный_журнал", blue), ("Спецификация", blue), ("Бирки", blue), ("Изменения", blue), ("Наклейки", blue), ("Каталог", grey),
                          ("Методика", green), ("Таблицы_норм", grey), ("Режимы", violet), ("Расчёт_нагрузок", violet)):
            try:
                wb.Worksheets(nm).Tab.Color = color
            except Exception:
                pass
        add_vba(wb, wsl)
        wb.Worksheets("Линии").Activate()
        wb.Worksheets("Линии").Range("D3").Select()
        if os.path.exists(out):
            os.remove(out)
        wb.SaveAs(os.path.abspath(out), 52)
        log("сохранено: %s (%.0f с)" % (out, time.time() - t0))
        wb.Close(False)
    try:
        os.remove(tmp)
    except Exception:
        pass
    return res


def compare(wsl, lol, wsb, lob, rows, boards):
    """сверка с результатами v3.12 (rows[i]['_ref'], boards[i]['_ref'])."""
    def colvals(lo, head):
        v = lo.ListColumns(head).DataBodyRange.Value
        return [x[0] for x in v] if isinstance(v, tuple) else [v]
    keys = (("p", "p"), ("i", "i"), ("qf", "in_grp"), ("s", "s"), ("l", "l"), ("du", "du"), ("du_vru", "du_vru"), ("grp", "grp"), ("cos", "cos"))
    got = {kg: colvals(lol, HEAD[kg]) for _, kg in keys}
    stats, diffs = {}, []
    for i, r in enumerate(rows):
        ref = r.get("_ref") or {}
        for kr, kg in keys:
            a, b = ref.get(kr), got[kg][i]
            if a in (None, "") and b in (None, ""):
                continue
            ok = (abs(float(a) - float(b)) <= max(0.011, 0.001 * abs(float(a)))) if isinstance(a, (int, float)) and isinstance(b, (int, float)) else (str(a) == str(b))
            st = stats.setdefault(kr, [0, 0])
            st[0 if ok else 1] += 1
            if not ok and len(diffs) < 40:
                diffs.append(f"{r['board']} {r.get('line')}: {kr} v3.12={a} v5={b}")
    log("СВЕРКА линий с v3.12 (совпало / не совпало):", {k: tuple(v) for k, v in stats.items()})
    for d in diffs:
        log("   ", d)
    bgot = {h: colvals(lob, h) for h in ("Руст, кВт", "Руст быт, кВт", "Рр сил, кВт", "Рр, кВт", "Iр, А", "cos φ", "ΔU до шин, %")}
    bkeys = (("p_inst", "Руст, кВт"), ("p_byt", "Руст быт, кВт"), ("p_sil", "Рр сил, кВт"), ("pr", "Рр, кВт"), ("ir", "Iр, А"), ("cos", "cos φ"), ("du_bus", "ΔU до шин, %"))
    for i, b in enumerate(boards):
        ref = b.get("_ref") or {}
        bad = []
        for kr, h in bkeys:
            a, g = ref.get(kr), bgot[h][i]
            if isinstance(a, (int, float)) and isinstance(g, (int, float)):
                if abs(a - g) > max(0.011, 0.001 * abs(a)):
                    bad.append(f"{kr}: {a:.3f} → {g:.3f}")
            elif a not in (None, "") or g not in (None, "", 0):
                if str(a) != str(g):
                    bad.append(f"{kr}: {a} → {g}")
        log("  щит %-12s %s" % (b["Щит"], "OK" if not bad else "; ".join(bad)))
    return stats


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--data", required=True)
    ap.add_argument("--base", required=True)
    ap.add_argument("--out", required=True, help="файл .xlsm")
    ap.add_argument("--mode", default="нормы 2026", help="v3.12 или «нормы 2026»")
    ap.add_argument("--defaults", action="store_true", help="заполнить «Коды» значениями v5 по умолчанию (шаблон, пример)")
    a = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    with open(a.data, encoding="utf-8") as fh:
        data = json.load(fh)
    build(data, a.base, a.out, a.mode, a.defaults)


if __name__ == "__main__":
    main()
