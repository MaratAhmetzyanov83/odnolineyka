# -*- coding: utf-8 -*-
"""Категория шины по цепочке источников: ввод(ы), ДГУ, ИБП (книги v3.12 после patch_inputs.py).

    py patch_categories.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит

Правило (ПУЭ 1.2.17–1.2.21): 1 источник — III; 2 и более, переключение вручную — II;
2 с автоматикой (АВР, ИБП) — I; третий автоматический источник (ДГУ с АВР, АКБ ИБП) — I особая.
Было: шина «ИБП» / «ДГУ» — всегда I особая (1 ввод + ДГУ — это I, а не I особая).

«Сеть_щита», справа от раздела 6 — H64:L71 «ДГУ И ИБП»: по щиту 1–3 (столбцы I:K)
  65 ДГУ есть (да / нет; пусто = да, если у щита есть шина ДГУ)
  66 ДГУ: АВР (да / нет; пусто = да)
  67 ДГУ резервирует (свою шину / всю основную; пусто = свою шину)
  68 ИБП питается от шины (основной / ДГУ; пусто = основной)
  69–71 итог: категория шины «Ввод», «ДГУ», «ИБП».
C68:E68 (СХ_кат_ввод) = строка 69; C69:E69 — «ДГУ / ИБП» для справки. Имена СХ_кат_ДГУ, СХ_кат_ИБП.
«Исходные данные» BU «Категория обеспечена» — по типу шины линии из этих итогов.
«Исходные данные» BP «Категория (итог)»: BN (вручную), пусто — категория шины линии (BU). Отсюда «Нагрузка_щитов» AS
  и «Режимы», блок «По категориям».
«Однолинейка» CA «Категория» и CB «Категория: проверка» — то же по строке линии (сами).
Свой скрытый экземпляр Excel; закрывается только он. COM — только позиционные аргументы.
"""
import os
import signal
import sys

PWD = "sx"
NET, SRC = "Сеть_щита", "Исходные данные"
MARK = "ДГУ И ИБП (для категорий)"
SCH = "CDE"          # раздел 6: щит 1 / 2 / 3
COL = "IJK"          # новый блок: щит 1 / 2 / 3


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def rgb(t):
    return t[0] + t[1] * 256 + t[2] * 65536


def up(x):
    """+1 автоматический источник (АВР ДГУ, АКБ ИБП)"""
    return f'IF(OR({x}="I",{x}="I особая"),"I особая","I")'


def build(wb, log):
    ws = wb.Worksheets(NET)
    if ws.Range("H64").Value == MARK:
        return False
    if wb.Application.WorksheetFunction.CountA(ws.Range("H64:L71")) > 0:
        raise RuntimeError(f"«{NET}» H64:L71 занято")
    ws.Range("H64").Value = MARK
    ws.Range("H64").Font.Bold = True
    rows = {65: ("ДГУ есть (да / нет)", "пусто = да, если у щита есть шина ДГУ («Однолинейка» D или таблица «Шины щита»)"),
            66: ("ДГУ: АВР (да / нет)", "автоматический пуск ДГУ и переключение; пусто = да. Нет — переключение вручную"),
            67: ("ДГУ резервирует (свою шину / всю основную)",
                 "свою шину — шина «ДГУ» (гарантированного питания): АВР ввод–ДГУ; всю основную — АВР ДГУ на основную шину"),
            68: ("ИБП питается от шины (основной / ДГУ)", "пусто = основной. Вход ИБП в «Режимах» — как раньше, по C67"),
            69: ("Категория шины «Ввод»", "1 источник — III; 2+ вручную — II; 2 с АВР — I; 3 автоматических — I особая"),
            70: ("Категория шины «ДГУ»", "ввод(ы) + ДГУ: с АВР — на ступень выше (III → I), без АВР — II"),
            71: ("Категория шины «ИБП»", "шина питания ИБП + АКБ: III/II → I, I → I особая")}
    for r, (lab, note) in rows.items():
        ws.Range(f"H{r}").Value = lab
        ws.Range(f"L{r}").Value = note
    for s in (1, 2, 3):
        c, sc = COL[s - 1], SCH[s - 1]
        ws.Range(f"{c}64").Formula = f"=$B${5 + s}"
        g = f'LOWER(TRIM({c}$65&""))<>"нет"'
        ga = f'LOWER(TRIM({c}$66&""))<>"нет"'
        gm = f'ISNUMBER(SEARCH("основн",{c}$67&""))'
        ws.Range(f"{c}69").Formula2 = (
            f'=IF({c}$64="","",LET(k_,MAX(1,N({sc}$65)),av_,LOWER(TRIM({sc}$66&""))="да",'
            f'd_,AND({g},{gm}),n_,k_+d_,a_,IF(av_,k_,1)+AND(d_,{ga}),'
            f'IF(a_>=3,"I особая",IF(a_=2,"I",IF(n_>=2,"II","III")))))')
        ws.Range(f"{c}70").Formula2 = (
            f'=IF({c}$64="","",IF({gm},{c}$69,IF({ga},{up(f"{c}$69")},IF({c}$69="III","II",{c}$69))))')
        ws.Range(f"{c}71").Formula2 = (
            f'=IF({c}$64="","",LET(f_,IF(ISNUMBER(SEARCH("ДГУ",{c}$68&"")),{c}$70,{c}$69),{up("f_")}))')
        # раздел 6: шина «Ввод» — отсюда; «ИБП, ДГУ» — для справки
        ws.Range(f"{sc}68").Formula = f'=IF({sc}$64="","",{c}$69)'
        ws.Range(f"{sc}69").Formula = f'=IF({sc}$64="","",{c}$70&" / "&{c}$71)'
    ws.Range("B69").Value = "Обеспечиваемая категория: шина «ДГУ» / «ИБП»"
    ws.Range("G68").Value = "с учётом ДГУ на всю основную шину (блок «ДГУ и ИБП» справа, H64)"
    ws.Range("G69").Value = "по цепочке источников (блок «ДГУ и ИБП» справа, H64); было: всегда I особая"
    inp = ws.Range("I65:K68")
    inp.Interior.Color = rgb((255, 242, 204))                        # жёлтое — ввод
    for r, lst in ((65, "да,нет"), (66, "да,нет"), (67, "свою шину,всю основную"), (68, "основной,ДГУ")):
        v = ws.Range(f"I{r}:K{r}").Validation
        v.Delete()
        v.Add(3, 1, 1, lst)
    ws.Range("H64:L64").Font.Bold = True
    ws.Range("H64:L64").Interior.Color = rgb((221, 235, 247))
    ws.Range("H69:K71").Font.Bold = True
    ws.Range("H64:L71").Borders.LineStyle = 1
    ws.Range("L65:L71").Font.Color = rgb((89, 89, 89))
    for col, w in zip("HIJKL", (34, 11, 11, 11, 60)):
        if ws.Columns(col).ColumnWidth < w:
            ws.Columns(col).ColumnWidth = w
    wb.Names.Add("СХ_кат_ДГУ", f"='{NET}'!$I$70:$K$70")
    wb.Names.Add("СХ_кат_ИБП", f"='{NET}'!$I$71:$K$71")
    log.append(f"«{NET}»: блок «ДГУ и ИБП» H64:L71, имена СХ_кат_ДГУ, СХ_кат_ИБП; C68:E69 — из блока")
    return True


BU = ('=IF($D3="","",IF($BS3="Одиночная","III",IF($BZ3="","III",IFERROR(INDEX('
      'IF($BS3="ИБП",СХ_кат_ИБП,IF($BS3="ДГУ",СХ_кат_ДГУ,СХ_кат_ввод)),MATCH($BZ3,ПЩ_щиты,0)),"III"))))')


def patch_bu(wb, log):
    ws = wb.Worksheets(SRC)
    if "СХ_кат_ИБП" in (ws.Range("BU3").Formula or ""):
        return
    ws.Range("BU3:BU1000").Formula = BU
    log.append(f"«{SRC}» BU: категория шины ДГУ / ИБП — по цепочке источников")


BP = ('=IF($D3="","",IF(TRIM($BN3&"")="",IF($BU3="","III",$BU3),'
      'IFERROR(INDEX({"I особая","I","II","III"},MATCH(TRIM($BN3&""),{"I особая","I","II","III"},0)),"III")))')
BP_HDR = "Категория (итог: BN, пусто = по шине)"


def patch_bp(wb, log):
    ws = wb.Worksheets(SRC)
    if "$BU3" in (ws.Range("BP3").Formula or ""):
        return
    ws.Range("BP3:BP1000").Formula = BP
    ws.Range("BP2").Value = BP_HDR
    log.append(f"«{SRC}» BP: пусто в BN = категория по шине (BU)")


ODN = "Однолинейка"
ODN_HDR = ("КАТЕГОРИЯ", "Категория\n(сама: по шине;\nвручную — «Исходные\nданные» BN)", "Категория:\nпроверка")


def patch_odn(wb, log):
    ws = wb.Worksheets(ODN)
    if ws.Range("CA1").Value == ODN_HDR[0]:
        return False
    if wb.Application.WorksheetFunction.CountA(ws.Range("CA1:CB1000")) > 0:
        raise RuntimeError(f"«{ODN}» CA:CB занято")
    look = "MATCH($R3,'Исходные данные'!$D$3:$D$1000,0)"
    ws.Range("CA1").Value = ODN_HDR[0]
    ws.Range("CA2").Value = ODN_HDR[1]
    ws.Range("CB2").Value = ODN_HDR[2]
    ws.Range("CA3:CA1000").Formula = f"=IF($R3=\"\",\"\",IFERROR(INDEX('Исходные данные'!$BP$3:$BP$1000,{look}),\"\"))"
    ws.Range("CB3:CB1000").Formula = f"=IF($R3=\"\",\"\",IFERROR(INDEX('Исходные данные'!$BV$3:$BV$1000,{look}),\"\"))"
    ws.Range("CA1:CB2").Font.Bold = True
    ws.Range("CA2:CB2").WrapText = True
    ws.Range("CA2:CB2").Interior.Color = rgb((221, 235, 247))
    fc = ws.Range("CB3:CB1000").FormatConditions.Add(1, 3, '="НЕ ОБЕСПЕЧЕНА"')
    fc.Font.Color = rgb((192, 0, 0))
    fc.Font.Bold = True
    ws.Columns("CA").ColumnWidth = 14
    ws.Columns("CB").ColumnWidth = 14
    log.append(f"«{ODN}» CA:CB: категория линии и проверка (сами)")
    return True


def unprot(ws):
    """снять защиту; вернуть её параметры (как в patch_audit.py) или None"""
    if not ws.ProtectContents:
        return None
    p = ws.Protection
    opts = (True, True, True, False,
            p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
            p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks,
            p.AllowDeletingColumns, p.AllowDeletingRows, p.AllowSorting,
            p.AllowFiltering, p.AllowUsingPivotTables)
    ws.Unprotect(PWD)
    return opts


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        app.Calculation = -4135
        try:
            wb.Worksheets(NET)
        except Exception:
            raise RuntimeError(f"нет листа «{NET}» — сначала merge_network.py")
        pr = [(ws, unprot(ws)) for ws in (wb.Worksheets(NET), wb.Worksheets(SRC), wb.Worksheets(ODN))]
        build(wb, log)
        patch_bu(wb, log)
        patch_bp(wb, log)
        patch_odn(wb, log)
        for ws, opts in pr:
            if opts:
                ws.Protect(PWD, *opts)
        app.Calculation = -4105
        app.CalculateFull()
        wb.Save()
        wb.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        try:
            os.kill(pid, signal.SIGTERM)
        except Exception:
            pass
    print(os.path.basename(path))
    for s in log or ["изменений нет"]:
        print("  " + s)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for p in sys.argv[1:]:
        run(p)
