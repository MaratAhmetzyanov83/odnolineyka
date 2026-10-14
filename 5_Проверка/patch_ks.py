# -*- coding: utf-8 -*-
"""Кс силовых линий — сам, по числу электроприёмников в щите (как «группы по характеру» в nanoCAD BIM Электро).

    py patch_ks.py <книга.xlsx> [...]      — книга закрыта; повторный запуск ничего не портит

Было: Кс линии «Сил» = «Исходные данные» BA (руками), пусто = 1.
Стало: Кс линии «Сил» = BA (ручной, главнее) -> иначе по таблице СП 256 по числу ЭП -> иначе 1.
  «Справочная» CE:CF — код линии -> таблица Кс (ЛФ — Лифт; НС, ПВ, ПУ, В — Двиг). Жёлтые — можно менять.
  «Справочная» CH:CN — Кс по числу ЭП n: Лифт (до 12 эт.), Лифт 12+, Двиг (сантехника, вентиляция, насосы);
                       CL:CN — пустые столбцы для своих таблиц (пищеблок и т. п.). Между точками — по меньшему n.
  «Справочная» CI20  — «Этажей 12 и более» (да/нет) — какой столбец лифтов брать.
  «Нагрузка_щитов» BB — таблица Кс линии, BC — n (линий «Сил» с той же таблицей в щите и секции, P > 0),
                   BD — Кс по таблице; B (Кс линии) берёт BD, если BA пусто.
Нижестоящие щиты приходят готовым Рр сил (Кс посчитан в своём щите) — n через щиты не суммируется (в запас).
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import signal
import sys

PWD = "sx"
SPR = "Справочная"
NAG = "Нагрузка_щитов"
SRC = "Исходные данные"
TITLE = "Кс СИЛОВЫХ (СП 256)"

MAP = [("ЛФ", "Лифт"), ("НС", "Двиг"), ("ПВ", "Двиг"), ("ПУ", "Двиг"), ("В", "Двиг")]
MAP_ROWS = 12                                    # CE4:CF15
# n, Лифт до 12 эт., Лифт 12 эт. и более, Двигатели (СП 256 / СП 31-110 табл. 6.4 и 6.9; между точками — по меньшему n)
TAB = [(1, 1, 1, 1), (2, 0.8, 0.9, 1), (3, 0.8, 0.9, 0.9), (4, 0.7, 0.8, 0.9), (5, 0.7, 0.8, 0.8),
       (6, 0.65, 0.75, 0.8), (8, 0.65, 0.75, 0.75), (10, 0.5, 0.6, 0.7), (15, 0.5, 0.6, 0.65),
       (20, 0.4, 0.5, 0.65), (25, 0.35, 0.4, 0.65), (30, 0.35, 0.4, 0.6), (50, 0.35, 0.4, 0.55),
       (100, 0.35, 0.4, 0.55), (200, 0.35, 0.4, 0.5)]
NAMES = {"КсС_код": "$CE$4:$CE$15", "КсС_табл": "$CF$4:$CF$15", "КсС_n": "$CH$4:$CH$18",
         "КсС_шапка": "$CI$3:$CN$3", "КсС_знач": "$CI$4:$CN$18", "КсС_12эт": "$CI$20"}
NOTES = [
    "Как «группы по характеру нагрузки» в nanoCAD: код линии (CE) -> таблица Кс (CF) -> Кс по числу ЭП в щите.",
    "n — число линий «Сил» с той же таблицей в щите и секции (Руст > 0). Резервные — группа «Нет» (не считаются).",
    "Ручной Кс — «Исходные данные» BA, он главнее таблицы. Код не в списке — Кс = 1.",
    "Лифты — СП 256 (СП 31-110 табл. 6.4); двигатели сантехники и вентиляции — СП 31-110 табл. 6.9.",
    "Между точками таблицы — значение для меньшего n (в запас). Значения сверить с действующей редакцией СП 256.",
    "Своя таблица (пищеблок и т. п.): заголовок в CL3, значения CL4:CL18, имя таблицы — в CF у кода.",
    "Нижестоящий щит приходит готовым Рр сил: его ЭП в n вышестоящего не входят (в запас).",
]
HDR = {"BB": "Таблица Кс сил («Справочная» CF)",
       "BC": "Число ЭП n (щит, секция, таблица)",
       "BD": "Кс по таблице (СП 256)"}
F_BB = ('=IF(OR($A3<>"Сил",N($AR3)=0),"",IF(ISNUMBER(INDEX(\'Исходные данные\'!$BA$1:$BA$1000,$AR3)),"ручной",'
        'IFERROR(INDEX(КсС_табл,MATCH(INDEX(\'Исходные данные\'!$AH$1:$AH$1000,$AR3)&"",КсС_код,0))&"","")))')
F_BC = ('=IF(OR($BB3="",$BB3="ручной"),"",COUNTIFS($C$3:$C$1000,$C3,$AJ$3:$AJ$1000,$AJ3,'
        '$BB$3:$BB$1000,$BB3,$H$3:$H$1000,">0"))')
F_BD = ('=IF(N($BC3)=0,"",IFERROR(INDEX(КсС_знач,MATCH($BC3,КсС_n,1),MATCH(IF($BB3="Лифт",'
        'IF(LOWER(TRIM(КсС_12эт&""))="да","Лифт 12+","Лифт"),$BB3),КсС_шапка,0)),""))')
OLD_B = '),1),1),IF(RC[-1]="Пост"'                # Сил: ISNUMBER(BA) -> BA, иначе 1 (R1C1 от столбца B)
NEW_B = '),IF(ISNUMBER(RC[54]),RC[54],1)),1),IF(RC[-1]="Пост"'   # иначе BD (B + 54), иначе 1
BA_HDR = "Коэффициент спроса Кс (для Сил; ручной, пусто = по таблице «Справочная» CE:CN)"


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def unprotect(ws):
    if not ws.ProtectContents:
        return None
    p = ws.Protection
    opts = (True, True, True, False, p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
            p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks, p.AllowDeletingColumns,
            p.AllowDeletingRows, p.AllowSorting, p.AllowFiltering, p.AllowUsingPivotTables)
    ws.Unprotect(PWD)
    return opts


def runs(ws, col, r0, r1):
    fs = [x[0] for x in ws.Range(f"{col}{r0}:{col}{r1}").Formula2R1C1]
    out, start = [], r0
    for i in range(1, len(fs) + 1):
        if i == len(fs) or fs[i] != fs[i - 1]:
            out.append((start, r0 + i - 1, fs[i - 1]))
            start = r0 + i
    return out


def fmt(ws, src, dst):
    ws.Range(src).Copy(ws.Range(dst))       # без буфера обмена (PasteSpecial падал); значения — стираем
    ws.Range(dst).ClearContents()


def spr(wb, log):
    ws = wb.Worksheets(SPR)
    t = str(ws.Range("CE2").Value or "").strip()
    if t == TITLE:
        return
    used = [c for c in ("CE", "CF", "CH", "CI", "CJ", "CK") if ws.Range(f"{c}1:{c}30").Find("*") is not None]
    if used:
        raise SystemExit(f"«{SPR}» {','.join(used)} заняты — книга не изменена")
    opts = unprotect(ws)
    fmt(ws, "CC1", "CE1")
    ws.Range("CE1").Value = "Кс силовых по числу ЭП в щите. Жёлтые — можно менять."
    fmt(ws, "CC2", "CE2")
    ws.Range("CE2").Value = TITLE
    fmt(ws, "BP3", "CE3:CF3")
    fmt(ws, "BP3", "CH3:CN3")
    ws.Range("CE3:CF3").Value = (("Код (из P)", "Таблица Кс"),)
    ws.Range("CH3:CK3").Value = (("Число ЭП n", "Лифт", "Лифт 12+", "Двиг"),)
    fmt(ws, "BQ4", f"CE4:CF{3 + MAP_ROWS}")
    for i, (c, tb) in enumerate(MAP):
        ws.Range(f"CE{4 + i}:CF{4 + i}").Value = ((c, tb),)
    fmt(ws, "BP4", f"CH4:CH{3 + len(TAB)}")
    fmt(ws, "BQ4", f"CI4:CN{3 + len(TAB)}")
    for i, row in enumerate(TAB):
        ws.Range(f"CH{4 + i}:CK{4 + i}").Value = (row,)
    fmt(ws, "BP14", "CH20")
    fmt(ws, "BQ14", "CI20")
    ws.Range("CH20").Value = "Этажей 12 и более"
    ws.Range("CI20").Value = "нет"
    fmt(ws, "BR14", f"CH22:CH{21 + len(NOTES)}")
    for i, s in enumerate(NOTES):
        ws.Range(f"CH{22 + i}").Value = s
    for c, w in (("CE", 10), ("CF", 11), ("CG", 3), ("CH", 11), ("CI", 8), ("CJ", 9), ("CK", 8),
                 ("CL", 8), ("CM", 8), ("CN", 8)):
        ws.Columns(c).ColumnWidth = w
    s = str(ws.Range("BR18").Value or "")
    if "Сил" in str(ws.Range("BP18").Value or "") and "CE" not in s:
        ws.Range("BR18").Value = ("Силовая: вентиляция, насосы, лифты, пищеблок… Кс — по числу ЭП в щите "
                                  "(таблицы CE:CN этой страницы); ручной — «Исходные данные» BA.")
    for n, ref in NAMES.items():
        try:
            wb.Names(n).Delete()
        except Exception:
            pass
        wb.Names.Add(n, f"='{SPR}'!{ref}")
    if opts:
        ws.Protect(PWD, *opts)
    log.append(f"«{SPR}» CE:CN — таблицы Кс силовых, имена КсС_*")


def nag(wb, log):
    ws = wb.Worksheets(NAG)
    opts = unprotect(ws)
    last = ws.Cells(ws.Rows.Count, "B").End(-4162).Row
    for col, f in (("BB", F_BB), ("BC", F_BC), ("BD", F_BD)):
        h = ws.Range(f"{col}2")
        if str(h.Value or "").strip() not in ("", HDR[col]):
            raise SystemExit(f"«{NAG}» {col}2 занят: «{h.Value}» — книга не изменена")
        if h.Value != HDR[col]:
            fmt(ws, "BA2", f"{col}2")
            h.Value = HDR[col]
            fmt(ws, "AZ3", f"{col}3:{col}{last}")
            ws.Columns(col).ColumnWidth = 11
        if ws.Range(f"{col}3").Formula2 != f or ws.Range(f"{col}{last}").Formula2 == "":
            ws.Range(f"{col}3:{col}{last}").Formula2 = f
            log.append(f"«{NAG}» {col}3:{col}{last} — {HDR[col]}")
    n = 0
    for a, b, f in runs(ws, "B", 3, last):
        if isinstance(f, str) and OLD_B in f and "RC[54]" not in f:
            ws.Range(f"B{a}:B{b}").Formula2R1C1 = f.replace(OLD_B, NEW_B, 1)
            n += b - a + 1
    if n:
        log.append(f"«{NAG}» B — Кс «Сил» по таблице (BD), если BA пусто: {n} строк")
    elif not any("RC[54]" in str(f) for _, _, f in runs(ws, "B", 3, last)):
        log.append(f"«{NAG}» B — формула не узнана, Кс по таблице не подключён")
    if opts:
        ws.Protect(PWD, *opts)


def src(wb, log):
    ws = wb.Worksheets(SRC)
    h = ws.Range("BA2")
    if "Кс" in str(h.Value or "") and h.Value != BA_HDR:
        opts = unprotect(ws)
        h.Value = BA_HDR
        if opts:
            ws.Protect(PWD, *opts)
        log.append(f"«{SRC}» BA2 — заголовок: ручной Кс, пусто = по таблице")


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        app.Calculation = -4135
        spr(wb, log)
        nag(wb, log)
        src(wb, log)
        app.Calculation = -4105
        app.CalculateFull()
        if log:
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
