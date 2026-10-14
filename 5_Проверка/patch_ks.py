# -*- coding: utf-8 -*-
"""Кс силовых линий — сам, по числу электроприёмников в щите (как «группы по характеру» в nanoCAD BIM Электро).

    py patch_ks.py <книга.xlsx> [...]      — книга закрыта; повторный запуск ничего не портит

Было: Кс линии «Сил» = «Исходные данные» BA (руками), пусто = 1.
Стало: Кс линии «Сил» = BA (ручной, главнее) -> иначе по таблице СП 256 по числу ЭП -> иначе 1.
Таблицы сверены с текстом СП 256.1325800.2016 (ред. 01.03.2022), между точками — интерполяция (примечания к таблицам):
  «Справочная» CE:CF — код линии -> таблица Кс (ЛФ — Лифт; НС, ПВ, ПУ, В — Двиг). Жёлтые — можно менять.
  «Справочная» CH:CI — табл. 7.4, лифты в домах до 12 этажей (12 и более — не используем).
  «Справочная» CK:CW — табл. 7.5, сантехника и холодильные машины: строка по доле их Руст в Руст силовых щита
                       (≥85 / 75 / 50 / 25 / меньше %), столбец по n; строка 9 — двигатели > 30 кВт (значения в скобках).
  «Нагрузка_щитов» BB — таблица Кс линии, BC — n (линий «Сил» с той же таблицей в щите и секции, P > 0),
                   BD — Кс по таблице, BE — доля сантехники, %, BF — строка табл. 7.5; B (Кс линии) берёт BD, если BA пусто.
Первая версия блока (с «Лифт 12+» и «Этажей 12 и более») заменяется. Нижестоящие щиты приходят готовым Рр сил —
n через щиты не суммируется (в запас). Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import signal
import sys

PWD = "sx"
SPR = "Справочная"
NAG = "Нагрузка_щитов"
SRC = "Исходные данные"
TITLE = "Кс СИЛОВЫХ (СП 256 табл. 7.4, 7.5)"
OLD_TITLES = ("Кс СИЛОВЫХ (СП 256)",)

MAP = [("ЛФ", "Лифт"), ("НС", "Двиг"), ("ПВ", "Двиг"), ("ПУ", "Двиг"), ("В", "Двиг")]
MAP_ROWS = 12                                    # CE4:CF15
# табл. 7.4, до 12 этажей: 2–3 -> 0,8; 4–5 -> 0,7; 6 -> 0,65; 10 -> 0,5; 20 -> 0,4; 25 и св. -> 0,35 (1 лифт — 1)
LIFT = [(1, 1), (2, 0.8), (3, 0.8), (4, 0.7), (5, 0.7), (6, 0.65), (10, 0.5), (20, 0.4), (25, 0.35)]
# табл. 7.5: n по столбцам; n = 1 — 1; прочерки (n = 2, 3 при доле < 85 %) — как в строке ≥85 %
DN = (1, 2, 3, 5, 8, 10, 15, 20, 30, 50, 100, 200)
DROWS = [(0, (1, 1, 0.9, 0.6, 0.6, 0.55, 0.5, 0.5, 0.5, 0.45, 0.45, 0.4)),       # 24 и менее
         (25, (1, 1, 0.9, 0.65, 0.6, 0.6, 0.55, 0.5, 0.5, 0.5, 0.45, 0.45)),     # 49–25
         (50, (1, 1, 0.9, 0.7, 0.65, 0.65, 0.6, 0.6, 0.55, 0.5, 0.5, 0.45)),     # 74–50
         (75, (1, 1, 0.9, 0.75, 0.7, 0.65, 0.6, 0.6, 0.6, 0.55, 0.55, 0.5)),     # 84–75
         (85, (1, 1, 0.9, 0.8, 0.75, 0.7, 0.65, 0.65, 0.6, 0.55, 0.55, 0.5)),    # 100–85
         (">30 кВт", (1, 0.8, 0.75, 0.7, 0.75, 0.7, 0.65, 0.65, 0.6, 0.55, 0.55, 0.5))]   # в скобках, ≥85 %
NAMES = {"КсС_код": "$CE$4:$CE$15", "КсС_табл": "$CF$4:$CF$15", "КсЛ_n": "$CH$4:$CH$12", "КсЛ_к": "$CI$4:$CI$12",
         "КсД_n": "$CL$3:$CW$3", "КсД_доля": "$CK$4:$CK$8", "КсД_к": "$CL$4:$CW$9"}
OLD_NAMES = ("КсС_n", "КсС_шапка", "КсС_знач", "КсС_12эт")
NOTES = [
    "Как «группы по характеру нагрузки» в nanoCAD: код линии (CE) -> таблица Кс (CF) -> Кс по числу ЭП в щите.",
    "n — число линий «Сил» с той же таблицей в щите и секции (Руст > 0). Резервные — группа «Нет» (СП 256 п. 7.1.9).",
    "Ручной Кс — «Исходные данные» BA, он главнее таблицы. Код не в списке — Кс = 1.",
    "Лифты — СП 256 табл. 7.4 (дома до 12 этажей). Сантехника, вентиляция, насосы — табл. 7.5.",
    "Табл. 7.5: строка — доля Руст «Двиг» в Руст силовых щита, %; строка 9 — двигатели > 30 кВт при доле ≥ 85 %.",
    "Между точками — интерполяция (примечания к табл. 7.4, 7.5). n = 1 и прочерки СП — как в строке ≥ 85 %.",
    "Противопожарные ЭП по п. 7.1.9 в Рр не входят — сейчас входят (в запас). Нижестоящие щиты в n не входят.",
]
HDR = {"BB": "Таблица Кс сил («Справочная» CF)",
       "BC": "Число ЭП n (щит, секция, таблица)",
       "BD": "Кс по таблице (СП 256 табл. 7.4 / 7.5)",
       "BE": "Доля сантехники в Руст сил, %",
       "BF": "Строка табл. 7.5 («Справочная» CK)"}
OLD_HDR = {"BD": "Кс по таблице (СП 256)"}


def interp(nv, val, n):
    """Линейная интерполяция по точкам nv (имя диапазона); val(k) — значение k-й точки."""
    m = f"MATCH({n},{nv},1)"
    return (f"IF({n}>=MAX({nv}),{val(f'COUNT({nv})')},{val(m)}+({n}-INDEX({nv},{m}))*"
            f"({val(m + '+1')}-{val(m)})/(INDEX({nv},{m}+1)-INDEX({nv},{m})))")


F_BB = ('=IF(OR($A3<>"Сил",N($AR3)=0),"",IF(ISNUMBER(INDEX(\'Исходные данные\'!$BA$1:$BA$1000,$AR3)),"ручной",'
        'IFERROR(INDEX(КсС_табл,MATCH(INDEX(\'Исходные данные\'!$AH$1:$AH$1000,$AR3)&"",КсС_код,0))&"","")))')
F_BC = ('=IF(OR($BB3="",$BB3="ручной"),"",COUNTIFS($C$3:$C$1000,$C3,$AJ$3:$AJ$1000,$AJ3,'
        '$BB$3:$BB$1000,$BB3,$H$3:$H$1000,">0"))')
F_BE = ('=IF($BB3<>"Двиг","",IFERROR(ROUND(100*SUMIFS($H$3:$H$1000,$C$3:$C$1000,$C3,$AJ$3:$AJ$1000,$AJ3,'
        '$BB$3:$BB$1000,"Двиг")/SUMIFS($H$3:$H$1000,$C$3:$C$1000,$C3,$AJ$3:$AJ$1000,$AJ3,$A$3:$A$1000,"Сил"),1),""))')
F_BF = ('=IF($BE3="","",IF(AND($BE3>=85,MAXIFS($H$3:$H$1000,$C$3:$C$1000,$C3,$AJ$3:$AJ$1000,$AJ3,'
        '$BB$3:$BB$1000,"Двиг")>30),6,MATCH($BE3,КсД_доля,1)))')
F_BD = ('=IF(N($BC3)=0,"",IFERROR(ROUND(IF($BB3="Лифт",'
        + interp("КсЛ_n", lambda k: f"INDEX(КсЛ_к,{k})", "$BC3")
        + ',IF(AND($BB3="Двиг",ISNUMBER($BF3)),'
        + interp("КсД_n", lambda k: f"INDEX(КсД_к,$BF3,{k})", "$BC3")
        + ',"")),2),""))')
OLD_B = '),1),1),IF(RC[-1]="Пост"'                # Сил: ISNUMBER(BA) -> BA, иначе 1 (R1C1 от столбца B)
NEW_B = '),IF(ISNUMBER(RC[54]),RC[54],1)),1),IF(RC[-1]="Пост"'   # иначе BD (B + 54), иначе 1
BA_HDR = "Коэффициент спроса Кс (для Сил; ручной, пусто = по таблице «Справочная» CE:CW)"


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
    opts = unprotect(ws)
    if t in OLD_TITLES:
        ws.Range("CE1:CW40").Clear()                 # первая версия блока — заменить целиком
        log.append(f"«{SPR}» CE:CN — первая версия таблиц Кс снята")
    else:
        used = [c for c in ("CE", "CF", "CH", "CI", "CK", "CL", "CW")
                if ws.Range(f"{c}1:{c}40").Find("*") is not None]
        if used:
            raise SystemExit(f"«{SPR}» {','.join(used)} заняты — книга не изменена")
    fmt(ws, "CC1", "CE1")
    ws.Range("CE1").Value = "Кс силовых по числу ЭП в щите (СП 256 табл. 7.4, 7.5). Жёлтые — можно менять."
    fmt(ws, "CC2", "CE2")
    ws.Range("CE2").Value = TITLE
    fmt(ws, "BP3", "CE3:CF3")
    ws.Range("CE3:CF3").Value = (("Код (из P)", "Таблица Кс"),)
    fmt(ws, "BQ4", f"CE4:CF{3 + MAP_ROWS}")
    for i, (c, tb) in enumerate(MAP):
        ws.Range(f"CE{4 + i}:CF{4 + i}").Value = ((c, tb),)
    fmt(ws, "BP3", "CH3:CI3")
    ws.Range("CH3:CI3").Value = (("Лифтов n", "Кс.л (табл. 7.4)"),)
    fmt(ws, "BP4", f"CH4:CH{3 + len(LIFT)}")
    fmt(ws, "BQ4", f"CI4:CI{3 + len(LIFT)}")
    ws.Range(f"CH4:CI{3 + len(LIFT)}").Value = tuple(LIFT)
    fmt(ws, "BP3", "CK3:CW3")
    ws.Range("CK3").Value = "Табл. 7.5: доля от, % \\ n"
    ws.Range("CL3:CW3").Value = (DN,)
    fmt(ws, "BP4", f"CK4:CK{3 + len(DROWS)}")
    fmt(ws, "BQ4", f"CL4:CW{3 + len(DROWS)}")
    ws.Range(f"CK4:CW{3 + len(DROWS)}").Value = tuple((a,) + v for a, v in DROWS)
    fmt(ws, "BR14", f"CE18:CE{17 + len(NOTES)}")
    for i, s in enumerate(NOTES):
        ws.Range(f"CE{18 + i}").Value = s
    for c, w in (("CE", 10), ("CF", 11), ("CG", 3), ("CH", 9), ("CI", 10), ("CJ", 3), ("CK", 12)):
        ws.Columns(c).ColumnWidth = w
    ws.Range("CL:CW").ColumnWidth = 6
    s = str(ws.Range("BR18").Value or "")
    if "Сил" in str(ws.Range("BP18").Value or "") and "CE:CW" not in s:
        ws.Range("BR18").Value = ("Силовая: вентиляция, насосы, лифты, пищеблок… Кс — по числу ЭП в щите "
                                  "(СП 256 табл. 7.4, 7.5 — CE:CW этой страницы); ручной — «Исходные данные» BA.")
    for n in OLD_NAMES + tuple(NAMES):
        try:
            wb.Names(n).Delete()
        except Exception:
            pass
    for n, ref in NAMES.items():
        wb.Names.Add(n, f"='{SPR}'!{ref}")
    if opts:
        ws.Protect(PWD, *opts)
    log.append(f"«{SPR}» CE:CW — таблицы Кс силовых (СП 256 табл. 7.4, 7.5), имена КсС_*, КсЛ_*, КсД_*")


def nag(wb, log):
    ws = wb.Worksheets(NAG)
    opts = unprotect(ws)
    last = ws.Cells(ws.Rows.Count, "B").End(-4162).Row
    for col, f in (("BB", F_BB), ("BC", F_BC), ("BE", F_BE), ("BF", F_BF), ("BD", F_BD)):
        h = ws.Range(f"{col}2")
        if str(h.Value or "").strip() not in ("", HDR[col], OLD_HDR.get(col)):
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
