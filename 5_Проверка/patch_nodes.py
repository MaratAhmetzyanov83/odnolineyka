# -*- coding: utf-8 -*-
"""Узлы (каскад, многоуровневая однолинейка): аппарат над группой автоматов — УЗО, АВДТ, автомат, рубильник.

    py patch_nodes.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит

Как пользоваться («Однолинейка»):
  * Узел — своя строка-место без линии: CL = тип узла (УЗО / АВДТ / Автомат / Рубильник), E = «авто» или текст
    аппарата («УЗО 2P 40А 30мА»). Имя узла — само, в CM: У1, У2…
  * Линии под узлом: D = имя узла (У1), у каждой свой E («авто»). Ставить правее узла (строки ниже).
  * Узел под узлом: в строке узла D = имя верхнего узла. Уровней — сколько нужно.
  * Узел на второй шине (ИБП, ДГУ): в строке узла D = имя шины; его линии наследуют шину (категории, режимы).
Что считается:
  * CN — Iр узла: сумма Iр линий под ним (все уровни) по наиболее нагруженной фазе, без Кс (в запас, как узел КМ).
    Номинал «авто» — по Iр узла (Z, AA, AO, AQ, H); полюса — по числу фаз линий под узлом (AR).
  * УЗО / АВДТ (30 мА): автоматам под ними своё 30 мА не нужно — QF, не QFD; нет замечания «розетка без 30 мА».
    Задать 30 мА автомату под УЗО всё равно можно: F = «да».
  * Обозначение: УЗО — QD, АВДТ — QFD, автомат — QF, рубильник — QS (как ввод). Текст H: «УЗО 2P 40А 30мА», «QS 3P 63А».
  * CR — проверка узла (идёт в статус AE): нет линий, линия левее узла, неселективно (автомат, АВДТ: In узла
    < 1,6 × наибольшего автомата под ним), УЗО: сумма In автоматов под ним > In УЗО — нужен автомат перед УЗО.
  * CS — шина линии с учётом узлов (корень цепочки): «Исходные данные» BY берёт её — категории и режимы видят
    узел на шине ИБП как ИБП. Линии узла КМ (KM-1) теперь на основной шине (раньше — «шина KM-1» типа «Ввод»).
Узел КМ (CE:CI, patch_km.py) не меняется. Служебные CO (путь узлов), CP (ключ узла), CQ (30 мА выше).
SXDRAW до v3.25 рисует узел как автомат на шине, его линии — на второй шине с именем узла.
Свой скрытый экземпляр Excel; закрывается только он. COM — только позиционные аргументы.
"""
import os
import signal
import sys

PWD = "sx"
ODN, SRC = "Однолинейка", "Исходные данные"
TYPES = ("УЗО", "АВДТ", "Автомат", "Рубильник")
GROUP = "УЗЕЛ (каскад)"
MARK = "RC90"                     # ссылка на CL (тип узла) — признак, что формула уже доработана

# новые столбцы: (заголовок строки 2, формула для строки 3 в A1; None — ввод)
COLS = {
    "CL": ("Узел: тип\n(УЗО / АВДТ /\nАвтомат / Рубильник)", None),
    "CM": ("Узел: имя\n(в D линий)",
           '=IF($CL3="","","У"&COUNTIFS($A$3:$A3,$A3,$CL$3:$CL3,"?*"))'),
    "CN": ("Iр узла, А\n(фаза, без Кс)",
           '=IF($CM3="","",ROUND(MAX('
           'SUMIFS($U$3:$U$1000,$A$3:$A$1000,$A3,$CO$3:$CO$1000,"*/"&$CM3&"/*",$V$3:$V$1000,"*L1*",$AN$3:$AN$1000,0),'
           'SUMIFS($U$3:$U$1000,$A$3:$A$1000,$A3,$CO$3:$CO$1000,"*/"&$CM3&"/*",$V$3:$V$1000,"*L2*",$AN$3:$AN$1000,0),'
           'SUMIFS($U$3:$U$1000,$A$3:$A$1000,$A3,$CO$3:$CO$1000,"*/"&$CM3&"/*",$V$3:$V$1000,"*L3*",$AN$3:$AN$1000,0)),2))'),
    # путь: цепочка узлов / шин над строкой («/У1/У2/»); родитель — только выше (без циклической ссылки)
    "CO": ("служебн.: путь\n(узлы над строкой)",
           '=IF($D3<>"",IFERROR(INDEX($CO$2:$CO2,MATCH($A3&"|"&$D3,$CP$2:$CP2,0)),"/")&$D3&"/",'
           'IF(AND($E3="",$AF3<>"",$AF3=$AF2,$A3=$A2),$CO2,"/"))'),
    "CP": ("служебн.: ключ узла",
           '=IF($CM3<>"",$A3&"|"&$CM3,IF($CE3<>"",$A3&"|"&$CE3,""))'),
    "CQ": ("служебн.: 30 мА выше\n(УЗО / АВДТ)",
           '=IF($D3<>"",IFERROR(MAX(N(INDEX($CQ$2:$CQ2,MATCH($A3&"|"&$D3,$CP$2:$CP2,0))),'
           'IF(AND(OR(INDEX($CL$2:$CL2,MATCH($A3&"|"&$D3,$CP$2:$CP2,0))="УЗО",'
           'INDEX($CL$2:$CL2,MATCH($A3&"|"&$D3,$CP$2:$CP2,0))="АВДТ"),'
           'N(INDEX($AS$2:$AS2,MATCH($A3&"|"&$D3,$CP$2:$CP2,0)))=1),1,0)),0),'
           'IF(AND($E3="",$AF3<>"",$AF3=$AF2,$A3=$A2),N($CQ2),0))'),
    "CR": ("Узел: проверка",
           '=IF($CM3="","",'
           'IF(COUNTIFS($A$3:$A$1000,$A3,$D$3:$D$1000,$CM3)=0,"узел "&$CM3&": нет линий — в D линий впишите "&$CM3,'
           'IF(COUNTIFS($A$3:$A$1000,$A3,$D$3:$D$1000,$CM3,$B$3:$B$1000,"<"&$B3)>0,"узел "&$CM3&": линия левее узла",'
           'IF($E3="","узел "&$CM3&": нет аппарата — E «авто» или текст",'
           'IF(AND(OR($CL3="Автомат",$CL3="АВДТ"),N($AU3)>0,'
           'N($AU3)<1.6*MAXIFS($AU$3:$AU$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CM3)),'
           '"узел "&$CM3&": неселективно — In "&$AU3&" А < 1,6 × "&MAXIFS($AU$3:$AU$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CM3)&" А",'
           'IF(AND($CL3="УЗО",N($AU3)>0,SUMIFS($AU$3:$AU$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CM3)>$AU3),'
           '"узел "&$CM3&": УЗО "&$AU3&" А, сумма автоматов под ним "&SUMIFS($AU$3:$AU$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CM3)'
           '&" А — нужен автомат перед УЗО не больше "&$AU3&" А",'
           '"ok"))))))'),
    "CS": ("служебн.: шина\n(с учётом узлов)",
           '=IF($CO3="/","",IF(COUNTIF($CP$3:$CP$1000,$A3&"|"&MID($CO3,2,FIND("/",$CO3,2)-2))>0,"",'
           'MID($CO3,2,FIND("/",$CO3,2)-2)))'),
}

# ---------------------------------------------------------------- доработка существующих формул (R1C1)
# столбцы: E 5, F 6, H 8, AF 32, AG 33, AK 37, AQ 43, AS 45, AT 46, CL 90, CM 91, CN 92, CO 93, CQ 95, CR 96, CS 97
NODE = 'AND(RC90<>"",RC5<>"")'                         # строка-узел с аппаратом


def wrap(f: str, head: str) -> str:
    """=X -> =IF(<head>,X): head — «условие,значение»; закрывающая скобка — в конце."""
    return "=IF(" + head + "," + f[1:] + ")"


def fix_z(f):          # ток группы: у узла — Iр узла
    return wrap(f, f"{NODE},N(RC92)")


def fix_ar(f):         # фаз группы: у узла — наибольшее число фаз линий под ним
    return wrap(f, f'{NODE},MAX(1,MAXIFS(R3C37:R1000C37,R3C1:R1000C1,RC1,R3C93:R1000C93,"*/"&RC91&"/*"))')


def fix_as(f):         # 30 мА: у узла — по типу; под УЗО / АВДТ — не нужно (кроме F = «да»)
    f = wrap(f, 'AND(N(RC95)=1,RC6<>"да",OR(RC5="авто",RC5="+")),0')
    return wrap(f, f'{NODE},IF(RC6="нет",0,IF(OR(RC90="УЗО",RC90="АВДТ"),1,IF(RC6="да",1,0)))')


def fix_g(f):          # обозначение на схеме: у узла нет линии и клеммы — показать всё равно
    return wrap(f, f'{NODE},IF(LEFT(RC8,3)="QS ",SUBSTITUTE(RC33,"QF","QS"),RC33)')


def fix_aj(f):         # обозначение: УЗО — QD
    return wrap(f, 'AND(RC90="УЗО",RC5<>""),"QD"&RC32')


def fix_h(f):          # текст аппарата узла «авто»: УЗО и рубильник; автомат и АВДТ — как обычно
    return wrap(f, f'AND({NODE},OR(RC5="авто",RC5="+"),OR(RC90="УЗО",RC90="Рубильник")),'
                   'IF(RC90="УЗО","УЗО "&RC46&" "&RC43&"А"&IF(N(RC45)=1," 30мА",""),"QS "&RC46&" "&RC43&"А")')


def fix_au(f):         # номинал из текста «УЗО 2P 40А …» / «QS 3P 63А» (в них нет «C»)
    num = ('--TRIM(RIGHT(SUBSTITUTE(LEFT(SUBSTITUTE(RC8,"A","А"),FIND("А",SUBSTITUTE(RC8,"A","А"))-1),'
           '" ",REPT(" ",50)),50))')
    return wrap(f, f'AND(RC90<>"",OR(LEFT(RC8,4)="УЗО ",LEFT(RC8,3)="QS ")),IFERROR({num},RC43)')


def fix_ae(f):         # статус: проверка узла; «розетка без 30 мА» — не под УЗО
    k = f.find('TEXTJOIN("; ",TRUE,')
    if k < 0:
        return None
    k += len('TEXTJOIN("; ",TRUE,')
    f = f[:k] + 'IF(OR(RC96="",LEFT(RC96,2)="ok"),"",RC96),' + f[k:]
    j = f.find('),"розетка без 30 мА"')
    if j > 0:
        f = f[:j] + ",N(RC95)=0" + f[j:]
    return f


ODN_FIX = {"G": fix_g, "Z": fix_z, "AR": fix_ar, "AS": fix_as, "AJ": fix_aj, "H": fix_h, "AU": fix_au, "AE": fix_ae}
SRC_OLD, SRC_NEW = "Однолинейка!R3C4:R1000C4", "Однолинейка!R3C97:R1000C97"     # «Исходные данные» BY: D -> CS


def fix_by(f):
    return f.replace(SRC_OLD, SRC_NEW) if SRC_OLD in f else None


# ---------------------------------------------------------------- Excel
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


def patch_runs(ws, col, last, fn, marker, log, sheet):
    n = bad = 0
    for a, b, f in runs(ws, col, 3, last):
        if not isinstance(f, str) or not f.startswith("=") or marker in f:
            continue
        g = fn(f)
        if g is None:
            bad += b - a + 1
            continue
        ws.Range(f"{col}{a}:{col}{b}").Formula2R1C1 = g
        n += b - a + 1
    if n:
        log.append(f"«{sheet}» {col} — узлы: {n} строк")
    if bad:
        log.append(f"«{sheet}» {col} — формула не узнана в {bad} строках, узлы там не учтены")


def odn(wb, log):
    ws = wb.Worksheets(ODN)
    opts = unprotect(ws)
    last = max(ws.Cells(ws.Rows.Count, c).End(-4162).Row for c in ("A", "H", "AF", "BH"))
    for col, (hdr, f) in COLS.items():
        h = ws.Range(f"{col}2")
        if str(h.Value or "").strip() not in ("", hdr):
            raise SystemExit(f"«{ODN}» {col}2 занят: «{h.Value}» — книга не изменена")
        if h.Value != hdr:
            ws.Range("CI2").Copy(h)
            h.Value = hdr
            ws.Range("CI3").Copy(ws.Range(f"{col}3:{col}{last}"))
            ws.Range(f"{col}3:{col}{last}").ClearContents()
            ws.Columns(col).ColumnWidth = {"CL": 12, "CR": 40}.get(col, 10)
            log.append(f"«{ODN}» {col}2 — «{hdr.replace(chr(10), ' ')}»")
        if f and (ws.Range(f"{col}3").Formula2 != f or ws.Range(f"{col}{last}").Formula2 == ""):
            ws.Range(f"{col}3:{col}{last}").Formula2 = f
            log.append(f"«{ODN}» {col}3:{col}{last} — формула")
    if not str(ws.Range("CL1").Value or ""):
        ws.Range("CL1").Value = GROUP
    # CL — ввод: белый фон и список
    rng = ws.Range(f"CL3:CL{last}")
    rng.Interior.Color = 0xFFFFFF
    rng.Locked = False
    try:
        rng.Validation.Delete()
        rng.Validation.Add(3, 1, 1, ",".join(TYPES))         # xlValidateList, AlertStop, Between
    except Exception:  # noqa: BLE001 — список не обязателен
        pass
    for col, fn in ODN_FIX.items():
        patch_runs(ws, col, last, fn, MARK if col != "AE" else "RC96", log, ODN)
    if opts:
        ws.Protect(PWD, *opts)


def src(wb, log):
    ws = wb.Worksheets(SRC)
    opts = unprotect(ws)
    patch_runs(ws, "BY", 1000, fix_by, "C97", log, SRC)
    if opts:
        ws.Protect(PWD, *opts)


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        if wb.ReadOnly:                              # открыта у другого — не трогаем
            wb.Close(False)
            raise SystemExit(f"{os.path.basename(path)}: открыта у другого пользователя — пропущена")
        app.Calculation = -4135
        odn(wb, log)
        src(wb, log)
        app.Calculation = -4105
        app.CalculateFull()
        if log:
            wb.Save()
        wb.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
        try:
            os.kill(pid, signal.SIGTERM)
        except Exception:  # noqa: BLE001
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
