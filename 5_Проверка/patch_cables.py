# -*- coding: utf-8 -*-
"""Кабель в столбце Q «Однолинейки»: ввод — по питающей линии, слаботочные — без ΔU.

    py patch_cables.py <книга.xlsx> [...]      — книга закрыта; повторный запуск ничего не портит

1. Ввод (код «Ввод», строка Ввод.n): Q = кабель питающей линии, её длина и ΔU от ВРУ до шин щита
   («ВВГнг(А)-LS (5х6) L=20-0,77%»), как у отходящих линий. Источник — «Сеть_щита»:
     ввод 1 — раздел 1 (H:L, ставит «Связать с вышестоящим»), пусто — раздел 2 вручную (C18 марка, C19 сечение,
     C20 жил, C22 длина: щит без вышестоящего шкафа в этой системе — ВРУ, щит от ТП, сторонний щит);
     вводы 2, 3 — «Вводы 2 и 3» (AN кабель, AO длина); ΔU — «Сеть_щита» AH по секции (ключ «щит|n»), иначе C33.
   Нет кабеля в «Сети_щита» — Q как раньше (строка «Ввод.n» в «Исходных данных»).
2. «Исходные данные» AO у слаботочной линии (CB = 1): кабель и длина без ΔU («L=24»). Раньше ΔU = #N/A
   (сечение 0,5 мм² нет в таблице сопротивлений) давало #N/A во всём кабеле — нет клеммы, линия не рисовалась.
3. «Однолинейка» AF / AG: слаботочная строка без своего автомата (BZ = 1, E пусто) не продолжает группу силовой
   строки выше (иначе датчики «висят» на автомате обогрева). Продолжает — после слаботочной строки или строки
   с промежуточным устройством в I:K (БП DALI, PS-KNX). На схеме — клеммы без PE, без питания сверху (к контроллеру).
   AE: «нет мощности» у слаботочной не пишется. Признак BZ ставит patch_lowcurrent.py (запускать раньше).
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import signal
import sys

PWD = "sx"
ODN, SRC, NET = "Однолинейка", "Исходные данные", "Сеть_щита"
MARK = "ПЩ_dU"                                    # метка новой формулы Q (имя есть и в A1, и в R1C1)
Q_OLD = '=IF(RC[1]="","",IFERROR(SUBSTITUTE(Разбивка_по_щитам!RC[-1],CHAR(10)," "),""))'
Q_NEW = (
    '=IF($R3="","",LET(old_,IFERROR(SUBSTITUTE(Разбивка_по_щитам!P3,CHAR(10)," "),""),'
    'IF($AL3<>"Ввод",old_,IFERROR(LET(n_,IFERROR(--MID($R3,6,2),1),k_,MATCH($A3,ПЩ_щиты,0),r_,(k_-1)*2+n_-1,'
    'cab_,IF(n_<=1,IF(INDEX(Сеть_щита!$J$6:$J$8,k_)&""<>"",INDEX(Сеть_щита!$J$6:$J$8,k_)&"",'
    'IF(INDEX(Сеть_щита!$C$18:$E$18,1,k_)&""<>"",INDEX(Сеть_щита!$C$18:$E$18,1,k_)&IF(N(INDEX(Сеть_щита!$C$19:$E$19,1,k_))>0,'
    '" ("&IF(N(INDEX(Сеть_щита!$C$20:$E$20,1,k_))=0,5,INDEX(Сеть_щита!$C$20:$E$20,1,k_))&"х"'
    '&SUBSTITUTE(INDEX(Сеть_щита!$C$19:$E$19,1,k_)&"",".",",")&")",""),"")),'
    'INDEX(Сеть_щита!$AN$6:$AN$11,r_)&""),'
    'l_,IF(n_<=1,IF(N(INDEX(Сеть_щита!$K$6:$K$8,k_))>0,N(INDEX(Сеть_щита!$K$6:$K$8,k_)),'
    'N(INDEX(Сеть_щита!$C$22:$E$22,1,k_))),N(INDEX(Сеть_щита!$AO$6:$AO$11,r_))),'
    'du_,IFERROR(VLOOKUP($A3&"|"&n_,Сеть_щита!$AE$6:$AH$40,4,FALSE),INDEX(ПЩ_dU,1,k_)),'
    'IF(cab_="",old_,cab_&" L="&l_&"-"&FIXED(N(du_),2)&"%")),old_))))')
AO_OLD = '=IF(RC[-19]=0,RC[-11],RC[-10]&"\nL="&RC[-28]&"-"&RC[-13]&"%")'
AO_NEW = '=IF(AND(N(RC[39])=1,RC[-19]<>0),RC[-10]&"\nL="&RC[-28],IF(RC[-19]=0,RC[-11],RC[-10]&"\nL="&RC[-28]&"-"&RC[-13]&"%"))'


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


def odn(wb, log):
    names = [n.Name for n in wb.Names]
    if NET not in [s.Name for s in wb.Worksheets] or "ПЩ_щиты" not in names or "ПЩ_dU" not in names:
        log.append(f"«{NET}» без разделов (merge_network.py) — ввод не изменён")
        return
    ws = wb.Worksheets(ODN)
    opts = unprotect(ws)
    last = max(ws.Cells(ws.Rows.Count, c).End(-4162).Row for c in ("A", "H", "AF", "BH"))
    cur = runs(ws, "Q", 3, last)
    if any(isinstance(f, str) and MARK in f for _, _, f in cur):
        n = bad = 0
    else:
        # A1-формула от строки 3 на весь столбец; ячейки не со старой формулой — вернуть как были
        keep = [(a, b, f) for a, b, f in cur if f != Q_OLD]
        ws.Range(f"Q3:Q{last}").Formula2 = Q_NEW
        for a, b, f in keep:
            ws.Range(f"Q{a}:Q{b}").Formula2R1C1 = f
        n = sum(b - a + 1 for a, b, f in cur if f == Q_OLD)
        bad = sum(b - a + 1 for a, b, f in keep if isinstance(f, str) and f.startswith("="))
    if n:
        log.append(f"«{ODN}» Q — кабель ввода из «{NET}» (длина, ΔU до шин): {n} строк")
    if bad:
        log.append(f"«{ODN}» Q — формула не узнана в {bad} строках, не изменена")
    if opts:
        ws.Protect(PWD, *opts)


# AF (32) / AG (33): BZ = 78, I:K = 9..11
GRP = {
    "AF": ("N(RC[46])=1", 'RC[-7]="вручную",',
           'RC[-7]="вручную",AND(N(RC[46])=1,N(R[-1]C[46])=0,R[-1]C[-23]&R[-1]C[-22]&R[-1]C[-21]=""),'),
    "AG": ("N(RC[45])=1", 'RC[-8]="вручную",',
           'RC[-8]="вручную",AND(N(RC[45])=1,N(R[-1]C[45])=0,R[-1]C[-24]&R[-1]C[-23]&R[-1]C[-22]=""),'),
}


def low(wb, log):
    ws = wb.Worksheets(ODN)
    if str(ws.Range("BZ2").Value or "").strip()[:7] != "слабото":
        return                                        # нет признака слаботочной (patch_lowcurrent.py)
    opts = unprotect(ws)
    last = max(ws.Cells(ws.Rows.Count, c).End(-4162).Row for c in ("A", "H", "AF", "BH"))
    # AE: «нет мощности» не у слаботочной (как в шаблоне: N(BZ)<>1)
    n = 0
    for a, b, f in runs(ws, "AE", 3, last):
        if isinstance(f, str) and '"нет мощности"' in f and "N(RC[47])<>1" not in f:
            k = f.find('"нет мощности"')
            j = f.rfind("N(RC[-11])=0", 0, k)
            if j >= 0:
                j += len("N(RC[-11])=0")
                ws.Range(f"AE{a}:AE{b}").Formula2R1C1 = f[:j] + ",N(RC[47])<>1" + f[j:]
                n += b - a + 1
    if n:
        log.append(f"«{ODN}» AE — «нет мощности» не у слаботочных: {n} строк")
    for col, (mark, old, new) in GRP.items():
        n = 0
        for a, b, f in runs(ws, col, 3, last):
            if isinstance(f, str) and old in f and mark not in f:
                ws.Range(f"{col}{a}:{col}{b}").Formula2R1C1 = f.replace(old, new, 1)
                n += b - a + 1
        if n:
            log.append(f"«{ODN}» {col} — слаботочная без автомата не в силовой группе: {n} строк")
    if opts:
        ws.Protect(PWD, *opts)


def src(wb, log):
    ws = wb.Worksheets(SRC)
    opts = unprotect(ws)
    n = 0
    for a, b, f in runs(ws, "AO", 3, 1000):
        if isinstance(f, str) and f.replace("\r", "") == AO_OLD:
            ws.Range(f"AO{a}:AO{b}").Formula2R1C1 = AO_NEW
            n += b - a + 1
    if n:
        log.append(f"«{SRC}» AO — слаботочные (CB = 1): кабель без ΔU: {n} строк")
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
        low(wb, log)
        app.Calculation = -4105
        app.CalculateFull()
        if any("строк" in s for s in log):
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
