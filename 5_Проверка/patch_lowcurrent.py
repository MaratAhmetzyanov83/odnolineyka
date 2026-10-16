# -*- coding: utf-8 -*-
"""Слаботочные линии (DALI, KNX, датчики…) в таблице v3.12 — признак по коду.

    py patch_lowcurrent.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит

«Справочная» CC: список «Слаботочные коды» (имя Слаботочные_коды): DALI, KNX, WD, VK — дополняйте сами.
Код KNX добавляется в таблицу потребителей; у DALI / WD / VK / KNX — марка кабеля по коду, если пусто.
У слаботочной линии («Исходные данные»):
  ток T = 0, потери AA / AB = 0, сечение V = по коду (Q) — без подбора по току и потерям;
  в нагрузку щита не идёт («Нагрузка_щитов» H = 0).
«Однолинейка»: BZ «слаботочн.» (1 — да; SXDRAW рисует клемму без PE);
  статус: нет «нет мощности»; «слаботочная: автомат без промежуточного устройства», если E задан, а I:K пусто.
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import subprocess
import sys

PWD = "sx"
REF, SRC, NAG, ODN = "Справочная", "Исходные данные", "Нагрузка_щитов", "Однолинейка"
CODES = ["DALI", "KNX", "WD", "VK"]
# код: (столбец «Справочной» -> значение), пишется только в пустые ячейки
CODE_CABLE = {
    "DALI": {"U": "мс", "V": 2, "X": 1, "Y": 1.5},
    "KNX": {"U": "кн", "V": 4, "X": 1},
    "WD": {"U": "кс", "V": 4, "X": 1},
    "VK": {"U": "кс", "V": 4, "X": 1},
}
KNX_ROW = {"P": "KNX", "Q": "Шина KNX", "R": 1, "S": 1, "T": 1, "AZ": "Нет"}
SRC_HDR = "Слаботочная линия (1 — да, по коду)"
ODN_COL, ODN_HDR = "BZ", "слаботочн."
MSG = "слаботочная: автомат без промежуточного устройства"


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
    opts = dict(DrawingObjects=True, Contents=True, Scenarios=True,
                AllowFormattingCells=p.AllowFormattingCells, AllowFormattingColumns=p.AllowFormattingColumns,
                AllowFormattingRows=p.AllowFormattingRows, AllowSorting=p.AllowSorting,
                AllowFiltering=p.AllowFiltering, AllowUsingPivotTables=p.AllowUsingPivotTables)
    ws.Unprotect(PWD)
    return opts


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


TITLES = ("", "СЛАБОТОЧНЫЕ ЛИНИИ", "СЛАБОТОЧНЫЕ КОДЫ", "Слаботочные коды (без тока, нагрузки и своего автомата)")


def manual_codes(wb):
    """Коды строк, где признак слаботочной вписан вручную (1 без формулы) — чтобы формула его не сбросила."""
    ws = wb.Worksheets(SRC)
    hdr = [str(x or "").strip() for x in ws.Range("A2:EZ2").Value[0]]
    c = hdr.index(SRC_HDR) + 1 if SRC_HDR in hdr else 80
    out = []
    for f, code in zip(ws.Range(f"{col_letter(c)}3:{col_letter(c)}1000").Formula,
                       ws.Range("AH3:AH1000").Value):
        f, code = f[0], str(code[0] or "").strip()
        if str(f).strip() == "1" and code and code not in out:
            out.append(code)
    return out


def ref_sheet(wb, log, extra=()):
    ws = wb.Worksheets(REF)
    opts = unprotect(ws)
    if "Слаботочные_коды" not in [n.Name for n in wb.Names]:
        # шапку «СЛАБОТОЧНЫЕ КОДЫ» ставит tidy_book.py — она не мешает
        if (any(ws.Range(f"CC{r}").Formula not in ("", None) for r in range(4, 41))
                or any(str(ws.Range(f"CC{r}").Formula or "").strip() not in TITLES for r in (2, 3))):
            raise RuntimeError("«Справочная» CC2:CC40 занят — некуда записать «Слаботочные коды»")
        if not ws.Range("CC2").Formula:
            ws.Range("CC2").Value = "СЛАБОТОЧНЫЕ ЛИНИИ"
        if not ws.Range("CC3").Formula:
            ws.Range("CC3").Value = "Слаботочные коды (без тока, нагрузки и своего автомата)"
        for i, c in enumerate(CODES):
            ws.Range(f"CC{4 + i}").Value = c
        ws.Range("CC2:CC3").Font.Bold = True
        ws.Range("CC3").WrapText = True
        ws.Columns("CC").ColumnWidth = 22
        wb.Names.Add(Name="Слаботочные_коды", RefersTo="='Справочная'!$CC$4:$CC$40")
        log.append("список «Слаботочные коды» добавлен")
    have = [str(x[0] or "").strip() for x in ws.Range("CC4:CC40").Value]
    add = [c for c in extra if c not in have]
    for c in add:
        if "" not in have:
            break
        r = 4 + have.index("")
        ws.Range(f"CC{r}").Value = c
        have[r - 4] = c
    if add:
        log.append(f"«Слаботочные коды» + {', '.join(add)} (был ручной признак 1)")
    codes = [str(x[0] or "").strip() for x in ws.Range("P4:P51").Value]
    if "KNX" not in codes:
        r = 4 + codes.index("")
        for c, v in KNX_ROW.items():
            ws.Range(f"{c}{r}").Value = v
        codes[r - 4] = "KNX"
        log.append(f"код KNX — строка {r}")
    for code, vals in CODE_CABLE.items():
        if code in codes:
            r = 4 + codes.index(code)
            for c, v in vals.items():
                if ws.Range(f"{c}{r}").Formula in ("", None):
                    ws.Range(f"{c}{r}").Value = v
    if opts:
        ws.Protect(PWD, **opts)


def src_sheet(wb, log):
    ws = wb.Worksheets(SRC)
    opts = unprotect(ws)
    hdr = [str(x or "").strip() for x in ws.Range("A2:EZ2").Value[0]]
    if SRC_HDR in hdr:
        c = hdr.index(SRC_HDR) + 1
    else:
        c = 80   # CB
        while hdr[c - 1]:
            c += 1
        ws.Cells(2, c).Value = SRC_HDR
        ws.Cells(2, c).WrapText = True
        log.append(f"«Исходные данные» {col_letter(c)}: признак слаботочной линии")
    ws.Range(f"{col_letter(c)}3:{col_letter(c)}1000").FormulaR1C1 = (
        '=IF(RC4="","",IF(ISNUMBER(MATCH(RC34,Слаботочные_коды,0)),1,0))')
    sl = f"RC{c}=1"
    n = 0
    for col, zero in (("T", "0"), ("AA", "0"), ("AB", "0"), ("V", 'IF(N(RC17)>0,RC17,"")')):
        fs = [x[0] for x in ws.Range(f"{col}3:{col}1000").Formula2R1C1]
        for i, f in enumerate(fs):
            if isinstance(f, str) and f.startswith("=") and sl not in f:
                ws.Range(f"{col}{3 + i}").Formula2R1C1 = f"=IF({sl},{zero},{f[1:]})"
                n += 1
    log.append(f"«Исходные данные»: формул T/V/AA/AB обёрнуто {n}")
    if opts:
        ws.Protect(PWD, **opts)
    return c


def nag_sheet(wb, c, log):
    ws = wb.Worksheets(NAG)
    opts = unprotect(ws)
    cond = f"IFERROR(INDEX('{SRC}'!C{c},MATCH(RC4,'{SRC}'!C4,0))=1,FALSE)"
    fs = [x[0] for x in ws.Range("H3:H1000").Formula2R1C1]
    n = 0
    for i, f in enumerate(fs):
        if isinstance(f, str) and f.startswith("=") and cond not in f:
            ws.Range(f"H{3 + i}").Formula2R1C1 = f"=IF({cond},0,{f[1:]})"
            n += 1
    log.append(f"«Нагрузка_щитов» H: обёрнуто {n}")
    if opts:
        ws.Protect(PWD, **opts)


def odn_sheet(wb, c, log):
    ws = wb.Worksheets(ODN)
    opts = unprotect(ws)
    head = str(ws.Range(f"{ODN_COL}2").Value or "").strip()
    if head not in ("", ODN_HDR):
        raise RuntimeError(f"«Однолинейка» {ODN_COL} занят («{head}») — SXDRAW ждёт там признак слаботочной линии")
    ws.Range(f"{ODN_COL}2").Value = ODN_HDR
    ws.Range(f"{ODN_COL}3:{ODN_COL}1000").FormulaR1C1 = (
        f'=IF(RC18="",0,IFERROR(INDEX(\'{SRC}\'!C{c},MATCH(RC18,\'{SRC}\'!C4,0))*1,0))')
    fs = [x[0] for x in ws.Range("AE3:AE1000").Formula]
    n = 0
    for i, f in enumerate(fs):
        r = 3 + i
        old = f'IF(AND(R{r}<>"",N(T{r})=0),"нет мощности","")'
        if isinstance(f, str) and old in f and MSG not in f:
            new = (f'IF(AND(R{r}<>"",N(T{r})=0,N({ODN_COL}{r})<>1),"нет мощности",""),'
                   f'IF(AND(N({ODN_COL}{r})=1,E{r}<>"",I{r}&J{r}&K{r}=""),"{MSG}","")')
            ws.Range(f"AE{r}").Formula2 = f.replace(old, new)
            n += 1
    log.append(f"«Однолинейка» AE: статус обновлён в {n} строках")
    if opts:
        ws.Protect(PWD, **opts)


def run(paths):
    app, pid = excel()
    try:
        for path in paths:
            wb = app.Workbooks.Open(os.path.abspath(path), 0, False)
            try:
                app.Calculation = -4135
                log = []
                if wb.ReadOnly:                      # открыта у другого — не трогаем
                    print(f"{os.path.basename(path)}: открыта у другого пользователя — пропущена")
                    continue
                ref_sheet(wb, log, manual_codes(wb))
                c = src_sheet(wb, log)
                nag_sheet(wb, c, log)
                odn_sheet(wb, c, log)
                app.Calculation = -4105
                app.CalculateFull()
                wb.Save()
                print(f"{os.path.basename(path)}: " + "; ".join(log))
            finally:
                wb.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    run(sys.argv[1:])
