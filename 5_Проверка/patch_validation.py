# -*- coding: utf-8 -*-
"""Исправления по валидации таблицы 11.10.2026 (аудит, раздел 6: ошибки 15–17) для книг v3.12.

    py patch_validation.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит

15. Строка ввода (код «Ввод») считалась от ручной мощности E — автомат ввода мог быть меньше тока щита.
    «Исходные данные» T: для кода «Ввод» ток = Iр щита по наиболее нагруженной фазе («Нагрузка_щитов» AF12 / AF21 / AF30,
    блок щита этой линии). Автомат и кабель ввода подбираются по нему.
16. В «Кабель_сопротивление» («Справочная» B:D) не было 120 мм² — ΔU давал #Н/Д. Добавлена строка 120 (x0 0,06; ρ 0,0225).
17. «Нагрузка_щитов»: фазы L1–L3 (O7:Q7, …), мощность по фазам (Y5:Y7, …), cos φ (V8, …) включали группу «Нет»
    (ввод, резерв) — теперь SUMIFS с условием «группа ≠ Нет», как в Iр (AF12).
18. «Исходные данные» I, K, L (длина по плану, гофра ПВХ / ПНД из DWG) были не во всех строках (дыры 36–37, 53–54,
    70–71, 79+): линия в такой строке оставалась без гофры. Формулы строки 21 дотянуты на пустые ячейки 22–1000.
Свой скрытый экземпляр Excel; закрывается только он. COM — только позиционные аргументы.
"""
import os
import re
import signal
import sys

PWD = "sx"
SRC, NAG, REF = "Исходные данные", "Нагрузка_щитов", "Справочная"
BOARD = "INDEX(Разбивка_по_щитам!R1C6:R1000C6,MATCH(RC4,Разбивка_по_щитам!R1C8:R1000C8,0))"
BOARD_I = (f"IFERROR(CHOOSE(MATCH({BOARD},ПЩ_щиты,0),"
           "Нагрузка_щитов!R12C32,Нагрузка_щитов!R21C32,Нагрузка_щитов!R30C32)*1,\"\")")
NOT_NET = '$A$3:$A$1000,"<>Нет"'


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


def runs(ws, col, r0=3, r1=1000):
    fs = [x[0] for x in ws.Range(f"{col}{r0}:{col}{r1}").Formula2R1C1]
    out, start = [], r0
    for i in range(1, len(fs) + 1):
        if i == len(fs) or fs[i] != fs[i - 1]:
            out.append((start, r0 + i - 1, fs[i - 1]))
            start = r0 + i
    return out


# ---- 15
def rule_t(f):
    if not isinstance(f, str) or not f.startswith("=") or "Нагрузка_щитов!R12C32" in f:
        return None
    return f'=IF(AND(RC34="Ввод",ISNUMBER({BOARD_I})),{BOARD_I},' + f[1:] + ")"


def patch_input_current(wb, log):
    ws = wb.Worksheets(SRC)
    opts = unprotect(ws)
    n = 0
    for a, b, f in runs(ws, "T"):
        g = rule_t(f)
        if g:
            ws.Range(f"T{a}:T{b}").Formula2R1C1 = g
            n += b - a + 1
    if opts:
        ws.Protect(PWD, *opts)
    if n:
        log.append(f"15: «{SRC}» T — ввод от Iр щита, {n} строк")


# ---- 16
def patch_120(wb, log):
    ws = wb.Worksheets(REF)
    secs = [ws.Cells(r, 2).Value for r in range(4, 30)]
    if 120 in secs or 120.0 in secs:
        return
    last = 4 + max(i for i, v in enumerate(secs) if v not in (None, ""))
    r = last + 1
    if any(ws.Cells(r, c).Value not in (None, "") for c in (2, 3, 4)):
        raise RuntimeError(f"«{REF}» B{r}:D{r} занято — некуда добавить 120 мм²")
    ws.Range(f"B{r}:D{r}").Value = (120, 0.06, 0.0225)
    ws.Range(f"B{last}:D{last}").Copy()
    ws.Range(f"B{r}:D{r}").PasteSpecial(-4122)            # форматы как у соседней строки
    wb.Application.CutCopyMode = False
    wb.Names("Кабель_сопротивление").RefersTo = f"='{REF}'!$B$4:$D${r}"
    log.append(f"16: «{REF}» B{r}: сечение 120 мм², «Кабель_сопротивление» до строки {r}")


# ---- 17
def split_args(s):
    out, depth, q, cur = [], 0, False, ""
    for ch in s:
        if ch == '"':
            q = not q
        if not q:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "," and depth == 0:
                out.append(cur)
                cur = ""
                continue
        cur += ch
    out.append(cur)
    return out


def to1000(rng):
    return re.sub(r"\$(\d+)$", "$1000", rng)


def fix_sumif(f):
    """SUMIF(диапазон, условие, сумма) -> SUMIFS(сумма, диапазон, условие, группа <> «Нет»)"""
    out, i, n = "", 0, 0
    while True:
        j = f.find("SUMIF(", i)
        if j < 0 or (j > 0 and f[j - 1].isalpha()):
            if j < 0:
                out += f[i:]
                break
            out += f[i:j + 6]
            i = j + 6
            continue
        k, depth = j + 6, 1                       # k — за закрывающей скобкой SUMIF(...)
        while depth:
            if f[k] == "(":
                depth += 1
            elif f[k] == ")":
                depth -= 1
            k += 1
        args = split_args(f[j + 6:k - 1])
        if len(args) == 3:
            out += f[i:j] + f"SUMIFS({to1000(args[2])},{to1000(args[0])},{args[1]},{NOT_NET})"
            n += 1
        else:
            out += f[i:k]
        i = k
    return out, n


def patch_phases(wb, log):
    ws = wb.Worksheets(NAG)
    opts = unprotect(ws)
    n = 0
    for b in (0, 9, 18):                                   # блоки щитов 1–3
        for addr in (f"O{7 + b}", f"P{7 + b}", f"Q{7 + b}", f"V{8 + b}"):
            c = ws.Range(addr)
            f = c.Formula
            if isinstance(f, str) and "SUMIF(" in f:
                g, k = fix_sumif(f)
                if k:
                    c.Formula = g
                    n += 1
    if opts:
        ws.Protect(PWD, *opts)
    if n:
        log.append(f"17: «{NAG}» фазы и cos φ без группы «Нет», {n} формул")


# ---- 18
def patch_holes(wb, log):
    ws = wb.Worksheets(SRC)
    opts = unprotect(ws)
    n = 0
    for col in ("I", "K", "L"):
        f21 = ws.Range(f"{col}21").FormulaR1C1
        if not (isinstance(f21, str) and f21.startswith("=")):
            continue
        for a, b, f in runs(ws, col, 22):
            if f in ("", None):
                ws.Range(f"{col}{a}:{col}{b}").FormulaR1C1 = f21
                n += b - a + 1
    if opts:
        ws.Protect(PWD, *opts)
    if n:
        log.append(f"18: «{SRC}» I, K, L — формулы дотянуты на {n} пустых ячеек")


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        app.Calculation = -4135
        patch_input_current(wb, log)
        patch_120(wb, log)
        patch_phases(wb, log)
        patch_holes(wb, log)
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
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for p in sys.argv[1:]:
        run(p)
