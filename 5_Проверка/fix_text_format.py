# -*- coding: utf-8 -*-
"""TEXT(x,"0.0") -> FIXED(x,1,TRUE): исправление #ЗНАЧ! на русском Excel.

    py fix_text_format.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит
    py fix_text_format.py --check <книга.xlsx>  — только показать, где

В русском Excel формат «0.0» в TEXT не понимается (разделитель — запятая) и TEXT даёт #ЗНАЧ!.
FIXED (ФИКСИРОВАННЫЙ) от языка не зависит; TRUE — без пробелов между тысячами.
Где было в v3.12: «Однолинейка» AE (статус «потери … %»), «Питание_щита» C40:E40
(термическая стойкость), «Расчёт_нагрузок» C44, C58. TEXT(x,"0") не трогается — он работает везде.
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import re
import signal
import sys

PWD = "sx"
FMT = {"0.0": 1, "0.00": 2, "0.000": 3}


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def _close(s, i):
    """индекс закрывающей скобки для открывающей s[i] == '(' (с учётом строк в кавычках)"""
    depth, q = 0, False
    for j in range(i, len(s)):
        ch = s[j]
        if ch == '"':
            q = not q
        elif not q:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    return j
    return -1


def _last_arg(s, a, b):
    """начало последнего аргумента верхнего уровня в s[a:b]"""
    depth, q, cut = 0, False, -1
    for j in range(a, b):
        ch = s[j]
        if ch == '"':
            q = not q
        elif not q:
            if ch in "({":
                depth += 1
            elif ch in ")}":
                depth -= 1
            elif ch == "," and depth == 0:
                cut = j
    return cut


def fix_formula(f):
    out, pos, n = [], 0, 0
    for m in re.finditer(r"(?<![A-Za-z_.])TEXT\(", f):
        if m.start() < pos:
            continue
        op = m.end() - 1
        cl = _close(f, op)
        if cl < 0:
            continue
        cut = _last_arg(f, op + 1, cl)
        if cut < 0:
            continue
        fmt = f[cut + 1:cl].strip()
        if not (fmt.startswith('"') and fmt.endswith('"') and fmt[1:-1] in FMT):
            continue
        inner = fix_formula(f[op + 1:cut])[0]
        out.append(f[pos:m.start()])
        out.append(f"FIXED({inner},{FMT[fmt[1:-1]]},TRUE)")
        pos = cl + 1
        n += 1
    out.append(f[pos:])
    return "".join(out), n


def run(path, check=False):
    path = os.path.abspath(path)
    app, pid = excel()
    found = {}
    try:
        wb = app.Workbooks.Open(path, 0, check)
        app.Calculation = -4135
        for ws in wb.Worksheets:
            # защищённый лист со скрытыми формулами отдаёт .Formula пустым — сначала снять защиту
            prot = ws.ProtectContents
            if prot:
                p = ws.Protection
                opts = (True, True, True, False,
                        p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
                        p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks,
                        p.AllowDeletingColumns, p.AllowDeletingRows, p.AllowSorting,
                        p.AllowFiltering, p.AllowUsingPivotTables)
                ws.Unprotect(PWD)
            try:
                rng = ws.UsedRange.SpecialCells(-4123)      # формулы
            except Exception:
                rng = None
            if rng is None or not any("TEXT(" in str(a.Formula) for a in rng.Areas):
                if prot:
                    ws.Protect(PWD, *opts)
                continue
            for area in rng.Areas:
                vals = area.Formula
                if not isinstance(vals, tuple):
                    vals = ((vals,),)
                for i, row in enumerate(vals):
                    for j, f in enumerate(row):
                        if not (isinstance(f, str) and "TEXT(" in f):
                            continue
                        g, n = fix_formula(f)
                        if n:
                            cell = area.Cells(i + 1, j + 1)
                            found.setdefault(ws.Name, []).append(cell.Address.replace("$", ""))
                            if not check:
                                if cell.HasArray:
                                    cell.FormulaArray = g
                                else:
                                    cell.Formula = g
            if prot:
                ws.Protect(PWD, *opts)
        if not check and found:
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
    print(os.path.basename(path) + (" (проверка)" if check else ""))
    if not found:
        print("  TEXT с форматом 0.0 / 0.00 / 0.000 нет")
    for name, cells in found.items():
        tail = ", ".join(cells[:4]) + (f" … всего {len(cells)}" if len(cells) > 4 else "")
        print(f"  «{name}»: {tail}")


if __name__ == "__main__":
    args = sys.argv[1:]
    chk = "--check" in args
    files = [a for a in args if a != "--check"]
    if not files:
        print(__doc__)
        sys.exit(1)
    for p in files:
        run(p, chk)
