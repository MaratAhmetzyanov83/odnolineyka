# Перенос книги 1.05 в шаблон v3.12: py migrate_105.py <старая.xlsx> <новая.xlsx>  (книги закрыты, старую — из резерва)
# «Исходные данные» A:S, «Разбивка_по_щитам» A:U: своё число / текст и своя простая формула (=G21*0.01) главнее
# формулы шаблона; служебные формулы 1.05 не переносятся.
# «Справочная»: коды P:T и марки кабелей L:N из старой книги (первая строка кода; свои значения главнее,
# новые — в свободные строки, сначала коды, что есть в линиях).
# «Спецификация»: в 1.05 ручная — B3:J (значения и формат) из старой книги, если в шаблоне нет позиций (D:H пусто).
# «Однолинейка»: E = «авто» у линий (кроме вводов). Затем сравнение ток / автомат / сечение / ΔU / кабель.
import os, sys, shutil, signal, tempfile
import win32com.client, win32process
sys.stdout.reconfigure(encoding="utf-8")
OLD, NEW = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
TPL = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "1_Таблицы_Excel", "Однолинейка_шаблон.xlsx")
PWD = "sx"
_tmp = os.path.join(tempfile.gettempdir(), "sx_old_105_" + os.path.basename(OLD))   # другое имя: Excel не откроет две книги с одним именем
shutil.copyfile(OLD, _tmp); OLD = _tmp
shutil.copyfile(TPL, NEW)


def unprot(ws):
    if not ws.ProtectContents:
        return None
    p = ws.Protection
    o = (True, True, True, False, p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
         p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks, p.AllowDeletingColumns,
         p.AllowDeletingRows, p.AllowSorting, p.AllowFiltering, p.AllowUsingPivotTables)
    ws.Unprotect(PWD)
    return o


import re
FUNC = re.compile(r"[A-Za-zА-Яа-я_][A-Za-zА-Яа-я_0-9.]*\(")


def copy_inputs(so, sn, rng):
    # своё число / текст или своя простая формула (=G21*0.01, =M21-J21) главнее формулы шаблона;
    # служебные формулы 1.05 (с функциями, листами) не переносим — остаётся формула шаблона
    fo, fn = so.Range(rng).Formula, sn.Range(rng).Formula
    out, n = [], 0
    for ro, rn in zip(fo, fn):
        row = []
        for a, b in zip(ro, rn):
            s = str(a) if a is not None else ""
            if s == "":
                row.append(b)
            elif not s.startswith("="):
                row.append(a); n += 1
            elif not FUNC.search(s) and "!" not in s and "[" not in s:
                row.append(a); n += 1
            else:
                row.append(b)
        out.append(row)
    sn.Range(rng).Formula = out
    return n


app = win32com.client.DispatchEx("Excel.Application")
pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
app.Visible = False; app.DisplayAlerts = False; app.AskToUpdateLinks = False
log = []
try:
    old = app.Workbooks.Open(OLD, 0, True)
    new = app.Workbooks.Open(NEW, 0, False)
    app.Calculation = -4135
    for sh, rng in (("Исходные данные", "A3:S1000"), ("Разбивка_по_щитам", "A3:U1000")):
        sn = new.Worksheets(sh); o = unprot(sn)
        log.append(f"{sh}: перенесено ячеек {copy_inputs(old.Worksheets(sh), sn, rng)}")
        if o: sn.Protect(PWD, *o)
    # «Справочная»: коды потребителей P:T и марки кабелей L:N
    ro, rn = old.Worksheets("Справочная"), new.Worksheets("Справочная"); o = unprot(rn)
    for c0, c1, last in (("P", "T", 51), ("L", "N", 51)):
        have = {str(rn.Range(f"{c0}{r}").Value): r for r in range(4, last + 1) if rn.Range(f"{c0}{r}").Value}
        free = [r for r in range(4, last + 1) if not rn.Range(f"{c0}{r}").Value]
        upd, add, skip = [], [], []
        # сначала коды, которые есть в линиях книги (D «Исходных данных» и N — марка кабеля)
        used = set()
        for r in range(3, 1001):
            v = str(old.Worksheets("Исходные данные").Range(("D" if c0 == "P" else "N") + str(r)).Value or "")
            used.add(v.split(".")[0] if c0 == "P" else v)
        order = sorted((r for r in range(4, 200) if ro.Range(f"{c0}{r}").Value),
                       key=lambda r: str(ro.Range(f"{c0}{r}").Value) not in used)
        seen = set()
        for r in order:
            k = ro.Range(f"{c0}{r}").Value
            if not k or str(k) in seen:
                continue                    # VLOOKUP берёт первую строку кода — повтор ниже не считается
            seen.add(str(k))
            vals = ro.Range(f"{c0}{r}:{c1}{r}").Value
            if str(k) in have:
                if rn.Range(f"{c0}{have[str(k)]}:{c1}{have[str(k)]}").Value != vals:
                    rn.Range(f"{c0}{have[str(k)]}:{c1}{have[str(k)]}").Value = vals; upd.append(str(k))
            elif free:
                row = free.pop(0)
                rn.Range(f"{c0}{row}:{c1}{row}").Value = vals; add.append(str(k)); have[str(k)] = row
            else:
                skip.append(str(k))
        log.append(f"Справочная {c0}:{c1}: свои значения {len(upd)} ({', '.join(upd)}); добавлены {', '.join(add) or '—'}"
                   + (f"; НЕ поместились: {', '.join(skip)}" if skip else ""))
    if o: rn.Protect(PWD, *o)
    # «Спецификация»: позиции (D:H — тип, код, завод, ед., кол-во) есть в старой и нет в шаблоне — переносим лист
    so, sn = old.Worksheets("Спецификация"), new.Worksheets("Спецификация")
    filled = lambda ws: app.WorksheetFunction.CountA(ws.Range("D3:H1000")) > 0   # noqa: E731
    if filled(so) and not filled(sn):
        last = so.Cells(so.Rows.Count, 3).End(-4162).Row
        o = unprot(sn)
        sn.Range("B3:J1000").Clear()
        so.Range(f"B3:J{last}").Copy(sn.Range("B3"))
        if o: sn.Protect(PWD, *o)
        log.append(f"Спецификация: перенесены строки 3–{last}")
    else:
        log.append("Спецификация: " + ("в шаблоне уже есть позиции — не тронута" if filled(sn) else "в старой книге пусто"))
    app.Calculation = -4105
    app.CalculateFull()
    # «Однолинейка»: «авто» у линий без автомата (вводы не трогаем)
    od = new.Worksheets("Однолинейка"); o = unprot(od); n = 0
    for r in range(3, 1001):
        line = str(od.Range(f"R{r}").Value or "").strip()
        if line and line not in ("0", "0.0") and not line.startswith(("Ввод", "#")) \
                and not str(od.Range(f"E{r}").Formula or "").strip():
            od.Range(f"E{r}").Value = "авто"; n += 1
    if o: od.Protect(PWD, *o)
    app.CalculateFull()
    log.append(f"Однолинейка: «авто» у {n} линий")
    # сравнение расчёта: старая и новая
    so, sn = old.Worksheets("Исходные данные"), new.Worksheets("Исходные данные")
    diff = 0; lines = 0
    for r in range(3, 1001):
        d = so.Range(f"D{r}").Text
        if not d or d in ("0", "#N/A"):
            continue
        lines += 1
        a = [so.Range(f"{c}{r}").Text for c in ("T", "U", "V", "AB", "AD")]
        b = [sn.Range(f"{c}{r}").Text for c in ("T", "U", "V", "AB", "AD")]
        if a != b:
            diff += 1
            print(f"   {d:14} было {a}\n   {'':14} стало {b}")
    log.append(f"сравнение: линий {lines}, отличаются {diff}")
    new.Save()
    old.Close(False); new.Close(False)
finally:
    try: app.Quit()
    except Exception: pass
    try: os.kill(pid, signal.SIGTERM)
    except Exception: pass
print(os.path.basename(OLD), "->", NEW)
for s in log:
    print("  " + s)
