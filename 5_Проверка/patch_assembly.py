# -*- coding: utf-8 -*-
"""Сборка щита для CPDRAW v2.0: лист «Сборка_щита» + связь с «Каталогом».

    py patch_assembly.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит

1. Лист «Сборка_щита» в конце книги (если нет): настройки проводов (цвета L1/L2/L3/N/PE, марка, запасы),
   таблица «сечение по току аппарата», итог для спецификации (A21:G) и перечень проводов (J21:Q).
   Заполняет CPDRAW / CPUPDATE (AutoCAD). Раскладка листа — как в SX_Panel.lsp (cp:asm-init).
2. «Каталог», столбец N: для разделов «Провод», «Шины», «Корпус» — количество с листа «Сборка_щита»
   (SUMIFS по наименованию G). Столбец O («вручную / запас») остаётся: прежние ручные числа там надо убрать.
3. «Каталог», Q3 — «Блок компоновки (CPDRAW)»: имя блока вида спереди для модели (главнее подбора).
4. «Каталог»: шины PIN Chint 1P, 1P+N, 3P, 4P (L — модулей в шине 1 м), заглушка, шина PE — если их нет.
   Артикулы не заполнены: впишите по прайсу.
Свой скрытый экземпляр Excel; закрывается только он. COM — только позиционные аргументы.
"""
import os
import signal
import sys

PWD = "sx"
ASM, CAT = "Сборка_щита", "Каталог"
BRANCH = ('IF(OR(A{r}="Провод",A{r}="Шины",A{r}="Корпус"),IF(G{r}="","",'
          'SUMIFS(Сборка_щита!$E$21:$E$5000,Сборка_щита!$C$21:$C$5000,G{r})),"")')
COMBS = [
    ("Шины", "Chint", "PIN 1P 63А", "Шина соединительная PIN 1P 63А, 1 м (под NB1-63 1P)", "1P", 56),
    ("Шины", "Chint", "PIN 1P+N 63А", "Шина соединительная PIN 1P+N 63А, 1 м (под NB1L 1P+N)", "1P+N", 56),
    ("Шины", "Chint", "PIN 3P 63А", "Шина соединительная PIN 3P 63А, 1 м", "3P", 56),
    ("Шины", "Chint", "PIN 4P 63А", "Шина соединительная PIN 4P 63А, 1 м", "4P", 56),
    ("Шины", "Chint", "Заглушка PIN", "Заглушка торцевая для шины PIN", "заглушка", None),
    ("Шины", "", "Шина PE 12", "Шина PE на изоляторах, 12 подключений", "PE", 12),
]


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def unprot(ws):
    """снять защиту; вернуть её параметры или None"""
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


def put(ws, addr, rows):
    ws.Range(addr).Value = rows


def build_sheet(wb, log):
    try:
        wb.Worksheets(ASM)
        return False
    except Exception:
        pass
    last = wb.Worksheets(wb.Worksheets.Count)
    ws = wb.Worksheets.Add(None, last)
    ws.Name = ASM
    put(ws, "A1", "СБОРКА ЩИТА — считает CPDRAW (AutoCAD). Настройки — столбцы B и D:E. "
                  "«Каталог», столбец N, берёт количество по наименованию (столбец C).")
    put(ws, "A3:B12", (("НАСТРОЙКИ", ""), ("Цвет L1", "белый"), ("Цвет L2", "черный"), ("Цвет L3", "красный"),
                       ("Цвет N", "синий"), ("Цвет PE", "Ж/З"), ("Марка провода", "ПуГВ"),
                       ("Запас на подключение (на провод), мм", 150), ("Запас длины, %", 10), ("Округлять до, м", 1)))
    put(ws, "D3:E16", (("СЕЧЕНИЕ ПО ТОКУ АППАРАТА", ""), ("In до, А", "Сечение, кв. мм"),
                       (10, 1.5), (16, 2.5), (25, 4), (32, 6), (40, 10), (63, 16), (80, 25), (100, 35),
                       (125, 50), (160, 70), (200, 95), (250, 120)))
    put(ws, "G3:G8", (("КАК СЧИТАЕТСЯ",),
                      ("Аппараты 1P и 1P+N стоят по фазам L1-L2-L3 подряд, питание — гребёнкой; фаза — из «Однолинейки», столбец V.",),
                      ("Провод: кросс-модуль -> гребёнка / аппарат -> аппараты группы -> клеммы L, N, PE.",),
                      ("Сечение — по номиналу аппарата (таблица D:E). Длина — по чертежу + запас на подключение + запас длины.",),
                      ("Цвета и марку провода пишите как в «Каталоге», столбец G: по ним ищется строка.",),
                      ("Строки щита перезаписываются при каждом CPDRAW / CPUPDATE.",)))
    put(ws, "A19", "ИТОГО ДЛЯ СПЕЦИФИКАЦИИ")
    put(ws, "A20:G20", (("Щит", "Вид", "Наименование (как в «Каталоге», G)", "Модель", "Кол-во", "Ед.", "Есть в «Каталоге»"),))
    put(ws, "J19", "ПРОВОДА ДЛЯ СБОРКИ")
    put(ws, "J20:Q20", (("Щит", "№", "Откуда", "Куда", "Проводник", "Цвет", "Сечение, кв. мм", "Длина, мм"),))
    for a in ("A3", "D3", "G3", "A19", "J19", "A20:G20", "J20:Q20"):
        ws.Range(a).Font.Bold = True
    for col, w in (("A:A", 34), ("B:B", 12), ("C:C", 52), ("D:D", 14), ("E:E", 8), ("G:G", 16), ("L:M", 10)):
        ws.Range(col).ColumnWidth = w
    ws.Range("B4:B12").Interior.Color = 0xCCFFFF
    ws.Range("D5:E16").Interior.Color = 0xCCFFFF
    log.append(f"лист «{ASM}» добавлен в конец книги")
    return True


def patch_catalog(wb, log):
    ws = wb.Worksheets(CAT)
    n = 0
    for r in range(4, 1001):
        f = ws.Range(f"N{r}").Formula
        if not isinstance(f, str) or not f.startswith("=") or "Сборка_щита" in f:
            continue
        if f.endswith(',""))))') and '"Клемма"' in f:
            ws.Range(f"N{r}").Formula = f[:-len('""))))')] + BRANCH.format(r=r) + "))))"
            n += 1
    if n:
        log.append(f"«{CAT}» N: провод, шины, корпус — с листа «{ASM}» ({n} строк)")
    if not ws.Range("Q3").Value:
        ws.Range("Q3").Value = "Блок компоновки (CPDRAW)"
        ws.Range("Q3").Font.Bold = True
        ws.Range("Q3").WrapText = True
        log.append(f"«{CAT}» Q3: «Блок компоновки (CPDRAW)»")
    models = {str(ws.Range(f"C{r}").Value or "").strip() for r in range(4, 1001)}
    last = 3
    for r in range(4, 1001):
        if ws.Range(f"A{r}").Value not in (None, ""):
            last = r
    added = 0
    for cat, maker, model, name, poles, mods in COMBS:
        if model in models:
            continue
        last += 1
        ws.Range(f"A{last}").Value = cat
        ws.Range(f"B{last}").Value = maker
        ws.Range(f"C{last}").Value = model
        ws.Range(f"G{last}").Value = name
        ws.Range(f"I{last}").Value = poles
        if mods:
            ws.Range(f"L{last}").Value = mods
        ws.Range(f"M{last}").Value = "шт."
        added += 1
    if added:
        log.append(f"«{CAT}»: шины PIN Chint, заглушка, шина PE — {added} строк (артикулы впишите)")


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        app.Calculation = -4135
        struct = wb.ProtectStructure
        if struct:
            wb.Unprotect(PWD)
        cat = wb.Worksheets(CAT)
        opts = unprot(cat)
        build_sheet(wb, log)
        patch_catalog(wb, log)
        if opts:
            cat.Protect(PWD, *opts)
        if struct:
            wb.Protect(PWD, True, False)
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
