# -*- coding: utf-8 -*-
"""Модель объекта: жилой дом, ГРЩ с двумя вводами от ТП (2 независимых трансформатора 630 кВА), АВР на СВ.

    py build_demo_object.py <папка> [--подпапки]   (папка создаётся; объекты держим вне репозитория)
    --подпапки: ГРЩ в корне, ЩР-1 и ЩК-1..3 — «Корпус 1», ЩС-1 и ЩО-1 — «Инженерия» (проверка связей между папками)

Щиты (одна книга = один щит, все в одной папке):
  ГРЩ  — 2 ввода от ТП-1 / ТП-2, АВР, I категория. Фидеры: ЩР-1 (секция 1), ЩС-1 двумя фидерами
         (секция 1 и секция 2), ЩО-1 (секция 2); свои линии — свет и розетка электрощитовой.
  ЩР-1 — распределительный щит жилой части: фидеры квартирных щитов ЩК-1, ЩК-2, ЩК-3, свет и розетки МОП.
  ЩС-1 — силовой: лифт, насосы, вентиляция; 2 ввода + АВР (I категория у лифта и насоса).
  ЩО-1 — общедомовое освещение.
  ЩК-1..3 — квартирные щиты (розетки, свет, варочная панель 3ф, водонагреватель, кондиционер, тёплый пол).
Освещение МОП — код LM (группа «Пост»). Книги — из шаблона v3.12 (1_Таблицы_Excel/Однолинейка_шаблон.xlsx). Связи ставит макрос надстройки
SX_LinkUpstreamTo («Связать с вышестоящим»), затем все книги пересчитываются вместе и строится «Дерево щитов».
Свой скрытый экземпляр Excel; закрывается только он. COM — только позиционные аргументы.
"""
import os
import shutil
import signal
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
TPL = os.environ.get("SX_TPL") or os.path.join(ROOT, "1_Таблицы_Excel", "Однолинейка_шаблон.xlsx")  # SX_TPL — другой шаблон (проба)
ADDIN = os.environ.get("SX_ADDIN") or os.path.join(ROOT, "3_Макросы_Excel", "Однолинейка.xlam")  # SX_ADDIN — другая сборка (проба)
PWD = "sx"
SRC0, ODN0 = 21, 3

# коды Щ, ЛФ, НС, LM — в шаблоне («Справочная», patch_review.py)
P3 = dict(F=400, H=3, N="мс", O=5)            # трёхфазная линия


def ln(num, kw, length, sec=None, name=None, **kw_):
    d = {"D": num, "E": kw, "I": length}
    if sec:
        d["BB"] = sec
    if name:
        d["C"] = name
    d.update(kw_)
    return d


def flat(i):
    """квартира i: свои мощности, чтобы щиты отличались"""
    k = 1 + 0.1 * (i - 1)
    return [
        ln("Ввод.1", 0, 2, None, "Ввод питающей линии", **P3),
        ln(f"R.{i}.01.1", round(1.0 * k, 2), 12), ln(f"R.{i}.02.1", round(0.8 * k, 2), 15),
        ln(f"R.{i}.03.1", round(1.5 * k, 2), 10, name="Розетки кухни"),
        ln(f"R.{i}.03.2", 7.0, 9, name="Варочная панель", **P3),
        ln(f"R.{i}.04.1", round(1.2 * k, 2), 14, name="Розетки санузла"),
        ln(f"L.{i}.01.1", 0.3, 14), ln(f"L.{i}.02.1", 0.25, 16), ln(f"L.{i}.03.1", 0.2, 11),
        ln(f"ПН.{i}.04.1", 3.5, 10, F=230, H=1, N="мс", O=3),
        ln(f"K.{i}.02.1", 1.2, 18, F=230, H=1, N="мс", O=3),
        ln(f"ТП.{i}.04.1", round(0.8 * k, 2), 12, F=230, H=1, N="мс", O=3),
    ]


BOOKS = {
    "ГРЩ": dict(inputs=2, avr="да", lines=[
        ln("Ввод.1", 0, 3, 1, "Ввод 1 (ТП-1)", **P3),
        ln("Ввод.2", 0, 3, 2, "Ввод 2 (ТП-2)", **P3),
        ln("Щ.0.01.1", 20, 25, 1, "ЩР-1 (жилая часть)", R=80, **P3),
        ln("Щ.0.02.1", 20, 40, 1, "ЩС-1, ввод 1", R=40, **P3),
        ln("Щ.0.02.2", 20, 42, 2, "ЩС-1, ввод 2", R=40, **P3),
        ln("Щ.0.03.1", 5, 30, 2, "ЩО-1 (освещение)", R=16, **P3),
        ln("LM.0.01.1", 0.2, 8, 1, "Освещение электрощитовой"),
        ln("R.0.01.1", 0.5, 6, 2, "Розетки электрощитовой"),
    ]),
    "ЩР-1": dict(inputs=1, lines=[
        ln("Ввод.1", 0, 2, None, "Ввод питающей линии", **P3),
        ln("Щ.1.01.1", 15, 20, None, "Квартира 1 (ЩК-1)", R=40, **P3),
        ln("Щ.1.02.1", 15, 26, None, "Квартира 2 (ЩК-2)", R=40, **P3),
        ln("Щ.1.03.1", 15, 32, None, "Квартира 3 (ЩК-3)", R=40, **P3),
        ln("LM.1.01.1", 0.4, 30, None, "Освещение МОП"),
        ln("R.1.01.1", 0.5, 25, None, "Розетки МОП"),
    ]),
    "ЩС-1": dict(inputs=2, avr="да", lines=[
        ln("Ввод.1", 0, 2, 1, "Ввод 1", **P3),
        ln("Ввод.2", 0, 2, 2, "Ввод 2", **P3),
        ln("ЛФ.0.01.1", 7.5, 35, 1, "Лифт пассажирский", BN="I"),
        ln("НС.0.01.1", 4.0, 15, 2, "Насос ХВС (повысительный)", BN="I"),
        ln("ПУ.0.01.1", 5.5, 20, 1, "Приточная установка П1", **P3),
        ln("ПВ.0.01.1", 3.0, 22, 2, "Вытяжная установка В1", **P3),
        ln("НС.0.02.1", 1.5, 12, 1, "Дренажный насос", **dict(P3, F=400)),
    ]),
    "ЩО-1": dict(inputs=1, lines=[
        ln("Ввод.1", 0, 2, None, "Ввод питающей линии", **P3),
        ln("LM.0.01.1", 0.6, 40, None, "Освещение лестничной клетки"),
        ln("LM.0.02.1", 0.4, 35, None, "Освещение входной группы"),
        ln("LM.0.03.1", 0.5, 45, None, "Освещение подвала"),
        ln("LM.0.04.1", 0.3, 30, None, "Освещение кровли"),
        ln("LED.0.01.1", 0.8, 50, None, "Фасадная подсветка", F=230, H=1, N="мс", O=3),
    ]),
    "ЩК-1": dict(inputs=1, lines=flat(1)),
    "ЩК-2": dict(inputs=1, lines=flat(2)),
    "ЩК-3": dict(inputs=1, lines=flat(3)),
}
# (нижестоящий, вышестоящий, фидер, ввод нижестоящего)
LINKS = [
    ("ЩР-1", "ГРЩ", "Щ.0.01.1", 1),
    ("ЩС-1", "ГРЩ", "Щ.0.02.1", 1),
    ("ЩС-1", "ГРЩ", "Щ.0.02.2", 2),
    ("ЩО-1", "ГРЩ", "Щ.0.03.1", 1),
    ("ЩК-1", "ЩР-1", "Щ.1.01.1", 1),
    ("ЩК-2", "ЩР-1", "Щ.1.02.1", 1),
    ("ЩК-3", "ЩР-1", "Щ.1.03.1", 1),
]
# ГРЩ: питание от ТП-1 (раздел 2) и ТП-2 (Ввод 2 — те же параметры)
GRSH_POWER = [("C12", "ТП"), ("C16", 630), ("C17", "Д/Yн"), ("C18", "АПвБбШв"), ("C19", 95), ("C20", 4),
              ("C21", 1), ("C22", 60), ("C23", "алюминий"), ("C24", 160), ("C25", "C")]


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 1          # макросы надстройки (своя копия Excel)
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def unprot(ws):
    if ws.ProtectContents:
        ws.Unprotect(PWD)


def fill_book(app, path, board, spec):
    wb = app.Workbooks.Open(path, 0, False)
    app.Calculation = -4135
    src, rz, odn, ref, net = (wb.Worksheets(n) for n in ("Исходные данные", "Разбивка_по_щитам", "Однолинейка",
                                                          "Справочная", "Сеть_щита"))
    p = odn.Protection
    opts = (True, True, True, False, p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
            p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks, p.AllowDeletingColumns,
            p.AllowDeletingRows, p.AllowSorting, p.AllowFiltering, p.AllowUsingPivotTables) if odn.ProtectContents else None
    for ws in (src, rz, odn, ref):
        unprot(ws)
    rz.Range("F3:F1000").ClearContents()
    rz.Range("H3:H1000").ClearContents()
    for c in ("D", "E", "F", "I", "J", "K", "L", "N", "P", "Y"):
        odn.Range(f"{c}3:{c}1000").ClearContents()
    for i, d in enumerate(spec["lines"]):
        r = SRC0 + i
        for c, v in d.items():
            src.Range(f"{c}{r}").Value = v
        q = ODN0 + i
        rz.Range(f"F{q}").Value = board
        rz.Range(f"H{q}").Value = d["D"]
        odn.Range(f"E{q}").Value = "авто"
        if not str(d["D"]).startswith("Ввод"):              # у ввода клеммы XT нет
            odn.Range(f"P{q}").Value = f"XT{i + 1}"
    net.Range("C65").Value = spec.get("inputs", 1)
    net.Range("C66").Value = spec.get("avr", "нет")
    if board == "ГРЩ":
        for a, v in GRSH_POWER:
            net.Range(a).Value = v
        # Ввод 2 — ТП-2, трансформатор и кабель как у Ввода 1
        net.Range("W6").Formula = "=C33"
        net.Range("X6").Formula = "=C34-R_пер"
        net.Range("Y6").Formula = "=C36"
        net.Range("Z6").Value = 160
        net.Range("C42").Value = 25                     # вводные и отходящие ВА ГРЩ: Icu 25 кА
        net.Range("AA6").Value = "ТП-2, Т2 630 кВА (как ТП-1)"
    if opts:
        odn.Protect(PWD, *opts)
    app.Calculation = -4105
    app.CalculateFull()
    wb.Save()
    wb.Close(False)


SUBDIRS = {"ЩР-1": "Корпус 1", "ЩК-1": "Корпус 1", "ЩК-2": "Корпус 1", "ЩК-3": "Корпус 1",
           "ЩС-1": "Инженерия", "ЩО-1": "Инженерия"}


def run(folder, subdirs=False):
    folder = os.path.abspath(folder)
    os.makedirs(folder, exist_ok=True)
    paths = {}
    for board in BOOKS:
        d = os.path.join(folder, SUBDIRS[board]) if subdirs and board in SUBDIRS else folder
        os.makedirs(d, exist_ok=True)
        paths[board] = os.path.join(d, f"{board}.xlsx")
        shutil.copyfile(TPL, paths[board])
    app, pid = excel()
    try:
        for board, spec in BOOKS.items():
            fill_book(app, paths[board], board, spec)
            print(f"  {board}: линий {len(spec['lines'])}, вводов {spec.get('inputs', 1)}")
        ad = app.Workbooks.Open(ADDIN)
        pre = "'" + ad.Name + "'!"
        app.Run(pre + "SX_SetQuiet", True)
        # все книги открыты вместе: связи считаются вживую
        wbs = {b: app.Workbooks.Open(p, 0, False) for b, p in paths.items()}
        for child, up, feeder, inp in LINKS:
            wbs[child].Activate()
            msg = app.Run(pre + "SX_LinkUpstreamTo", paths[up], feeder, 1, inp)
            print("  " + msg.split("\n")[0])
        for _ in range(3):                      # снизу вверх и сверху вниз — до устойчивых чисел
            app.CalculateFull()
        for w in wbs.values():
            w.Save()
        wbs["ГРЩ"].Activate()
        app.Run(pre + "SX_Tree")
        print("  " + app.Run(pre + "SX_GetLastMsg"))
        wbs["ГРЩ"].Save()
        t = wbs["ГРЩ"].Worksheets("Дерево_щитов")
        r = 4
        while t.Cells(r, 2).Value:
            print("   ", " | ".join(str(t.Cells(r, j).Text) for j in (1, 3, 4, 5, 6, 7, 8)))
            r += 1
        wbs["ГРЩ"].Activate()
        app.Run(pre + "SX_PZ")                  # ПЗ по объекту: ПЗ_<папка>.xlsx и .docx
        print("  " + app.Run(pre + "SX_GetLastMsg"))
        for w in list(app.Workbooks):
            w.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        try:
            os.kill(pid, signal.SIGTERM)
        except Exception:
            pass
    print("папка:", folder)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    run(sys.argv[1], "--подпапки" in sys.argv[2:])
