# -*- coding: utf-8 -*-
"""Сравнение 1.05 -> 1.06 на реальных таблицах (копиях).

py compare_106.py <папка с таблицами 1.05> <отчёт.xlsx> [папка для копий]
Каждая таблица копируется, пересчитывается в своём скрытом Excel до и после patch_106.patch(),
отчёт — какие линии поменяли автомат, сечение, потери; итоги щитов.
"""
import glob
import os
import shutil
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import patch_106  # noqa: E402


def num(x):
    return x if isinstance(x, (int, float)) and not isinstance(x, bool) and abs(x) < 1e12 and x > -1e9 else None


def show(x):
    if x is None:
        return ""
    if isinstance(x, float):
        return round(x, 2)
    if isinstance(x, int) and x < -2000000000:
        return "#ошибка"
    return x


def snap(wb):
    app = wb.Application
    app.CalculateFull()
    src = wb.Worksheets("Исходные данные").Range("A3:AT1000").Value
    lines = {}
    for i, r in enumerate(src):
        d = r[3]
        if d in (None, "") or not isinstance(d, str):
            continue
        lines[(d, i)] = {
            "name": r[35], "E": r[4], "F": r[5], "H": r[7], "O": r[14], "M": r[12],
            "T": r[19], "U": r[20], "V": r[21], "AB": r[27], "AO": r[40],
        }
    rz = wb.Worksheets("Разбивка_по_щитам").Range("E3:M1000").Value
    phases = {i: (r[3], r[8]) for i, r in enumerate(rz) if r[3]}
    nz = wb.Worksheets("Нагрузка_щитов")
    tot = {
        "щит": nz.Range("R4").Value,
        "Руст": nz.Range("V5").Value, "Рр": nz.Range("V6").Value, "Iр": nz.Range("V7").Value,
        "cos": nz.Range("V8").Value, "Кс": nz.Range("V9").Value,
    }
    return lines, phases, tot


def reason(a, b):
    out = []
    f = num(a["F"])
    if f and f <= 50:
        out.append(f"линия {int(f)} В — ток и потери на своём напряжении")
    if a["V"] != b["V"]:
        out.append("сечение по ПУЭ 1.3.6 по числу нагруженных жил, In ≤ Iдоп")
    if a["U"] != b["U"] and not (f and f <= 50):
        out.append("ток с cos φ и 230/400 В")
    if a["AB"] != b["AB"]:
        out.append("потери: итоговая длина, cos φ, ρ 0,0225")
    return "; ".join(out)


def main(folder, report, workdir=None):
    workdir = workdir or os.path.join(os.path.dirname(os.path.abspath(report)), "_копии")
    os.makedirs(workdir, exist_ok=True)
    files = [f for f in sorted(glob.glob(os.path.join(folder, "*.xlsx")))
             if not os.path.basename(f).startswith("~$")]
    rows, totals, summary = [], [], []
    app = patch_106.excel()
    try:
        for f in files:
            base = os.path.basename(f)
            dst = os.path.join(workdir, base)
            shutil.copyfile(f, dst)
            wb = app.Workbooks.Open(os.path.abspath(dst), UpdateLinks=0)
            try:
                ws_names = [ws.Name for ws in wb.Worksheets]
                if "Исходные данные" not in ws_names:
                    continue
                l0, p0, t0 = snap(wb)
                patch_106.patch(wb, log=lambda *_: None)
                l1, p1, t1 = snap(wb)
                wb.Save()
            finally:
                wb.Close(False)
            n_u = n_v = n_ab = n_ph = 0
            for key, a in l0.items():
                b = l1.get(key)
                if not b:
                    continue
                du = a["U"] != b["U"]
                dv = a["V"] != b["V"]
                ab0, ab1 = num(a["AB"]), num(b["AB"])
                dab = (ab0 is None) != (ab1 is None) or (ab0 is not None and abs(ab1 - ab0) >= 0.5)
                if du:
                    n_u += 1
                if dv:
                    n_v += 1
                if dab:
                    n_ab += 1
                if du or dv or dab:
                    rows.append([base, key[0], show(a["name"]), show(a["F"]), show(a["H"]), show(a["E"]),
                                 show(a["M"]), show(a["T"]), show(b["T"]), show(a["U"]), show(b["U"]),
                                 show(a["V"]), show(b["V"]), show(a["AB"]), show(b["AB"]), reason(a, b)])
            for i, (e, ph0) in p0.items():
                ph1 = p1.get(i, (None, None))[1]
                if ph0 != ph1:
                    n_ph += 1
            totals.append([base, show(t0["щит"]), show(t0["Руст"]), show(t1["Руст"]), show(t0["Рр"]),
                           show(t1["Рр"]), show(t0["Iр"]), show(t1["Iр"]), show(t0["cos"]), show(t0["Кс"])])
            summary.append([base, len(l0), n_u, n_v, n_ab, n_ph])
            print(base, "линий", len(l0), "автомат", n_u, "сечение", n_v, "потери", n_ab, "фаза", n_ph, flush=True)
    finally:
        patch_106.quit_excel(app)

    out = openpyxl.Workbook()
    head = Font(bold=True)
    fill = PatternFill("solid", fgColor="FFF2CC")
    s = out.active
    s.title = "Итог"
    s.append(["Сравнение 1.05 → 1.06 на копиях таблиц: " + folder])
    s.append(["Исходные файлы не менялись. Изменения — в листе «Линии». Потери показаны, если разница ≥ 0,5 %."])
    s.append([])
    s.append(["Файл", "Линий", "Автомат поменялся", "Сечение поменялось", "Потери (≥0,5 %)", "Фаза в «Разбивке»"])
    for c in s[4]:
        c.font = head
    for r in summary:
        s.append(r)
    s.append([])
    s.append(["Что исправлено в 1.06 — «Краткое руководство», строка 1.06."])
    s.column_dimensions["A"].width = 22
    for col in "BCDEF":
        s.column_dimensions[col].width = 18

    l = out.create_sheet("Линии")
    l.append(["Файл", "Линия", "Наименование", "U, В", "Фаз", "P, кВт", "Длина итог, м", "Ток было", "Ток стало",
              "Автомат было", "Автомат стало", "Сечение было", "Сечение стало", "Потери % было", "Потери % стало",
              "Почему"])
    for c in l[1]:
        c.font = head
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in rows:
        l.append(r)
        i = l.max_row
        for a, b in ((10, 11), (12, 13), (14, 15)):
            if l.cell(i, a).value != l.cell(i, b).value:
                l.cell(i, b).fill = fill
    widths = [16, 16, 40, 7, 6, 8, 10, 9, 9, 10, 10, 10, 10, 10, 10, 60]
    for i, w in enumerate(widths, 1):
        l.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    l.freeze_panes = "C2"
    l.auto_filter.ref = f"A1:P{l.max_row}"

    t = out.create_sheet("Итоги щитов")
    t.append(["Файл", "Щит", "Руст было", "Руст стало", "Рр было", "Рр стало", "Iр было", "Iр стало", "cos φ", "Кс"])
    for c in t[1]:
        c.font = head
    for r in totals:
        t.append(r)
    t.column_dimensions["A"].width = 18
    out.save(report)
    print("отчёт:", report)


if __name__ == "__main__":
    main(*sys.argv[1:])
