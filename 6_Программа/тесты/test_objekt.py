# -*- coding: utf-8 -*-
"""Проверка программы на тестовом объекте (make_demo.py).

    python тесты/test_objekt.py [папка]   — папки нет: создаст во временной (≈30 с)
"""
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)

from odn.project import load_project  # noqa: E402


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else os.path.join(tempfile.gettempdir(), "odn_demo")
    if not os.path.isfile(os.path.join(folder, "ГРЩ.xlsx")):
        import make_demo
        make_demo.main(folder)
    prj = load_project(folder, progress=lambda s: None, use_cache=False)
    panels = {p.name: p for p in prj.panels()}
    obj = [(i.level, panels_name(prj, i.panel), i.line, i.text.split(":")[0]) for i in prj.issues if i.source == "объект"]
    want = [
        (2, "ГРЩ", "Щ.1", "нагрузка ЩР-1 в книге ГРЩ устарела"),
        (2, "ЩР-2", "Щ.2", "связь устарела"),
        (1, "ГРЩ", "Щ.1", "селективность"),
        (1, "ГРЩ", "Щ.2", "селективность"),
        (1, "ГРЩ", "Щ.3", "фидер «Питание ЩР-3»"),
    ]
    fails = []
    if sorted(panels) != ["ГРЩ", "ЩР-1", "ЩР-2"]:
        fails.append(f"щиты: {sorted(panels)} (копия в _резерв не должна попасть)")
    if panels["ЩР-1"].parent != panels["ГРЩ"].id or panels["ЩР-1"].parent_line != "Щ.1":
        fails.append("ЩР-1 не связан с ГРЩ / Щ.1")
    if panels["ЩР-2"].parent != panels["ГРЩ"].id or panels["ЩР-2"].parent_line != "Щ.2":
        fails.append("ЩР-2 (подпапка, путь ..\\) не связан с ГРЩ / Щ.2")
    if sorted(obj) != sorted(want):
        fails.append(f"замечания объекта:\n  есть  {sorted(obj)}\n  нужно {sorted(want)}")
    xl = [i for i in prj.issues if i.source == "Excel" and i.level == 2 and panels_name(prj, i.panel) == "ЩР-1"]
    if len(xl) != 6:
        fails.append(f"ошибок Excel в ЩР-1: {len(xl)}, в книге 6")
    print("ОК" if not fails else "ОШИБКИ:\n" + "\n".join(fails))
    return 1 if fails else 0


def panels_name(prj, pid):
    return next((p.name for p in prj.panels() if p.id == pid), pid)


if __name__ == "__main__":
    sys.exit(main())
