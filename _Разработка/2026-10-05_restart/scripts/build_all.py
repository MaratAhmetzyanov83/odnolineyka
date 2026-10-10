# -*- coding: utf-8 -*-
"""Перевести таблицу 1.05 (или 1.06 / 1.07) в последнюю версию — все патчи по очереди в одном Excel.

py build_all.py <таблица.xlsx> <результат.xlsx>
Без аргументов — собирает чистую таблицу: 1_Таблицы_Excel/1.05 эталон -> 1_Таблицы_Excel/Однолинейка 1.08 чистая.xlsx
Исходный файл не меняется.
"""
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import patch_106  # noqa: E402
import patch_107  # noqa: E402
import patch_108  # noqa: E402
import patch_layout  # noqa: E402
import patch_order  # noqa: E402

ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
PATCHES = [patch_106.patch, patch_107.patch, patch_108.patch, patch_order.patch, patch_layout.layout]
VERSION = "1.08"


def build(src, dst, log=print):
    src, dst = os.path.abspath(src), os.path.abspath(dst)
    shutil.copyfile(src, dst)
    app = patch_106.excel()
    try:
        wb = app.Workbooks.Open(dst, UpdateLinks=0)
        app.Calculation = -4135
        for p in PATCHES:
            p(wb, log=log)
        app.Calculation = -4105
        app.CalculateFull()
        wb.Worksheets(1).Activate()
        wb.Save()
        wb.Close(False)
    finally:
        patch_106.quit_excel(app)
    return dst


if __name__ == "__main__":
    if len(sys.argv) >= 3:
        build(sys.argv[1], sys.argv[2])
    else:
        build(os.path.join(ROOT, "1_Таблицы_Excel", "1.05 эталон", "Однолинейка 1.05 чистая.xlsx"),
              os.path.join(ROOT, "1_Таблицы_Excel", f"Однолинейка {VERSION} чистая.xlsx"))
