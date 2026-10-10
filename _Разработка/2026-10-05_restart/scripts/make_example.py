# -*- coding: utf-8 -*-
"""Пример 1.08: чистая таблица + демо-щит ЩР (ввод в узле, розетки с «гр», свет Relay/DIM, LED 24 В с БП, шторы, 3ф, тёплый пол).
py make_example.py [чистая 1.08.xlsx] [пример.xlsx]"""
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import demo_fill  # noqa: E402
import patch_106  # noqa: E402

ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "1_Таблицы_Excel", "Однолинейка 1.08 чистая.xlsx")
dst = sys.argv[2] if len(sys.argv) > 2 else os.path.join(ROOT, "1_Таблицы_Excel", "Однолинейка 1.08 пример.xlsx")
shutil.copyfile(src, dst)
app = patch_106.excel()
try:
    wb = app.Workbooks.Open(os.path.abspath(dst), UpdateLinks=0)
    demo_fill.fill(wb)
    wb.Worksheets("Экспликация").Range("A3:C4").Value = (("1", "01", "Гостиная"), ("1", "02", "Кухня"))
    app.CalculateFull()
    wb.Worksheets("Разбивка_по_щитам").Activate()
    wb.Save()
    wb.Close(False)
finally:
    patch_106.quit_excel(app)
print("пример:", dst)
