# -*- coding: utf-8 -*-
"""fix_phase_rows.py - исправить сдвиг формул фазы в «Нагрузка_щитов» (v3.12).

    py fix_phase_rows.py <книга.xlsx> [...]

Ошибка шаблона: в строках 29-127 формулы G («Фаза (Итог)») и M (служебная) ссылаются на ручную фазу E следующей
строки (G29 = ...E30...). Из-за этого с 27-й линии щита фаза берётся от соседней линии. Скрипт находит строки,
где G/M ссылаются не на E своей строки, и копирует в них формулы из ближайшей правильной строки выше.
Через Excel COM; строки и столбцы не вставляются. Повторный запуск ничего не меняет.
"""
import os
import re
import sys

from model_v312 import ExcelSession, log

SH = "Нагрузка_щитов"


def bad_rows(ws, col):
    vals = ws.Range(f"{col}3:{col}1000").Formula
    out = []
    for i, (f,) in enumerate(vals):
        r = i + 3
        m = re.search(r"E(\d+)", f or "")
        if m and int(m.group(1)) != r:
            out.append(r)
    return out


def fix(path):
    with ExcelSession() as xl:
        wb = xl.Workbooks.Open(os.path.abspath(path), 0, False)
        try:
            ws = wb.Worksheets(SH)
            prot = ws.ProtectContents
            if prot:
                ws.Unprotect("sx")
            total = 0
            for col in ("G", "M"):
                rows = bad_rows(ws, col)
                for r in rows:
                    src = r - 1
                    while src in rows:
                        src -= 1
                    ws.Range(f"{col}{src}").Copy(ws.Range(f"{col}{r}"))
                total += len(rows)
                left = bad_rows(ws, col)
                if left:
                    log("  ОСТАЛОСЬ %s: %s" % (col, left[:5]))
            if prot:
                ws.Protect("sx", DrawingObjects=True, Contents=True, Scenarios=True, AllowFormattingColumns=True)
            xl.CalculateFull()
            wb.Save()
            log("%s: исправлено формул %d" % (os.path.basename(path), total))
        finally:
            wb.Close(False)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    for p in sys.argv[1:]:
        fix(p)
