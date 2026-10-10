# -*- coding: utf-8 -*-
"""methodology_sheet.py - лист «Методика» из 4_Инструкции/Методика_расчёта.md (Markdown -> ячейки Excel).

    py methodology_sheet.py <книга.xlsx> [--md <Методика_расчёта.md>]

Поддерживается: # / ## / ### заголовки, абзацы, списки (-, 1.), цитаты-примечания (>), таблицы (| … |), **жирный** снимается.
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "2026-09-29_speedup-automation", "scripts"))
from model_v312 import ExcelSession, log  # noqa: E402

MD = os.path.join(os.path.dirname(__file__), "..", "..", "..", "4_Инструкции", "Методика_расчёта.md")
C_SEC = 14857357
C_HEAD = 14277081
C_NOTE = 15921906
C_WARN = 10284031


def clean(t):
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", t)
    t = re.sub(r"`(.+?)`", r"\1", t)
    t = re.sub(r"<(https?://[^>]+)>", r"\1", t)
    return t.strip()


def render(wb, md=MD, name="Методика"):
    try:
        wb.Worksheets(name).Delete()
    except Exception:
        pass
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Cells.VerticalAlignment = -4160
    widths = {"A": 3, "B": 30, "C": 34, "D": 22, "E": 34, "F": 22}
    for c, w in widths.items():
        ws.Columns(c).ColumnWidth = w
    lines = open(md, encoding="utf-8").read().split("\n")
    r = 1
    i = 0
    while i < len(lines):
        ln = lines[i].rstrip()
        if not ln.strip():
            i += 1
            continue
        if ln.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                cells = [clean(c) for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{2,}:?", c.strip()) for c in cells if c.strip()):
                    rows.append(cells)
                i += 1
            for k, cells in enumerate(rows):
                for j, c in enumerate(cells[:5]):
                    cell = ws.Cells(r, 2 + j)
                    cell.Value = c
                    cell.WrapText = True
                    for b in (7, 8, 9, 10):
                        cell.Borders(b).LineStyle = 1
                if k == 0:
                    rng = ws.Range(ws.Cells(r, 2), ws.Cells(r, 1 + min(5, len(cells))))
                    rng.Font.Bold = True
                    rng.Interior.Color = C_HEAD
                r += 1
            r += 1
            continue
        rng = ws.Range(ws.Cells(r, 2), ws.Cells(r, 6))
        rng.Merge()
        rng.WrapText = True
        if ln.startswith("# "):
            rng.Value = clean(ln[2:])
            rng.Font.Size = 16
            rng.Font.Bold = True
        elif ln.startswith("## ") or ln.startswith("### "):
            r += 1 if ln.startswith("## ") else 0
            rng.UnMerge()
            rng = ws.Range(ws.Cells(r, 2), ws.Cells(r, 6))
            rng.Merge()
            rng.Value = clean(ln.lstrip("#"))
            rng.Font.Bold = True
            rng.Font.Size = 12 if ln.startswith("## ") else 11
            if ln.startswith("## "):
                rng.Interior.Color = C_SEC
        elif ln.startswith(">"):
            t = clean(ln.lstrip("> "))
            rng.Value = t
            rng.Interior.Color = C_WARN if t.startswith("ВНИМАНИЕ") or t.startswith("ОСТОРОЖНО") else C_NOTE
        elif re.match(r"^\s*(-|\d+\.)\s", ln):
            rng.Value = "•  " + clean(re.sub(r"^\s*(-|\d+\.)\s", "", ln)) if ln.lstrip().startswith("-") else clean(ln)
        else:
            rng.Value = clean(ln)
        n = len(str(rng.Cells(1, 1).Value or ""))
        ws.Rows(r).RowHeight = max(13.5, 13.5 * -(-n // 135))
        r += 1
        i += 1
    try:
        ws.Activate()
        wb.Application.ActiveWindow.DisplayGridlines = False
    except Exception:
        pass
    return r


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("book")
    ap.add_argument("--md", default=MD)
    a = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    with ExcelSession() as xl:
        wb = xl.Workbooks.Open(os.path.abspath(a.book), 0, False)
        n = render(wb, a.md)
        wb.Save()
        wb.Close(False)
    log("методика: %d строк" % n)


if __name__ == "__main__":
    main()
