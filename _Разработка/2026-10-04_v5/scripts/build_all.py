# -*- coding: utf-8 -*-
"""build_all.py - собрать четыре книги v5 (шаблон, пример «все случаи», ЖК «Остров», ЖД) в ../out; журналы — ../out/build_*.log."""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
T = os.path.join(ROOT, "1_Таблицы_Excel")
OUT = os.path.join(HERE, "..", "out")
V4 = os.path.join(HERE, "..", "..", "2026-10-03_v4")
ZHD_BASE = os.environ.get("ZHD_BASE") or os.path.join(HERE, "..", "base", "Однолинейка ГРЩ-ЖД.xlsx")
JOBS = [("demo", os.path.join(T, "Однолинейка_шаблон.xlsx"), "Однолинейка_v5_пример_все_случаи.xlsm", ["--defaults"]),
        ("tpl", os.path.join(T, "Однолинейка_шаблон.xlsx"), "Однолинейка_v5_шаблон.xlsm", ["--defaults"]),
        ("ostrov", os.path.join(T, "Однолинейка_пример_ЖК-Остров.xlsx"), "Однолинейка_v5_пример_ЖК-Остров.xlsm", []),
        ("zhd", ZHD_BASE, "Однолинейка_v5_ЖД.xlsm", [])]
only = sys.argv[1:]
for name, base, out, extra in JOBS:
    if only and name not in only:
        continue
    log = os.path.join(OUT, "build_%s.log" % name)
    cmd = [sys.executable, os.path.join(HERE, "build_v5.py"), "--data", os.path.join(V4, "data_%s.json" % name), "--base", base,
           "--out", os.path.join(OUT, out)] + extra
    with open(log, "w", encoding="utf-8") as fh:
        subprocess.run(cmd, stdout=fh, stderr=subprocess.STDOUT, env=dict(os.environ, PYTHONIOENCODING="utf-8"))
    txt = open(log, encoding="utf-8").read()
    keep = [ln for ln in txt.splitlines() if any(k in ln for k in ("ошибки", "ЦИКЛ", "ОШИБКА", "сохранено", "кнопка", "срез", "СВЕРКА", "Traceback", "Error"))]
    print("==", out)
    print("\n".join(keep))
