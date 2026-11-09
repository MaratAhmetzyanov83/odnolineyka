"""Копирует рабочие файлы из проекта (2_AutoCAD, 4_Инструкции) в SX_Tools. Запуск: py sync_from_project.py"""
import shutil, sys
from pathlib import Path

here = Path(__file__).resolve().parent
proj = here.parent
src = proj / "2_AutoCAD"
cad = here / "CAD"

lisp = {
    "Однолинейка": ["SX_Schema.lsp", "SX_Schema.dcl"],
    "Компоновка": ["SX_Panel.lsp", "SX_Panel.dcl", "SX_Корпуса.csv"],
    "Длины_трасс": ["DLN.lsp"],
    "Блоки": ["SX_Каталог.lsp", "Каталог_блоков.csv", "Размеры_для_проверки.csv",
              "SX_База_блоков.dwg", "02_Однолинейка.dwg", "SX_Компоновка_пользовательские.dwg"],
}
tools = {"Кабельный_журнал": [("CableJournal_1.4.exe", "CableJournal.exe"), ("КЖ_шаблон.dwg", "КЖ_шаблон.dwg")]}
docs = ["Программа_SXDRAW.md", "Компоновка_щита.md", "Кабельный_журнал.md", "База_блоков_компоновки.md",
        "Как_заполнять_строки.md", "Категории_режимы_ИБП.md", "Спецификация_и_изменения.md", "Надстройка_на_ленте.md"]

n = 0
def cp(a, b):
    global n
    if not a.exists():
        print("нет файла:", a); return
    b.parent.mkdir(parents=True, exist_ok=True)
    if not b.exists() or a.stat().st_mtime > b.stat().st_mtime or a.stat().st_size != b.stat().st_size:
        shutil.copy2(a, b); n += 1; print("копия:", b.relative_to(here))

for d, files in lisp.items():
    for f in files:
        cp(src / d / f, cad / d / f)
for d, files in tools.items():
    for a, b in files:
        cp(src / d / a, here / "Tools" / b)
for f in docs:
    cp(proj / "4_Инструкции" / f, here / "Docs" / f)
print("Обновлено файлов:", n)
