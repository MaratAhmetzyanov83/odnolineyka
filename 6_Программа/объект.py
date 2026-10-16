# -*- coding: utf-8 -*-
"""Однолинейка — объект в одном окне.

    py объект.py [папка объекта]        — окно в браузере (без папки — спросит, запомнит последнюю)
    py объект.py --json <папка> [файл]  — только выгрузить объект в JSON (для проверок)

Книги Excel только читаются. Нужен Python 3.10+ и openpyxl; для перехода к ячейке в Excel — pywin32.
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
SETTINGS = os.path.join(HERE, "настройки.json")


def settings():
    try:
        with open(SETTINGS, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def ask_folder(last):
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        d = filedialog.askdirectory(title="Папка объекта (книги щитов)", initialdir=last or os.path.expanduser("~"))
        root.destroy()
        return d
    except Exception:  # noqa: BLE001 — нет tkinter
        return input("Папка объекта: ").strip().strip('"')


def main(argv):
    if argv and argv[0] == "--json":
        from odn.project import load_project
        prj = load_project(argv[1])
        out = argv[2] if len(argv) > 2 else os.path.join(argv[1], "объект.json")
        with open(out, "w", encoding="utf-8") as f:
            json.dump(prj.to_dict(), f, ensure_ascii=False, indent=1)
        print("записано:", out, "| замечаний:", len(prj.issues))
        return
    st = settings()
    folder = argv[0] if argv else ask_folder(st.get("last"))
    if not folder or not os.path.isdir(folder):
        print("Папка не выбрана.")
        return
    st["last"] = os.path.abspath(folder)
    try:
        with open(SETTINGS, "w", encoding="utf-8") as f:
            json.dump(st, f, ensure_ascii=False)
    except OSError:
        pass
    from odn.server import serve
    serve(folder)


if __name__ == "__main__":
    main(sys.argv[1:])
