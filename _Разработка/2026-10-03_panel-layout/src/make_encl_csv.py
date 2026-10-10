"""База корпусов DKC для SXLAYOUT: data/*.json (собрано с dkc.ru) -> 2_AutoCAD/Компоновка/SX_Корпуса.csv

Запуск: py make_encl_csv.py
CSV: UTF-8 с BOM, разделитель «;», открывается в Excel. Столбцы — см. SX_Layout.src.lsp (sxl:load-encl).
"""
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "..", "data")
OUT = os.path.join(HERE, "..", "..", "..", "2_AutoCAD", "Компоновка", "SX_Корпуса.csv")

HEAD = ["Серия", "Артикул", "Наименование", "Тип", "Установка", "Ширина", "Высота", "Глубина",
        "Рядов", "Модулей в ряду", "Секций", "Шаг реек", "Короб", "Поле", "Примечание"]


def load(name):
    p = os.path.join(DATA, name)
    return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else []


def main():
    rows = []
    # навесные ST / STE с монтажной платой: [серия, артикул, В, Ш, Г, артикул с прозрачной дверью]
    for s, art, h, w, d, glass in load("dkc_st_ste.json"):
        small = h <= 400
        name = f"Корпус навесной {s} с М/П ВxШxГ {h}x{w}x{d} мм, IP66"
        note = f"с прозрачной дверью: {glass}" if glass else ""
        rows.append([s, art, name, "PLATE", "навесной", w, h, d, "", "", 1,
                     125 if small else 150, 25 if small else 40, 25 if small else 40, note])
    # напольные CQE N собранные: [артикул, В, Ш, Г, створок]
    for art, h, w, d, doors in load("dkc_cqen.json"):
        name = f"Напольный собранный корпус CQE N ВxШxГ {h}x{w}x{d} мм" + (", двустворчатая дверь" if doors == 2 else "")
        rows.append(["CQE N", art, name, "CAB", "напольный", w, h, d, "", "", 2 if w >= 1200 else 1, 205, 60, 30, ""])
    # модульные щитки: [серия, артикул, наименование, установка, рядов, модулей всего, Ш, В, Г]
    for s, art, name, inst, nr, nm, w, h, d in load("dkc_modular.json"):
        if not (nr and nm and w and h):
            continue
        rows.append([s, art, name, "MOD", inst, w, h, d, nr, nm // nr, 1, "", "", "", ""])
    with open(OUT, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write("# База корпусов DKC для SXLAYOUT (собрано с dkc.ru 04.10.2026). Можно править в Excel: сохранять как «CSV UTF-8».\n")
        f.write("# Тип: MOD — модульный щиток (рядов x модулей); PLATE — навесной с монтажной платой; CAB — напольный шкаф.\n")
        f.write("# Шаг реек / Короб / Поле, мм — раскладка на монтажной плате: шаг между осями реек, ширина кабель-канала, отступ от края.\n")
        f.write(";".join(HEAD) + "\n")
        for r in rows:
            f.write(";".join(str(x).replace(";", ",") for x in r) + "\n")
    print(len(rows), "rows ->", os.path.abspath(OUT))


if __name__ == "__main__":
    main()
