# -*- coding: utf-8 -*-
"""make_demo.py - данные книги-примера «все случаи» (формат extract_v3.py) -> data_demo.json.

    py make_demo.py            -> ../data_demo.json
    py build_v4.py --data ../data_demo.json --base "<шаблон v3.12>" --out "<Однолинейка_v4_пример_все_случаи.xlsx>"

Объект из 5 щитов, показывает каждую возможность книги (см. колонку «что показывает» в списке ниже).
Справочники и настройки — из шаблона (data_tpl.json) + код «НВК» (насос).
"""
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
F = "demo"

tpl = json.load(open(os.path.join(HERE, "..", "data_tpl.json"), encoding="utf-8"))
ref = tpl["ref"]
ref["codes"]["НВК"] = ["НВК", "Насос", 1, 1, 2, "мс", 3, 230, 1, None, "Насосы (водоснабжение, дренаж)", "Сил", "нет"]
ref["ugo"]["НВК"] = ["НВК", "Привод (двигатель)", "нет", None, None]

boards = [
    dict(name="ВРУ", nin=4, scheme=dict(inputs=2, avr="да", ups_from="Ввод 1"),
         manual={"from": "сеть", "ikz_min_up": 2500, "ikz_max_up": 8, "s": 95, "cores": 4, "n_par": 1, "l": 60, "mat": "медь",
                 "qf_in": 250, "qf_char": "C"}),
    dict(name="ЩР-1", nin=3, scheme=dict(inputs=1, avr="нет", ups_from="Ввод 1"), manual={}),
    dict(name="ЩС", nin=2, scheme=dict(inputs=1, avr="нет", ups_from="Ввод 1"), manual={}),
    dict(name="ЩГП", nin=2, scheme=dict(inputs=1, avr="нет", ups_from="Ввод 1"), manual={}),
    dict(name="ЩСПЗ", nin=3, scheme=dict(inputs=2, avr="да", ups_from="Ввод 1"), manual={}),
    dict(name="ЩБП", nin=2, scheme=dict(inputs=1, avr="нет", ups_from="Ввод 1"), manual={}),
]

# (щит, место, номер линии, ввод схемы, ввод линии)  — место None = следующее по порядку
P = []


def add(board, pos, line, sch=None, ln=None):
    P.append((board, pos, line, sch or {}, ln or {}))


V3 = dict(ph=3, u=400, cable="мс", cores=5, kc_group="Нет")
# ---- ВРУ: вводной узел (места 1–4) и фидеры
add("ВРУ", 1, "Ввод.1", dict(mode="вручную"), dict(V3, s=95, name="Ввод 1 от ТП"))                     # узел, только таблица
add("ВРУ", 3, "Ввод.2", dict(mode="вручную"), dict(V3, s=95, name="Ввод 2 от ТП"))
add("ВРУ", 4, "АСУ.0.001.1", dict(mode="рисовать", qf_txt="авто"), dict(p=0.1, l_plan=2, name="Питание автоматики АВР"))  # строка узла, рисуется
add("ВРУ", None, "ЩР-1", dict(qf_txt="авто"), dict(ph=3, u=400, cable="мс", cores=5, l_plan=25, cat="II"))     # фидер: номер = имя щита
add("ВРУ", None, "ЩС", dict(qf_txt="авто"), dict(ph=3, u=400, cable="мс", cores=5, l_plan=40, cat="II"))
add("ВРУ", None, "ЩГП", dict(qf_txt="авто"), dict(ph=3, u=400, cable="мс", cores=5, l_plan=15, cat="I"))
add("ВРУ", None, "ЩСПЗ.1", dict(qf_txt="авто", feeds="ЩСПЗ"), dict(ph=1, u=230, cable="мфр", cores=3, l_plan=12, cat="I", fire="да",
                                                                 name="Щит СПЗ, ввод 1"))                       # фидер вручную
add("ВРУ", None, "ЩСПЗ.2", dict(qf_txt="авто"), dict(ph=1, u=230, cable="мфр", cores=3, l_plan=12, cat="I", fire="да",
                                                    kc_group="Нет", name="Щит СПЗ, ввод 2 (резерв)"))
add("ВРУ", None, "КМН.0.001.1", dict(qf_txt="авто"), dict(p=2.0, kc_group="Отоп", section=2, l_plan=18))           # отопление, секция 2
add("ВРУ", None, "ПН.0.002.1", dict(qf_txt="авто"), dict(p=5.0, kc_group="ВНГ", qf=32, s=6, cable="мс", cores=3, l_plan=10))  # ВНГ, свой автомат/сечение
add("ВРУ", 20, "R.0.004.1", dict(qf_txt="авто"), dict(p=1.5, n=2, l_plan=12, name="Розетки техпомещения"))       # своё место: пропуск колонок
add("ВРУ", "-", "LD.0.003.1", {}, dict(p=0.3, l_plan=5))                                                        # нет на схеме
# ---- ЩР-1
add("ЩР-1", 1, "Ввод.1", dict(mode="вручную"), dict(V3, s=10))
add("ЩР-1", None, "R.1.01.1", dict(qf_txt="авто"), dict(p=1.8, n=4, l_plan=22, floor=2))                        # своя линия, дифзащита по коду
add("ЩР-1", None, "R.1.02.1", dict(qf_txt="авто"), dict(p=1.5, n=3, l_plan=15, floor=2))                        # группа: 2 линии на 1 автомате
add("ЩР-1", None, "R.1.02.2", {}, dict(p=1.0, n=2, l_plan=12, floor=2))
add("ЩР-1", None, "R.1.03.1", dict(qf_txt="авто", rcd=10), dict(p=2.0, l_plan=9, name="Розетки ванной (дифзащита 10 мА)"))
add("ЩР-1", None, "L.1.01.1", dict(qf_txt="авто", module="ZIOMB16V3"), dict(p=0.3, l_plan=14, floor=1))      # модуль: канал 1
add("ЩР-1", None, "L.1.02.1", {}, dict(p=0.2, l_plan=16, floor=1))                                             # канал 2
add("ЩР-1", None, "L.1.02.2", dict(ch="="), dict(p=0.1, l_plan=6, floor=1))                                   # параллельно на тот же канал
add("ЩР-1", None, None, dict(ch="р"))                                                                       # резерв канала
add("ЩР-1", None, "L.1.03.1", {}, dict(p=0.25, l_plan=11, floor=1))                                          # последний канал модуля
add("ЩР-1", None, "LED.1.01.1", dict(qf_txt="авто", dev1="БП 24V 75W", module="-"), dict(p=0.07, l_plan=8, cable="мс", cores=3, s=1.5))  # «-» = модуль закончился; БП после автомата
add("ЩР-1", None, "ТП.1.01.1", dict(qf_txt="авто", dev1="КМ NCH8-40/40"), dict(p=2.5, kc_group="ВНГ", l_plan=14, cable="мс", cores=3))  # контактор
add("ЩР-1", None, None, dict(ch="+"))                                                                       # ещё клемма в тот же кабель
add("ЩР-1", None, "ПВ.1.01.1", dict(qf_txt="авто"), dict(p=3.0, ph=3, u=400, cable="мс", cores=5, kc=0.8, l_plan=20))  # 3ф, Кс вручную
add("ЩР-1", None, "K.1.01.1", dict(qf_txt="авто"), dict(p=1.2, l_plan=18, cable="мс", cores=3, name="Кондиционер (cos φ вручную)"))
add("ЩР-1", None, "КУП.1.01.1", dict(qf_txt="авто", ch="б/к"), dict(cable="мс", cores=3, s=4, l_plan=6))       # без клеммы
add("ЩР-1", None, "УВ.1.01.1", dict(qf_txt="авто", xt="нет", phase="2"), dict(p=0.4, l_plan=10, cable="мс", cores=3))  # клемма без номера, фаза вручную
add("ЩР-1", None, "R.1.04.1", dict(qf_txt="авто"), dict(p=2.0, n=1, l_plan=140, name="Розетки гаража (длинная линия)"))
add("ЩР-1", None, "УЗО.1", dict(qf_txt="авто", dev1="УЗО 2P 40А 30мА"), dict(p=0, cable=None, cores=None, kc_group="Нет",
                                                                           name="Шина УЗО1"))             # УЗО на группу автоматов
add("ЩР-1", None, "R.1.05.1", dict(qf_txt="авто", grp="узо"), dict(p=1.5, n=3, l_plan=18, floor=2))         # на шине УЗО: свой автомат
add("ЩР-1", None, "R.1.05.2", dict(qf_txt="авто", grp="узо"), dict(p=1.2, n=2, l_plan=14, floor=2))
add("ЩР-1", None, "L.1.04.1", dict(qf_txt="авто", grp="узо"), dict(p=0.2, l_plan=10, floor=1))
add("ВРУ", None, "ЩБП", dict(qf_txt="авто"), dict(ph=1, u=230, cable="мс", cores=3, l_plan=10, cat="III",
                                                  name="Щит блоков питания LED"))                     # фидер 1ф
# ---- ЩБП: LED 24 В — автомат → МК → БП → DALI-диммер, линии 24 В «гр» под автоматом (постоянный ток, ΔU от БП ≤ 5 %)
add("ЩБП", 1, "Ввод.1", dict(mode="вручную"), dict(ph=1, u=230, cable="мс", cores=3, kc_group="Нет", s=2.5))
L24 = dict(ph=1, u=24, cable="мс", cores=2)
add("ЩБП", None, "LD.1.01.1", dict(qf_txt="авто", dev1="МК-5-1", dev2="БП 24V 480W", module="SMART-DALI-104"),
    dict(L24, p=0.1, l_plan=18, name="LED-лента 24 В, пом. 01"))                 # канал 1
add("ЩБП", None, "LD.1.01.2", {}, dict(L24, p=0.1, l_plan=24, name="LED-лента 24 В, пом. 01"))          # канал 2
add("ЩБП", None, "LD.1.02.1", {}, dict(L24, p=0.12, l_plan=33, name="LED-лента 24 В, пом. 02"))         # канал 3: дальняя — сечение больше
add("ЩБП", None, "LD.1.03.1", {}, dict(L24, p=0.15, l_plan=40, name="LED-лента 24 В, пом. 03"))         # канал 4
add("ЩБП", None, "LD.1.04.1", dict(qf_txt="авто", dev1="МК-5-1", dev2="БП 24V 75W", module="-"),
    dict(L24, p=0.05, l_plan=12, name="LED-подсветка ниши, пом. 04"))           # мал БП: 75 Вт < нужно
add("ЩБП", None, "LD.1.04.2", {}, dict(L24, p=0.04, l_plan=10, name="LED-подсветка ниши, пом. 04"))
# ---- ЩС
add("ЩС", 1, "Ввод.1", dict(mode="вручную"), dict(V3, s=16))
add("ЩС", None, "ПУ.0.01.1", dict(qf_txt="авто"), dict(p=5.5, ph=3, u=400, cable="мс", cores=5, l_plan=30, cat="II"))
add("ЩС", None, "ПВ.0.02.1", dict(qf_txt="авто"), dict(p=4.0, ph=3, u=400, cable="мс", cores=5, l_plan=25, cat="II"))  # Кс по табл. 7.5
add("ЩС", None, "НВК.0.03.1", dict(qf_txt="авто", bus="ДГУ"), dict(p=1.1, l_plan=15, cat="I особая", name="Дренажный насос (шина ДГУ)"))
add("ЩС", None, "НВК.0.04.1", dict(qf_txt="авто"), dict(p=1.1, l_plan=15, kc_group="Нет", name="Дренажный насос резервный"))
add("ЩС", None, "ПУ.0.05.1", dict(qf_txt="авто", bus="Насосная"), dict(p=0.75, l_plan=12, cable="мс", cores=3, cat="II",
                                                                     name="Вентилятор насосной (шина «Насосная» — Одиночная)"))
# ---- ЩГП (шина ИБП)
add("ЩГП", 1, "Ввод.1", dict(mode="вручную"), dict(V3, s=6))
add("ЩГП", None, "АПС.0.01.1", dict(qf_txt="авто", bus="ИБП"), dict(p=0.2, l_plan=8, cable="мс", cores=3, cat="I особая"))
add("ЩГП", None, "L.0.02.1", dict(qf_txt="авто", bus="ИБП"), dict(p=0.3, l_plan=20, cat="I", name="Аварийное освещение (шина ИБП)"))
# ---- ЩСПЗ (2 ввода + АВР, пожарные)
add("ЩСПЗ", 1, "Ввод.1", dict(mode="вручную"), dict(V3, s=2.5))
add("ЩСПЗ", 2, "Ввод.2", dict(mode="вручную"), dict(V3, s=2.5))
add("ЩСПЗ", None, "АПС.0.02.1", dict(qf_txt="авто"), dict(p=0.4, l_plan=10, cable="мфр", cores=3, cat="I", fire="да",
                                                          name="Прибор пожарной сигнализации"))
add("ЩСПЗ", None, "ПУ.0.03.1", dict(qf_txt="авто"), dict(p=3.0, ph=3, u=400, cable="мфр", cores=5, l_plan=35, cat="I", fire="да",
                                                         name="Вентилятор подпора воздуха"))

places, lines = [], []
for board, pos, line, sch, ln in P:
    pl = {"board": board, "pos": None if pos in (None, "-") else pos, "line": line, "file": F}
    if pos == "-":
        pl["unplaced"] = True
    for k, v in sch.items():
        pl[k] = v
    places.append(pl)
    if line:
        rec = {"line": line, "file": F, "board": board}
        if not line.startswith("Ввод") and "ph" not in ln:          # коды шаблона без фаз и кабеля: 1ф, 230 В, ВВГнг 3 жилы
            rec.update(ph=1, u=230, cable="мс", cores=3)
        rec.update(ln)
        lines.append(rec)

data = {"settings": tpl["settings"], "kz": tpl["kz"], "ref": ref, "issues": [], "feeders": [], "buses": [
    {"bus": "Насосная", "board": "ЩС", "type": "Одиночная", "file": F}],
        "boards": [dict(b, file=F, col="C") for b in boards], "places": places, "lines": lines}
out = os.path.join(HERE, "..", "data_demo.json")
json.dump(data, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("щитов", len(boards), "мест", len(places), "линий", len(lines), "->", os.path.abspath(out))
