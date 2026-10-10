# -*- coding: utf-8 -*-
"""build_v4.py - собрать книгу однолинейки v4 (один файл на объект) из данных extract_v3.py.

    py build_v4.py --data ../data_zhd.json --base "<книга v3.12 объекта (ГРЩ)>" --out "<Однолинейка_v4_объект.xlsx>"

Основа - копия книги v3.12 (оттуда берутся «Справочная», «Справочная_потребители», «Каталог», «Спецификация»,
«Спецификация_авто», «Бирки»); старые листы ввода и расчёта удаляются. Новые листы:
  «Объект»  - настройки объекта и ТКЗ;
  «Щиты»    - таблица щитов: схема, путь до ГРЩ, нагрузка по составу (вверх), ΔU и Zпетли (вниз);
  «Линии»   - таблица: строка = место схемы, ввод слева, расчёт справа;
  «Однолинейка» - расчётный лист с раскладкой v3.12 (интерфейс для SX_Schema.lsp, SX_Layout.lsp, «Каталога», «Бирок»).
Режим расчёта - совместимость с v3.12 (формулы перенесены 1:1, см. 4_Инструкции/Методика_расчёта.md).
"""
import argparse
import json
import os
import re
import shutil
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "2026-09-29_speedup-automation", "scripts"))
from model_v312 import ExcelSession, log  # noqa: E402

PW = "sx"
DELETE = ["Гайд", "Краткое руководство", "Исходные данные", "Разбивка_по_щитам", "Питание_щита", "Нижестоящие_щиты",
          "Нагрузка_щитов", "Режимы", "ИБП", "Категории_по_щитам", "Расчёт_нагрузок", "ТКЗ", "Кабельный_журнал",
          "В Акад", "Диапазоны"]
R0 = 3                       # первая строка данных таблиц (строка 2 - заголовки, как ищет DLN.lsp)

C_IN = 13431551              # светло-жёлтый - ввод
C_HEAD = 14277081
C_CALC = 15921906            # светло-серый - расчёт
C_GRP = 14599344             # голубой - заголовки групп

# ------------------------------------------------------------------------------------------------
# Таблица «Линии»: (ключ, заголовок, вид, формула). Вид: in - ввод, c - расчёт (видимый), h - служебный.
# Плейсхолдеры формул: {k} - буква столбца ключа k (строка 3 = первая строка данных); {P:k} - значение строки выше
# (пусто в первой строке); {NB} - новая группа строк (первая строка или другой щит); {T:k} - Линии[заголовок k].
# ------------------------------------------------------------------------------------------------
U_F = "(Uн_лин/SQRT(3))"      # кВ
POT = "Потребители_обозначение"
L = []


def col(key, head, kind, f=None, note=None, grp=None):
    L.append(dict(key=key, head=head, kind=kind, f=f, note=note, grp=grp))


# --- ввод линии
G1 = "ЛИНИЯ (ввод)"
col("board", "Щит", "in", grp=G1, note="Имя щита. Все места щита — подряд, в порядке мест на схеме.")
col("pos", "№ места", "in", grp=G1, note="Колонка на схеме (шаг 20 мм). Пусто — следующее место по порядку (после вводного узла). Число — своё место (пропуск = пустая колонка). «-» — строки нет на схеме.")
col("place_lbl", "Лист / место", "c", '=IF({pos_eff}3="","",IF({zone}3="вводной узел","узел ","")&"л."&(INT(({pos_eff}3-1)/Столбцов_на_лист)+1)&" / "&(MOD({pos_eff}3-1,Столбцов_на_лист)+1))', grp=G1,
    note="Лист схемы A3 и место на нём (мест на лист — «Справочная»/«Таблицы_норм» BF4). «узел» — вводной узел щита.")
col("line", "Номер линии", "in", grp=G1, note="Код.Помещение.№ (пример: R.1.05.4). Код — «Справочная» P.")
col("name_m", "Наименование", "in", grp=G1, note="Подставляется по коду автоматически. Можно переписать своим текстом. Если стереть — снова по коду (формулу можно протянуть из соседней ячейки).")
col("feeds", "Питает щит", "in", grp=G1, note="Обычно не нужно: если номер линии совпадает с именем щита на листе «Щиты», фидер определяется сам. Заполняйте, только если номер другой (пример: линия ЩСПЗ.1 питает щит ЩСПЗ).")
col("p_in", "Мощность, кВт", "in", grp=G1)
col("n", "Кол-во потребителей (для длины)", "in", grp=G1)
col("ph_in", "Кол-во фаз", "in", grp=G1, note="1 или 3. Пусто — по коду.")
col("u_in", "Напряжение", "in", grp=G1, note="230 или 400. Пусто — по коду.")
col("grp_in", "Группа Кс (Быт/Сил/Пост/Отоп/ВНГ/Нет; пусто = по коду)", "in", grp=G1, note="Быт — табл. 7.2 по сумме; Сил — Кс линии или табл. 7.5; Пост — 1; Отоп — 1 и ВНГ — 0,7 отдельно (СП 256, Изм. 9); Нет — не учитывается.")
col("kc_in", "Кс (для Сил)", "in", grp=G1, note="Кс силовой линии (СП 256 табл. 7.5). Пусто — 1.")
col("cos_in", "cos φ (ручн.)", "in", grp=G1, note="Пусто — 0,95, для двигателей по таблице.")
col("sec_in", "Секция (1/2/3; пусто = 1)", "in", grp=G1)
col("phase_in", "Фаза (ручная)", "in", grp=G1, note="1, 2, 3 или 1,2,3. Пусто — по кругу L1→L2→L3.")
G2 = "КАБЕЛЬ (ввод)"
col("lplan_in", "Длина линии (по плану)", "in", grp=G2, note="м. Пусто — сумма длин из чертежа (DWG).")
col("tray", "Лоток", "in", grp=G2)
col("pvc_in", "Гофра ПВХ", "in", grp=G2)
col("pnd_in", "Гофра ПНД", "in", grp=G2)
col("floor", "Потолок - 1/Пол - 2", "in", grp=G2)
col("cable_in", "Марка кабеля", "in", grp=G2, note="Обозначение из «Справочная» L (мс, мфр…). Пусто — по коду.")
col("cores_in", "Кол-во жил", "in", grp=G2)
col("s_in", "Сечение (ручное)", "in", grp=G2, note="Пусто — по коду или подбор по автомату.")
col("qf_in", "автомат (ручное)", "in", grp=G2, note="Номинал автомата линии вручную, А.")
col("klay_in", "К прокл. линии (пусто = общий)", "in", grp=G2)
G3 = "КАТЕГОРИИ (ввод)"
col("catpz_in", "Вид нагрузки (пусто = по коду)", "in", grp=G3, note="Строка таблицы нагрузок для пояснительной записки: рабочее освещение, розеточная сеть, вентиляция. Пусто — по коду.")
col("cat_in", "Категория надёжности (ввод: I особая / I / II / III; пусто = III)", "in", grp=G3)
col("fire_in", "Пожарная (ввод: да / нет; пусто = по коду)", "in", grp=G3)
G4 = "СХЕМА (ввод)"
col("bus", "Шина", "in", grp=G4, note="Пусто — основная шина. Пишется в строке с автоматом.")
col("qf_txt", "Автомат «авто» или текст", "in", grp=G4, note="«авто» или текст (1P C16A) в первой строке группы.")
col("rcd_in", "Дифзащита, мА", "in", grp=G4, note="Пусто — по коду (лист «Коды», таблица УГО). Число — ток утечки: 10, 30, 100, 300. «нет» — без дифзащиты. «да» = 30 мА.")
col("dev1", "Аппарат 1", "in", grp=G4)
col("dev2", "Аппарат 2", "in", grp=G4)
col("dev3", "Аппарат 3", "in", grp=G4)
col("module", "Модуль", "in", grp=G4, note="Модель — в первой строке модуля; следующие строки идут на его каналы. «-» — в первой строке ПОСЛЕ модуля (модуль закончился).")
col("ch_in", "Канал", "in", grp=G4, note="Пусто — авто; № вручную; = параллельно; + в тот же кабель; р — резерв; б/к — без клеммы.")
col("xt", "Клемма", "in", grp=G4, note="Пусто — номер XT по порядку; «нет» — клемма без номера.")
col("mode", "Режим", "in", grp=G4, note="Пусто — как обычно. «вручную» — строка только в таблице, в схеме не рисуется. «рисовать» — нарисовать строку вводного узла (по умолчанию вводной узел не рисуется).")
G5 = "DWG (заполняет DLN)"
col("dwg_pvc", "DWG: гофра ПВХ по плану, м", "in", grp=G5)
col("dwg_pnd", "DWG: гофра ПНД по плану, м", "in", grp=G5)
col("dwg_h", "DWG: хэндлы полилиний", "in", grp=G5)
col("dwg_note", "DWG: примечание", "in", grp=G5)
col("dwg_label", "Доп информация из выноски", "in", grp=G5)

# --- расчёт линии
G6 = "РАСЧЁТ ЛИНИИ"
col("_n", "_n", "h", "=ROW()-ROW(Линии[#Headers])")
col("code", "Код", "h", '=IF({line}3="","",IFERROR(LEFT({line}3,SEARCH(".",{line}3)-1),{line}3))')
col("rest1", "_часть1", "h", '=IFERROR(RIGHT({line}3,LEN({line}3)-SEARCH(".",{line}3)),"")')
col("floorcode", "Этаж/Шлейф", "h", '=IFERROR(LEFT({rest1}3,SEARCH(".",{rest1}3)-1),{rest1}3)')
col("rest2", "_часть2", "h", '=IFERROR(RIGHT({rest1}3,LEN({rest1}3)-SEARCH(".",{rest1}3)),"")')
col("room", "Помещение", "h", '=IFERROR(LEFT({rest2}3,SEARCH(".",{rest2}3)-1),{rest2}3)')
col("name_code", "Наименование по коду", "h",
    '=IF({line}3="","",IF(ISERROR(SEARCH(".",{line}3)),IFERROR(VLOOKUP({line}3,' + POT + ',2,FALSE)&"",""),IFERROR(IF(VLOOKUP({code}3,' + POT + ',4,FALSE)=1,VLOOKUP({code}3,' + POT + ',2,FALSE)&" "&'
    'VLOOKUP({floorcode}3,Этаж_обозначение,2,FALSE)&" Пом. №"&{room}3,VLOOKUP({code}3,' + POT + ',2,FALSE)&" Шлейф №"&{floorcode}3),'
    'IFERROR(VLOOKUP({line}3,' + POT + ',2,FALSE),""))))')
col("name", "Наименование (итог)", "h", '=IF({name_m}3<>"",{name_m}3,{name_code}3)')
col("feeder", "_фидер", "h", '=IF({feeds}3<>"",{feeds}3,IF(AND({line}3<>"",{line}3<>{board}3,ISNUMBER(XMATCH({line}3,Щиты[Щит]))),{line}3,""))')
col("bno", "_№ щита", "h", '=IFERROR(XMATCH({board}3,Щиты[Щит]),0)')
col("nin", "_мест узла", "h", '=IFERROR(N(XLOOKUP({board}3,Щиты[Щит],Щиты[Мест под вводной узел])),15)')
col("pos_eff", "_№ места (итог)", "h",
    '=IF({pos}3="-","",IF(ISNUMBER({pos}3),{pos}3,LET(prev,IF({NB},0,N({P:pos_eff})),nx,prev+1,IF(AND({mode}3<>"вручную",nx<={nin}3),{nin}3+1,nx))))')
col("zone", "_зона", "h", '=IF({pos_eff}3="","",IF({pos_eff}3<={nin}3,"вводной узел","линии"))')
col("p", "Р, кВт", "c", '=IF({line}3="","",IF({feeder}3<>"",IFERROR(XLOOKUP({feeder}3,Щиты[Щит],Щиты[Руст, кВт]),0),N({p_in}3)))',
    note="Мощность линии; для фидера — Руст нижестоящего щита.")
col("u", "U, В", "h", '=IF({u_in}3<>"",{u_in}3,IF({line}3="","",IFERROR(IF(VLOOKUP({code}3,' + POT + ',8,FALSE)=0,"",VLOOKUP({code}3,' + POT + ',8,FALSE)),"")))')
col("ph", "Фаз", "c", '=IF({ph_in}3<>"",{ph_in}3,IF({line}3="","",IFERROR(IF(VLOOKUP({code}3,' + POT + ',9,FALSE)=0,"",VLOOKUP({code}3,' + POT + ',9,FALSE)),"")))')
col("cable", "Марка (итог)", "h", '=IF({cable_in}3<>"",{cable_in}3,IF({line}3="","",IFERROR(IF(VLOOKUP({code}3,' + POT + ',6,FALSE)=0,"",VLOOKUP({code}3,' + POT + ',6,FALSE)),"")))')
col("cores", "Жил", "h", '=IF({cores_in}3<>"",{cores_in}3,IF({line}3="","",IFERROR(IF(VLOOKUP({code}3,' + POT + ',7,FALSE)=0,"",VLOOKUP({code}3,' + POT + ',7,FALSE)),"")))')
col("s_set", "Сечение задано", "h", '=IF({s_in}3<>"",{s_in}3,IF({line}3="","",IFERROR(IF(VLOOKUP({code}3,' + POT + ',10,FALSE)=0,"",VLOOKUP({code}3,' + POT + ',10,FALSE)),"")))')
col("tstart", "_тип пуска", "h", '=IFERROR(VLOOKUP({code}3,' + POT + ',5,FALSE),0)')
col("cos", "cos φ", "c",
    '=IF({line}3="","",IF({feeder}3<>"",IFERROR(XLOOKUP({feeder}3,Щиты[Щит],Щиты[cos φ]),0.95),IF(ISNUMBER({cos_in}3),{cos_in}3,'
    'IFERROR(IF({tstart}3=2,IFS({ph}3=3,VLOOKUP({p}3,Справочная_двигатели,8,FALSE),{ph}3=1,VLOOKUP({p}3,Справочная_двигатели,10,TRUE)),'
    'IF(AND(Методика="нормы 2026",{grp}3="Сил"),0.8,0.95)),0.95))))',
    note="Ручной; для двигателей — по мощности; иначе 0,95. В режиме «нормы 2026» силовые — 0,8 (СП 256 п. 7.1.12).")
col("i", "Ток, А", "c",
    '=IF({line}3="","",IF({feeder}3<>"",IFERROR(XLOOKUP({feeder}3,Щиты[Щит],Щиты[Iр, А]),""),'
    'IFERROR(IFS({ph}3=1,{p}3/(' + U_F + '*{cos}3),{ph}3=3,{p}3/(SQRT(3)*Uн_лин*{cos}3)),"")))',
    note="1ф: P/(Uф·cos φ); 3ф: P/(√3·Uн·cos φ). Фидер — Iр нижестоящего щита.")
col("qf_line", "Автомат", "c", '=IF({line}3="","",IF(N({qf_in}3)=0,IFERROR(VLOOKUP({i}3,Автомат_номиналы,2,TRUE),""),{qf_in}3))',
    note="Наименьший номинал с Iр ≤ 0,9·In (ГОСТ Р 50571.4.43 п. 433.1 + запас 10 %).")
col("kgr", "_k жил", "h", '=IF(N({cores}3)=1,1,IF({ph}3=1,1,IF(N({cores}3)>=4,0.92,1)))')
col("klay", "_К прокл", "h", '=IF(ISNUMBER({klay_in}3),{klay_in}3,К_прокл)')
col("tcol", "_столбец Iдоп", "h", '=IF(N({cores}3)=1,1,IF({ph}3=1,2,3))')
col("s", "Сечение", "c",
    '=IF({line}3="","",IF(N({s_set}3)<>0,{s_set}3,IFERROR(LET(need,{qf_line}3/({kgr}3*{klay}3),t,CHOOSE({tcol}3,Iдоп_1ж,Iдоп_2ж,Iдоп_3ж),'
    'k,XMATCH(need,t,1),IF(ISNA(k),"Не исп",INDEX(Iдоп_S,k))),"Не исп")))',
    note="Наименьшее сечение с Iдоп·k ≥ In. Таблица — ПУЭ 1.3.7 (медь, воздух), 4–5 жил ×0,92 (ПУЭ 1.3.10).")
col("r", "_ρ", "h", '=IFERROR(VLOOKUP({s}3,Кабель_сопротивление,3,FALSE),"")')
col("x", "_x", "h", '=IFERROR(VLOOKUP({s}3,Кабель_сопротивление,2,FALSE),"")')
col("lplan", "_Lплан", "h", '=IF({lplan_in}3<>"",{lplan_in}3,IF(COUNT({dwg_pvc}3,{dwg_pnd}3)=0,"",ROUNDUP(SUM({dwg_pvc}3,{dwg_pnd}3),0)))')
col("l", "Длина, м", "c",
    '=IF({line}3="","",IFERROR(IF(VLOOKUP({code}3,' + POT + ',4,FALSE)=1,VLOOKUP({code}3,' + POT + ',3,FALSE)*({n}3*2-1)+{lplan}3+'
    'VLOOKUP({floorcode}3,Этаж_обозначение,3,FALSE)+VLOOKUP({floorcode}3,Этаж_обозначение,4,FALSE),VLOOKUP({code}3,' + POT + ',3,FALSE)*({n}3*2-1)+{lplan}3),{lplan}3))',
    note="Опуск·(2n−1) + длина по плану + подъём на этаж.")
col("pvc", "_гофра ПВХ", "h", '=IF({pvc_in}3<>"",{pvc_in}3,IF(OR({lplan}3="",{floor}3=""),"",IF({floor}3=1,MAX(0,{l}3-ROUNDUP(N({dwg_pnd}3),0)-4),ROUNDUP(N({dwg_pvc}3),0))))')
col("pnd", "_гофра ПНД", "h", '=IF({pnd_in}3<>"",{pnd_in}3,IF(OR({lplan}3="",{floor}3=""),"",IF({floor}3=2,MAX(0,{l}3-ROUNDUP(N({dwg_pvc}3),0)-4),ROUNDUP(N({dwg_pnd}3),0))))')
col("b", "_b", "h", '=IF({ph}3=1,2,IF({ph}3=3,1,""))')
col("sin", "_sin", "h", '=IFERROR(ROUND(SIN(ACOS({cos}3)),2),"")')
col("du_v", "_ΔU, В", "h", '=IFERROR({b}3*({r}3*{l}3/{s}3*{cos}3+{x}3*{l}3*{sin}3/1000)*{i}3,"")')
col("du", "ΔU, %", "c", '=IFERROR(ROUNDUP({du_v}3/(' + U_F + '*1000)*100,2),"")',
    note="ΔU = b·(ρ·L/S·cos φ + x·L·sin φ)·I (ГОСТ Р 50571.5.52 прил. G), % от Uф.")
col("du_vru", "ΔU от ВРУ, %", "c", '=IF(ISNUMBER({du}3),{du}3+IFERROR(XLOOKUP({board}3,Щиты[Щит],Щиты[ΔU до шин, %]),0),"")')
col("du_max", "Допустимо ΔU, %", "c", '=IF({line}3="","",IF(IFERROR(VLOOKUP({code}3,УГО_спр,2,FALSE),"")="Светильник",3,4)+IF(Методика="нормы 2026",MIN(0.5,MAX(0,N({l}3)-100)*0.005),0))',
    note="СП 256 п. 8.23: от ВРУ до светильников ≤ 3 %, до прочих ≤ 4 %.")
col("cab", "_кабель", "h", '=IFERROR(VLOOKUP({cable}3,Кабель_обозначение,2,FALSE),"")')
col("cab_sec", "_кабель с сечением", "h", '=IFERROR(IF(VLOOKUP({cable}3,Кабель_обозначение,3,FALSE)=1,{cab}3&" ("&{cores}3&"х"&{s}3&")",{cab}3),"")')
col("cable_full", "Кабель", "c", '=IF(OR({line}3="",{cab}3="",AND({p_in}3<>"",NOT(ISNUMBER({p_in}3)))),"",IF({s}3=0,{cab}3,{cab_sec}3&CHAR(10)&"L="&{l}3&"-"&{du}3&"%"))')
col("zcab", "_Zкабеля", "h", '=IFERROR(2*ρкз_Cu*{l}3/{s}3,"")')
col("zmcab", "_Zmax кабеля", "h", '=IFERROR(0.0175*N({l}3)/{s}3,"")')
col("zloop", "Zпетли в конце, Ом", "c",
    '=IFERROR(IF(OR({line}3="",NOT(ISNUMBER({s}3)),N({l}3)=0),"",XLOOKUP({board}3,Щиты[Щит],Щиты[Zпетли на шинах, Ом])+{zcab}3),"")')
col("ikz_min", "Iкз min в конце линии, А", "c", '=IF(ISNUMBER({zloop}3),c_min*(Uн_лин*1000/SQRT(3))/{zloop}3,"")',
    note="Iкз = c_min·Uф / (Zшин + 2·ρкз·L/S), ρкз = 1,5·ρ20 (ГОСТ 28249 п. 2.4.2).")
col("ikz_max", "Iкз max в конце, кА", "c",
    '=IFERROR(IF(OR({line}3="",NOT(ISNUMBER({s}3))),"",1.05*(Uн_лин*1000/SQRT(3))/(XLOOKUP({board}3,Щиты[Щит],Щиты[Zmax на шинах, Ом])+{zmcab}3)/1000),"")')
col("grp", "Группа Кс (итог)", "c",
    '=IF({line}3="","",IF({feeder}3<>"","Щит",IF({grp_in}3<>"",{grp_in}3,IF(ISNUMBER({kc_in}3),"Сил",'
    'IF(IFERROR(XLOOKUP({code}3,Коды_код,Коды_группа)&"","")="","Быт",XLOOKUP({code}3,Коды_код,Коды_группа)&"")))))')
col("kc", "Кс", "c",
    '=IF({grp}3="Щит","щит",IF({grp}3="","",IF({grp}3="Быт",IFERROR(XLOOKUP({board}3,Щиты[Щит],Щиты[Кс быт]),1),'
    'IF({grp}3="Сил",IF(ISNUMBER({kc_in}3),{kc_in}3,IF(Методика="нормы 2026",IFERROR(XLOOKUP({board}3,Щиты[Щит],Щиты[Кс силовых (табл. 7.5)]),1),1)),'
    'IF(OR({grp}3="Пост",{grp}3="Отоп"),1,IF({grp}3="ВНГ",0.7,0))))))')
col("sec", "Секция", "h", '=IF({grp}3="","",IFERROR(MIN(3,MAX(1,INT(--{sec_in}3))),1))')
col("c_all", "_Руст (без Нет)", "h", '=IF(OR({grp}3="",{grp}3="Нет",{grp}3="Щит"),0,N({p}3))')
col("c_byt", "_Руст быт", "h", '=IF({grp}3="Быт",N({p}3),0)')
col("c_sil", "_Рр сил", "h", '=IF({grp}3="Сил",N({p}3)*N({kc}3),0)')
col("c_post", "_Руст пост", "h", '=IF({grp}3="Пост",N({p}3),0)')
col("c_otop", "_Р отопления", "h", '=IF({grp}3="Отоп",N({p}3),0)')
col("c_vng", "_Рр ВНГ", "h", '=IF({grp}3="ВНГ",N({p}3)*0.7,0)')
col("c_pw", "_P для cos", "h", '=IF(OR({line}3="",{grp}3="Щит"),0,N({p}3))')
col("c_pcos", "_P·cos", "h", '=IF(OR({line}3="",{grp}3="Щит"),0,N({p}3)*N({cos}3))')
col("phs", "_фаза №", "h", '=LET(prev,IF(OR({NB}),0,N({P:phs})),IF(OR({line}3="",{ph}3=3),prev,IF(prev>=3,1,prev+1)))')
col("phase_auto", "_фаза авто", "h", '=IF({line}3="","",IF({ph}3=3,"L1,L2,L3","L"&MAX(1,{phs}3)))')
col("ip1", "_Iпуск двиг.", "h", '=IFERROR(IF({tstart}3=2,IFS({ph}3=3,VLOOKUP({p}3,Справочная_двигатели,3,FALSE),{ph}3=1,VLOOKUP({p}3,Справочная_двигатели,5,TRUE)),0),0)')
col("ip2", "_Iпуск LED", "h", '=IFERROR(IF({tstart}3=3,VLOOKUP({p}3,Справочная_ЛЭД,2,TRUE),0),0)')
col("ipusk", "Iпуск, А", "h", '=IF({line}3="","",IF({ip1}3=0,{ip2}3,{ip1}3))')
col("phase", "Фаза", "c", '=IF({line}3="","",IF({phase_in}3="упр","",IF({phase_in}3="",{phase_auto}3,IF({phase_in}3&""="1,2,3","L1,L2,L3","L"&{phase_in}3))))')

# --- категории надёжности и шины (порт v3.12 BP:CA)
G8 = "КАТЕГОРИИ (расчёт)"
RANK = '{"III","II","I","I особая"}'
col("cat", "Категория (итог)", "c", '=IF({line}3="","",IFERROR(INDEX({"I особая","I","II","III"},MATCH(TRIM({cat_in}3&""),{"I особая","I","II","III"},0)),"III"))', grp=G8)
col("fire", "Пожарная (итог)", "c",
    '=IF({line}3="","",IF(TRIM({fire_in}3&"")<>"",IF(LOWER(TRIM({fire_in}3&""))="да","да","нет"),IF(IFERROR(XLOOKUP({code}3,Коды_код,Коды_пожарная),"")="да","да","нет")))')
col("bus_grp", "_шина группы", "h", '=IF({bus}3<>"",{bus}3,IF({qf_txt}3<>"","",IF({ag}3="","",{P:bus_grp})))')
col("bus_name", "Шина (итог)", "c", '=IF({line}3="","",IF({bus_grp}3="","основная",{bus_grp}3))')
col("sec_n", "_секция №", "h", '=IFERROR(MIN(3,MAX(1,INT(--{sec_in}3))),1)')
col("bus_type", "Тип шины", "c",
    '=IF({line}3="","",LET(t,IFERROR(XLOOKUP(1,(Шины[Шина]={bus_name}3)*(((Шины[Щит]={board}3)+(Шины[Щит]=""))>0),Шины[Тип]),"")&"",'
    'IF(t<>"",t,IF(ISNUMBER(SEARCH("ИБП",{bus_name}3)),"ИБП",IF(ISNUMBER(SEARCH("ДГУ",{bus_name}3)),"ДГУ","Ввод "&{sec_n}3)))))')
col("src", "Источник для режимов", "c", '=IF({line}3="","",IF(OR({bus_type}3="ИБП",{bus_type}3="ДГУ"),{bus_type}3,IF(LEFT({bus_type}3,5)="Ввод ",{bus_type}3,"Ввод "&{sec_n}3)))')
col("cat_prov", "Категория обеспечена", "c",
    '=IF({line}3="","",IF(OR({bus_type}3="ИБП",{bus_type}3="ДГУ"),"I особая",IF({bus_type}3="Одиночная","III",XLOOKUP({board}3,Щиты[Щит],Щиты[Категория шины «Ввод»],"III"))))')
col("cat_rank", "_ранг категории", "h", '=IF({line}3="",0,IFERROR(MATCH({cat}3,' + RANK + ',0),1))')
col("cat_rank_f", "_ранг для фидера", "h", '=IF(OR({bus_type}3="ИБП",{bus_type}3="ДГУ"),MIN({cat_rank}3,3),{cat_rank}3)')
col("cat_chk", "Категория: проверка", "c",
    '=IF({line}3="","",IF({cat_rank}3>IFERROR(MATCH({cat_prov}3,' + RANK + ',0),1),"НЕ ОБЕСПЕЧЕНА","ОК"))',
    note="Требуемая категория выше, чем даёт схема (ПУЭ 1.2.17–1.2.21, СП 256 п. 6.13, 8.10).")
col("bpath", "_путь щита", "h", '=IF({board}3="","",XLOOKUP({board}3,Щиты[Щит],Щиты[Путь],""))')
col("kind", "Вид нагрузки (итог)", "h",
    '=IF({line}3="","",IF({catpz_in}3<>"",{catpz_in}3,IF(IFERROR(XLOOKUP({code}3,Коды_код,Коды_вид)&"","")="","Прочее",XLOOKUP({code}3,Коды_код,Коды_вид)&"")))')

# --- схема (порт листа «Однолинейка» v3.12)
G7 = "СХЕМА (расчёт)"
col("sheetno", "Лист схемы", "h", '=IF({pos_eff}3="","",INT(({pos_eff}3-1)/Столбцов_на_лист)+1)')
col("q", "Кабель (строкой)", "h", '=IF({line}3="","",SUBSTITUTE({cable_full}3,CHAR(10)," "))')
col("t", "_Р схемы", "h", '=IF({line}3="","",IFERROR(--{p}3,""))')
col("uu", "_Ток схемы", "h", '=IF({line}3="","",IFERROR(ROUNDUP({i}3,2),""))')
col("ak", "_фаз", "h", '=IF({line}3="","",IFERROR(--{ph}3,1))')
col("am", "_дифзащита по коду, мА", "h",
    '=IF({code}3="",0,LET(v,IFERROR(VLOOKUP({line}3,УГО_спр,3,FALSE),IFERROR(VLOOKUP({code}3,УГО_спр,3,FALSE),IFERROR(VLOOKUP(LEFT({code}3,2),УГО_спр,3,FALSE),""))),'
    'IF(v="да",30,IF(ISNUMBER(v),v,0))))')
col("an", "_в кабель выше", "h", '=IF(OR({ch_in}3="+",AND({line}3<>"",{line}3={P:line})),1,0)')
col("af", "_гр.", "h", '=IF({pos_eff}3="","",IF({qf_txt}3<>"",COUNTIFS({board}$3:{board}3,{board}3,{qf_txt}$3:{qf_txt}3,"?*"),IF(OR({NB},{mode}3="вручную"),"",{P:af_c})))')
col("af_c", "_гр. (перенос)", "h", '=IF({pos_eff}3="",IF({NB},"",{P:af_c}),{af}3)')
col("bb", "_мин. по типу", "h",
    '=IF({line}3="",0,MAX(IFERROR(--{qf_in}3,0),IFERROR(--VLOOKUP({line}3,УГО_спр,4,FALSE),IFERROR(--VLOOKUP({code}3,УГО_спр,4,FALSE),'
    'IFERROR(--VLOOKUP(LEFT({code}3,2),УГО_спр,4,FALSE),0)))))')
col("ab", "_макс. автомат по кабелю", "h",
    '=IF({line}3="","",IFERROR(LOOKUP(INDEX(CHOOSE({tcol}3,Iдоп_1ж,Iдоп_2ж,Iдоп_3ж),MATCH({s}3,Iдоп_S,0))*{kgr}3*{klay}3,'
    '{6,10,16,20,25,32,40,50,63,80,100,125,160,200,250,315,400,500,630}),""))')
col("z", "_ток группы", "h", '=IF({qf_txt}3="","",SUMIFS({T:uu},{T:af},{af}3,{T:board},{board}3,{T:an},0))')
col("aa", "_номинал по току", "h", '=IF({qf_txt}3="","",IFERROR(VLOOKUP({z}3,Автомат_номиналы,2,TRUE),""))')
col("ao", "_по току", "h", '=IF({qf_txt}3="","",MAX(N({aa}3),MAXIFS({T:bb},{T:af},{af}3,{T:board},{board}3)))')
col("ap", "_мин. по кабелям", "h", '=IF({qf_txt}3="","",MINIFS({T:ab},{T:af},{af}3,{T:board},{board}3))')
col("ar", "_фаз гр.", "h", '=IF({qf_txt}3="","",MAX(1,MAXIFS({T:ak},{T:af},{af}3,{T:board},{board}3)))')
col("rcd_ma", "Дифзащита группы, мА", "h",
    '=IF({qf_txt}3="","",IF(AND({qf_txt}3<>"авто",{qf_txt}3<>"+"),IF(ISNUMBER(SEARCH("мА",{qf_txt}3)),'
    'IFERROR(--TRIM(RIGHT(SUBSTITUTE(LEFT({qf_txt}3,SEARCH("мА",{qf_txt}3)-1)," ",REPT(" ",9)),6)),30),0),'
    'IF({rcd_in}3="нет",0,IF(ISNUMBER({rcd_in}3),{rcd_in}3,IF({rcd_in}3="да",30,MINIFS({T:am},{T:af},{af}3,{T:board},{board}3,{T:am},">0"))))))',
    note="Самая строгая уставка среди линий группы (меньший ток утечки).")
col("as_", "_дифзащита гр. (0/1)", "h", '=IF({qf_txt}3="","",IF(N({rcd_ma}3)>0,1,0))')
col("aj", "_QF нач.", "h", '=IF({qf_txt}3="","",IF({as_}3=1,"QFD","QF")&{af}3)')
col("ag", "_QF гр.", "h", '=IF({aj}3<>"",{aj}3,IF(OR({NB},{mode}3="вручную"),"",{P:ag}))')
col("az", "_аппараты гр.", "h", '=IF({aj}3<>"",TEXTJOIN(" → ",TRUE,{dev1}3,{dev2}3,{dev3}3),IF({ag}3="","",{P:az}))')
col("ba", "_БП/тр-р", "h",
    '=IF({az}3="",0,IF(OR(LEFT({az}3,2)="T ",LEFT({az}3,2)="Т ",ISNUMBER(SEARCH("→ T ",{az}3)),ISNUMBER(SEARCH("→ Т ",{az}3)),'
    'ISNUMBER(SEARCH("БП",{az}3))),1,0))')
col("aq", "_номинал", "h", '=IF({qf_txt}3="","",IF({ba}3=1,MAX({ao}3,Мин_автомат_БП),{ao}3))')
col("at", "_полюса", "h", '=IF({qf_txt}3="","",IF({ar}3>=3,IF({as_}3=1,"4P","3P"),IF({as_}3=1,"2P","1P")))')
col("h", "Автомат (схема)", "c", '=IF({qf_txt}3="","",IF(AND({qf_txt}3<>"авто",{qf_txt}3<>"+"),{qf_txt}3,{at}3&" С"&{aq}3&"А"&IF({as_}3=1," "&{rcd_ma}3&"мА","")))',
    grp=G7, note="Групповой автомат: ток группы, мин. автомат по типу, дифзащита (мА), полюса.")
col("au", "_номинал итог", "h", '=IF({h}3="","",LET(t,SUBSTITUTE(SUBSTITUTE(UPPER({h}3),"С","C"),"А","A"),c,FIND("C",t),IFERROR(--MID(t,c+1,FIND("A",t&"A",c)-c-1),"")))')
col("av", "_префикс", "h", '=IF(OR({module}3="",{module}3="-"),"",IFERROR(VLOOKUP({module}3,Модули_спр,2,FALSE),"R"))')
col("aw", "_модуль нач.", "h", '=IF({av}3="","",{av}3&"."&COUNTIFS({board}$3:{board}3,{board}3,{av}$3:{av}3,{av}3))')
col("ah", "_модуль гр.", "h", '=IF({module}3="-","",IF({aw}3<>"",{aw}3,IF({NB},"",{P:ah})))')
col("ai", "_модель гр.", "h", '=IF({ah}3="","",IF(AND({module}3<>"",{module}3<>"-"),{module}3,{P:ai}))')
col("o", "Канал (итог)", "h",
    '=IF({ch_in}3="=",{P:ax},IF(ISNUMBER({ch_in}3),{ch_in}3,IF({ah}3="","",IF(OR({line}3<>"",{ch_in}3="+",{ch_in}3="р"),'
    'IF(OR({aw}3<>"",{P:ax}=""),1,{P:ax}+1),""))))')
col("ax", "_посл. канал", "h", '=IF({o}3<>"",{o}3,IF(OR({NB},{module}3="-",{aw}3<>""),"",{P:ax}))')
col("ay", "_клемма", "h", '=IF({ch_in}3="б/к",0,IF(OR(ISNUMBER(SEARCH("L=",{cable_full}3)),{ch_in}3="+",{ch_in}3="="),1,0))')
col("xt_eff", "Клемма (итог)", "h", '=IF({xt}3="нет","",IF({xt}3<>"",{xt}3,IF({ay}3=1,"XT"&COUNTIFS({board}$3:{board}3,{board}3,{ay}$3:{ay}3,1),"")))')
col("g", "QF", "c", '=IF(OR({line}3<>"",{o}3<>"",{ay}3=1),{ag}3,"")')
col("m", "_Модуль (схема)", "h", '=IF(OR({o}3<>"",AND({module}3<>"",{module}3<>"-")),{ah}3,"")')
col("w", "Проверка автомата", "c",
    '=IF({qf_txt}3="","",IF(NOT(ISNUMBER({au}3)),"",IF({z}3>{au}3,"ПЕРЕГРУЗ",IF(AND({ba}3=0,{ap}3>0,{au}3>{ap}3),"КАБЕЛЬ НЕ ЗАЩИЩЁН","ok"))))',
    note="ПЕРЕГРУЗ: ток группы > In. КАБЕЛЬ НЕ ЗАЩИЩЁН: In > Iдоп кабеля (ГОСТ Р 50571.4.43 п. 433.1).")
col("x_", "Проверка канала", "h",
    '=IF(OR({o}3="",{ah}3=""),"",IF({o}3>IFERROR(VLOOKUP({ai}3,Модули_спр,3,FALSE),99),"ЛИШНИЙ КАНАЛ",'
    'IF(AND({ch_in}3<>"=",COUNTIFS({T:ah},{ah}3,{T:o},{o}3,{T:ch_in},"<>=",{T:board},{board}3)>1),"ДУБЛЬ КАНАЛА","")))')
col("in_grp", "_In гр.", "h", '=IF({qf_txt}3<>"",{au}3,IF({ag}3="","",{P:in_grp}))')
col("h_grp", "_автомат гр.", "h", '=IF({qf_txt}3<>"",{h}3,IF({ag}3="","",{P:h_grp}))')
col("rcd_grp", "_дифзащита гр. (строки)", "h", '=IF({qf_txt}3<>"",{as_}3,IF({ag}3="","",{P:rcd_grp}))')
col("rcd_ma_grp", "_мА гр. (строки)", "h", '=IF({qf_txt}3<>"",{rcd_ma}3,IF({ag}3="","",{P:rcd_ma_grp}))')
col("ikz_ratio", "Iкз / In", "c", '=IFERROR(IF(ISNUMBER({ikz_min}3),{ikz_min}3/IFERROR(--{in_grp}3,N({qf_line}3)),""),"")')
col("mult", "_кратность", "h", '=IF(OR(ISNUMBER(SEARCH(" B"," "&{h_grp}3)),ISNUMBER(SEARCH(" В"," "&{h_grp}3))),5,IF(ISNUMBER(SEARCH(" D"," "&{h_grp}3)),20,10))')
col("trip", "Отключение ≤ 0,4 с", "c",
    '=IF(NOT(ISNUMBER({ikz_ratio}3)),"",IF({ikz_ratio}3>={mult}3,"ok",IF(N({rcd_grp}3)=1,"ok (дифзащита "&{rcd_ma_grp}3&" мА)","мало: Iкз < "&{mult}3&"·In")))',
    note="ПУЭ 1.7.79: ≤ 0,4 с. Iкз/In ≥ 5 (B), 10 (C), 20 (D) — ГОСТ IEC 60898-1.")
col("ae", "Статус", "c",
    '=IF(AND({line}3="",{g}3=""),"",TEXTJOIN("; ",TRUE,IF(AND({line}3<>"",{phase}3=""),"нет фазы",""),'
    'IF(AND({line}3<>"",N({t}3)=0,{grp}3<>"Нет"),"нет мощности",""),IF(AND({xt_eff}3<>"",COUNTIFS({T:board},{board}3,{T:xt_eff},{xt_eff}3)>1),"дубль клеммы "&{xt_eff}3,""),'
    'IF(AND({w}3<>"",{w}3<>"ok"),{w}3,""),IF({x_}3<>"",{x_}3,""),IF(AND(ISNUMBER(SEARCH("(5х",{q}3)),N({ak}3)=1),"5 жил на 1 фазе",""),'
    'IF(AND({code}3="R",{qf_txt}3<>"",OR(N({rcd_ma}3)=0,N({rcd_ma}3)>30)),"розетка без дифзащиты до 30 мА",""),'
    'IFERROR(IF(AND(ISNUMBER({du_vru}3),{du_vru}3>{du_max}3),"потери "&FIXED({du_vru}3,1)&" % > "&{du_max}3&" %",""),""),'
    'IFERROR(IF(LEFT({trip}3,4)="мало","Iкз мало — отключение > 0,4 с",""),""),IF(AND({line}3<>"",{pos_eff}3=""),"не на схеме","")))',
    note="Красный — ошибка (дубль, перегруз, кабель не защищён, канал, без 30 мА), жёлтый — проверить.")
col("icu", "Icu автомата, кА", "h",
    '=IF({qf_txt}3="","",LET(t,IFERROR(XLOOKUP({h}3,Каталог!$U$4:$U$500,Каталог!$G$4:$G$500),""),k,SEARCH("кА",t),'
    'IFERROR(--TRIM(RIGHT(SUBSTITUTE(MID(t,MAX(1,k-4),k-MAX(1,k-4))," ",REPT(" ",9)),5)),"")))')
col("icu_chk", "Проверка Icu", "c",
    '=IF(OR({icu}3="",{qf_txt}3=""),"",IF(XLOOKUP({board}3,Щиты[Щит],Щиты[Iкз max на шинах, кА],"")="","",'
    'IF({icu}3<XLOOKUP({board}3,Щиты[Щит],Щиты[Iкз max на шинах, кА]),"Icu "&{icu}3&" кА < Iкз max","ok")))',
    note="Отключающая способность автомата не меньше Iкз max на шинах (ГОСТ Р 50571.4.43, СП 256).")
col("issues", "Замечания", "c",
    '=TEXTJOIN("; ",TRUE,{ae}3,IF({cat_chk}3="НЕ ОБЕСПЕЧЕНА","категория "&{cat}3&" не обеспечена (есть "&{cat_prov}3&")",""),'
    'IF(LEFT({icu_chk}3,3)="Icu",{icu_chk}3,""),IF(AND({line}3<>"",{s}3="Не исп"),"сечение не подобрано",""))',
    note="Все замечания по строке. Сводка по объекту — лист «Проверка».")
col("chain", "_цепь для бирок", "h", '=IF(OR({mode}3="вручную",{pos_eff}3="",AND({line}3="",{xt_eff}3="")),"",TEXTJOIN("|",TRUE,{ag}3,SUBSTITUTE({az}3," → ","|"),IF({ah}3<>"",{ah}3&IF({o}3<>""," к."&{o}3,""),""),{xt_eff}3))')
col("chain2", "_цепь с кабелем", "h", '=IF({chain}3="","",{chain}3&IF({line}3<>"","|кабель "&{cab_sec}3&" → "&{name}3,""))')
col("tg1", "_бирка 1", "h", '=IF({chain2}3="","",LET(nd,TEXTSPLIT({chain2}3,"|"),IF(COLUMNS(nd)>1,{board}3&"¦"&INDEX(nd,1)&"¦"&INDEX(nd,2),"")))')
col("ti1", "_бирка 1 линия", "h", '=IF({tg1}3="","",{line}3&"¦"&{name}3)')
col("tg2", "_бирка 2", "h", '=IF({chain2}3="","",LET(nd,TEXTSPLIT({chain2}3,"|"),IF(COLUMNS(nd)>2,{board}3&"¦"&INDEX(nd,2)&"¦"&INDEX(nd,3),"")))')
col("ti2", "_бирка 2 линия", "h", '=IF({tg2}3="","",{line}3&"¦"&{name}3)')
col("tg3", "_бирка 3", "h", '=IF({chain2}3="","",LET(nd,TEXTSPLIT({chain2}3,"|"),IF(COLUMNS(nd)>3,{board}3&"¦"&INDEX(nd,3)&"¦"&INDEX(nd,4),"")))')
col("ti3", "_бирка 3 линия", "h", '=IF({tg3}3="","",{line}3&"¦"&{name}3)')
col("tg4", "_бирка 4", "h", '=IF({chain2}3="","",LET(nd,TEXTSPLIT({chain2}3,"|"),IF(COLUMNS(nd)>4,{board}3&"¦"&INDEX(nd,4)&"¦"&INDEX(nd,5),"")))')
col("ti4", "_бирка 4 линия", "h", '=IF({tg4}3="","",{line}3&"¦"&{name}3)')
col("tg5", "_бирка 5", "h", '=IF({chain2}3="","",LET(nd,TEXTSPLIT({chain2}3,"|"),IF(COLUMNS(nd)>5,{board}3&"¦"&INDEX(nd,5)&"¦"&INDEX(nd,6),"")))')
col("ti5", "_бирка 5 линия", "h", '=IF({tg5}3="","",{line}3&"¦"&{name}3)')
col("tg6", "_бирка 6", "h", '=IF({chain2}3="","",LET(nd,TEXTSPLIT({chain2}3,"|"),IF(COLUMNS(nd)>6,{board}3&"¦"&INDEX(nd,6)&"¦"&INDEX(nd,7),"")))')
col("ti6", "_бирка 6 линия", "h", '=IF({tg6}3="","",{line}3&"¦"&{name}3)')
col("ac", "_модель (схема)", "h", '=IF({m}3="","",{ai}3)')
col("ad", "_цепь", "h",
    '=IF(AND({line}3="",{o}3="",{xt_eff}3=""),"",TEXTJOIN(" → ",TRUE,{ag}3,{az}3,IF({ah}3<>"",{ah}3&IF({o}3<>""," к."&{o}3,""),'
    'IF({o}3<>"","к."&{o}3,"")),{xt_eff}3,IF({an}3=1,"(в кабель выше)","")))')
col("bd", "_draw", "h", '=IF(OR({pos_eff}3="",{mode}3="вручную",AND({zone}3="вводной узел",{mode}3<>"рисовать")),0,1)')
col("be", "_bus2", "h", '=IF({bus}3="",0,1)')
col("bf", "_rcd", "h", '=IF({aj}3="",0,IF(LEFT({aj}3,3)="QFD",1,0))')
col("bg", "_chtype", "h", '=IF({ch_in}3="=","=",IF({ch_in}3="+","+",IF({ch_in}3="р","R","")))')
DEVT = ('=IF({d}3="","",IF(OR(LEFT({d}3,1)="T",LEFT({d}3,1)="Т"),"T",IF(LEFT(UPPER({d}3),2)="БП","PS",IF(OR(LEFT(UPPER({d}3),2)="МК",'
        'LEFT(UPPER({d}3),2)="MK"),"MK",IF(OR(LEFT(UPPER({d}3),2)="КМ",LEFT(UPPER({d}3),2)="KM"),"KM","X")))))')
col("bh", "_dev1", "h", DEVT.replace("{d}", "{dev1}"))
col("bi", "_dev2", "h", DEVT.replace("{d}", "{dev2}"))
col("bj", "_dev3", "h", DEVT.replace("{d}", "{dev3}"))
col("by", "_УГО (рус)", "h",
    '=IF({line}3="","",IFERROR(VLOOKUP({line}3,УГО_спр,2,FALSE),IFERROR(VLOOKUP({code}3,УГО_спр,2,FALSE),IFERROR(VLOOKUP(LEFT({code}3,2),УГО_спр,2,FALSE),""))))')
col("bk", "_sym", "h",
    '=IF({line}3="","",IF(OR({by}3="",{by}3=0),"BOX",IFERROR(IF(VLOOKUP({by}3,УГО_блоки,3,FALSE)&""<>"",VLOOKUP({by}3,УГО_блоки,3,FALSE),'
    'VLOOKUP({by}3,УГО_блоки,2,FALSE)),{by}3)))')
col("bl", "_hasline", "h", '=IF({line}3="",0,1)')
col("bm", "_тип автомата №", "h", '=IF({h}3="","",IF(COUNTIF({h}$3:{h}3,{h}3)=1,MAX({bm}$2:{bm}2)+1,""))')
col("bn", "_аппарат 1 №", "h", '=IF({dev1}3="","",IF(COUNTIF({dev1}$3:{dev1}3,{dev1}3)+COUNTIF({dev2}$2:{dev2}2,{dev1}3)+COUNTIF({dev3}$2:{dev3}2,{dev1}3)=1,MAX({bn}$2:{bp}2)+1,""))')
col("bo", "_аппарат 2 №", "h", '=IF({dev2}3="","",IF(COUNTIF({dev1}$3:{dev1}3,{dev2}3)+COUNTIF({dev2}$3:{dev2}3,{dev2}3)+COUNTIF({dev3}$2:{dev3}2,{dev2}3)=1,MAX({bn}$2:{bp}2,{bn}3)+1,""))')
col("bp", "_аппарат 3 №", "h", '=IF({dev3}3="","",IF(COUNTIF({dev1}$3:{dev1}3,{dev3}3)+COUNTIF({dev2}$3:{dev2}3,{dev3}3)+COUNTIF({dev3}$3:{dev3}3,{dev3}3)=1,MAX({bn}$2:{bp}2,{bn}3,{bo}3)+1,""))')
col("bq", "_кабель клеммы", "h", '=IF({ay}3=1,IFERROR(MID({q}3,FIND(" (",{q}3)+2,FIND(")",{q}3,FIND(" (",{q}3))-FIND(" (",{q}3)-2),"?"),"")')
col("br", "_вид клеммы №", "h", '=IF({bq}3="","",IF(COUNTIF({bq}$3:{bq}3,{bq}3)=1,MAX({br}$2:{br}2)+1,""))')
col("bs", "_модуль №", "h", '=IF({aw}3="","",COUNTIF({aw}$3:{aw}3,"?*"))')
col("bt", "_модель №", "h", '=IF(OR({aw}3="",{ai}3=""),"",IF(COUNTIFS({aw}$3:{aw}3,"?*",{ai}$3:{ai}3,{ai}3)=1,MAX({bt}$2:{bt}2)+1,""))')
col("bu", "_уровень", "h",
    '=IF({ae}3="",0,IF(OR(ISNUMBER(SEARCH("дубль",{ae}3)),ISNUMBER(SEARCH("ПЕРЕГРУЗ",{ae}3)),ISNUMBER(SEARCH("ЗАЩИЩ",{ae}3)),'
    'ISNUMBER(SEARCH("КАНАЛ",{ae}3)),ISNUMBER(SEARCH("без диф",{ae}3))),2,1))')
col("bv", "_№ линии", "h", '=IF({line}3="","",COUNTIF({line}$3:{line}3,"?*"))')
col("bw", "_3L клемм", "h",
    '=IF({ay}3<>1,0,LET(k,MAX(IFERROR(--VLOOKUP({line}3,УГО_спр,5,FALSE),0),IFERROR(--VLOOKUP({code}3,УГО_спр,5,FALSE),0),'
    'IFERROR(--VLOOKUP(LEFT({code}3,2),УГО_спр,5,FALSE),0)),IF(k>0,k,IF(IFERROR(--LEFT(SUBSTITUTE({bq}3,"x","х"),FIND("х",SUBSTITUTE({bq}3,"x","х"))-1),3)>3,Каталог!$S$9,1))))')
col("bx", "_доп. проходных", "h", '=IF({ay}3<>1,0,MAX(0,IFERROR(--LEFT(SUBSTITUTE({bq}3,"x","х"),FIND("х",SUBSTITUTE({bq}3,"x","х"))-1),3)-3*{bw}3))')

# «Замечания» — первым расчётным столбцом, сразу за вводом (видно без прокрутки)
_ix = {c["key"]: i for i, c in enumerate(L)}
_iss = L.pop(_ix["issues"])
_iss["grp"] = "ЗАМЕЧАНИЯ"
L.insert(max(i for i, c in enumerate(L) if c["kind"] == "in") + 1, _iss)

# Однолинейка v3.12: столбец -> ключ «Линий» (пусто = столбец пустой)
ONE_MAP = {"A": "board", "B": "pos_eff", "C": "sheetno", "D": "bus", "E": "qf_txt", "F": "rcd_in", "G": "g", "H": "h", "I": "dev1",
           "J": "dev2", "K": "dev3", "L": "module", "M": "m", "N": "ch_in", "O": "o", "P": "xt_eff", "Q": "q", "R": "line", "S": "name",
           "T": "t", "U": "uu", "V": "phase", "W": "w", "X": "x_", "Y": "mode", "Z": "z", "AA": "aa", "AB": "ab", "AC": "ac",
           "AD": "ad", "AE": "ae", "AF": "af", "AG": "ag", "AH": "ah", "AI": "ai", "AJ": "aj", "AK": "ak", "AL": "code", "AM": "am",
           "AN": "an", "AO": "ao", "AP": "ap", "AQ": "aq", "AR": "ar", "AS": "as_", "AT": "at", "AU": "au", "AV": "av", "AW": "aw",
           "AX": "ax", "AY": "ay", "AZ": "az", "BA": "ba", "BB": "bb", "BD": "bd", "BE": "be", "BF": "bf", "BG": "bg", "BH": "bh",
           "BI": "bi", "BJ": "bj", "BK": "bk", "BL": "bl", "BM": "bm", "BN": "bn", "BO": "bo", "BP": "bp", "BQ": "bq", "BR": "br",
           "BS": "bs", "BT": "bt", "BU": "bu", "BV": "bv", "BW": "bw", "BX": "bx", "BY": "by", "BZ": "bno", "CA": "zone"}

# ------------------------------------------------------------------------------------------------
# Таблица «Щиты»
# ------------------------------------------------------------------------------------------------
S = []


def scol(head, kind, f=None, note=None, grp=None):
    S.append(dict(head=head, kind=kind, f=f, note=note, grp=grp))


PATH = ('=LET(f,LAMBDA(x,IF(x="","",XLOOKUP(x,[Щит],[Питается от (итог)],"")&"")),п0,[@Щит],п1,f(п0),п2,f(п1),п3,f(п2),п4,f(п3),'
        'п5,f(п4),п6,f(п5),п7,f(п6),"|"&TEXTJOIN("|",TRUE,п7,п6,п5,п4,п3,п2,п1,п0)&"|")')
SUB = 'SUMIFS([{c}],[Путь],"*|"&[@Щит]&"|*")'          # сумма по поддереву (сам щит + все нижние)
ONPATH = '--ISNUMBER(SEARCH("|"&[Щит]&"|",[@Путь]))'     # щиты на пути от ГРЩ до этого щита
scol("Щит", "in", note="Имя щита (как в «Линии» столбец «Щит»).", grp="СХЕМА ЩИТА (ввод для каждого щита)")
scol("Питается от", "in", note="Обычно пусто: определяется сам по линии-фидеру в «Линиях». Вписывайте, только если фидера в книге нет. Для ВРУ / ГРЩ — пусто.")
scol("Вводов", "in", note="Число вводов 1 / 2 / 3.")
scol("АВР", "in", note="да / нет")
scol("ИБП от", "in", note="Ввод 1 / 2 / 3 — от какого ввода питается ИБП (ЩГП).")
scol("Мест под вводной узел", "in", note="Первые места схемы под вводной узел (вводы, АВР, байпас) — рисуете сами. По умолчанию 15. Строки узла в схеме не рисуются (режим «рисовать» — нарисовать).")
scol("Питание: от (ТП / сеть)", "in", note="Только для ВРУ / ГРЩ: «ТП» — щит питается от своего трансформатора; иначе — от сети (введите Iкз).", grp="ТОЛЬКО ДЛЯ ВРУ / ГРЩ: сеть и питающий кабель (ввод)")
scol("ΔU до питающего, %", "in")
scol("Iкз min от сети, А", "in", note="Iкз min (1ф) на вводе — от сетевой организации.")
scol("Iкз max от сети, кА", "in", note="Iкз max (3ф) на вводе — от сетевой организации.")
scol("Sтр, кВА", "in")
scol("Схема тр-ра", "in", note="Д/Yн или Y/Yн")
scol("Питающий кабель: S, мм²", "in")
scol("Жил", "in")
scol("Кабелей параллельно", "in")
scol("Длина, м", "in")
scol("Материал", "in", note="медь / алюминий")
scol("In питающего аппарата, А", "in")
scol("Хар-ка аппарата", "in", note="B / C / D")
scol("Кат. питающей линии (ручн.)", "in", note="Для щита без вышестоящего в книге: какую категорию даёт питающая линия. Пусто — по своей схеме.")
scol("Путь", "h", PATH, note="Цепочка щитов от ГРЩ до этого щита.")
scol("Корень", "h", '=TEXTBEFORE(MID([@Путь],2,999),"|")')
scol("Питается от (итог)", "c", '=IF([@[Питается от]]<>"",[@[Питается от]],IFERROR(INDEX(Линии[Щит],[@[_строка фидера]]),""))',
     note="Вышестоящий щит: вручную или по линии-фидеру.", grp="РЕЗУЛЬТАТ (считается сам)")
scol("Уровень", "c", '=LEN([@Путь])-LEN(SUBSTITUTE([@Путь],"|",""))-2')
scol("_строка фидера", "h", '=IFERROR(XMATCH([@Щит],Линии[_фидер]),"")')
scol("Линия фидера", "c", '=IF([@[_строка фидера]]="","",INDEX(Линии[Номер линии],[@[_строка фидера]]))')
for h, c in (("_Руст собств", "_Руст (без Нет)"), ("_быт собств", "_Руст быт"), ("_сил собств", "_Рр сил"), ("_пост собств", "_Руст пост"),
             ("_P собств", "_P для cos"), ("_Pcos собств", "_P·cos"), ("_отоп собств", "_Р отопления"), ("_внг собств", "_Рр ВНГ")):
    scol(h, "h", f'=SUMIFS(Линии[{c}],Линии[Щит],[@Щит])')
scol("Руст, кВт", "c", "=" + SUB.format(c="_Руст собств"), note="Сумма установленной мощности щита и всех нижестоящих (без группы «Нет»).")
scol("Руст быт, кВт", "c", "=" + SUB.format(c="_быт собств"))
scol("Кс быт", "c",
     '=LET(x,[@[Руст быт, кВт]],m,MATCH(x,Кс_Р,1),IF(x>=INDEX(Кс_Р,ROWS(Кс_Р)),INDEX(Кс_К,ROWS(Кс_К)),INDEX(Кс_К,m)+(x-INDEX(Кс_Р,m))*'
     '(INDEX(Кс_К,m+1)-INDEX(Кс_К,m))/(INDEX(Кс_Р,m+1)-INDEX(Кс_Р,m))))',
     note="СП 256 табл. 7.2 по сумме Руст быт щита, линейная интерполяция.")
scol("Рр сил, кВт", "c", "=" + SUB.format(c="_сил собств"), note="Σ P·Кс силовых линий.")
scol("Руст пост, кВт", "c", "=" + SUB.format(c="_пост собств"))
scol("Р отопления, кВт", "c", "=" + SUB.format(c="_отоп собств"), note="Электроотопление, Кс = 1 (СП 256, Изм. 9).")
scol("Рр ВНГ, кВт", "c", "=" + SUB.format(c="_внг собств"), note="Водонагреватели × 0,7 (СП 256, Изм. 9).")
scol("_n силовых", "h", '=COUNTIFS(Линии[Щит],[@Щит],Линии[Группа Кс (итог)],"Сил")')
scol("Кс силовых (табл. 7.5)", "c", '=LET(n,[@[_n силовых]],IF(n<=1,1,IF(n>=MAX(Кс75_n),INDEX(Кс75_k,ROWS(Кс75_k)),'
     'LET(m,MATCH(n,Кс75_n,1),INDEX(Кс75_k,m)+(n-INDEX(Кс75_n,m))*(INDEX(Кс75_k,m+1)-INDEX(Кс75_k,m))/(INDEX(Кс75_n,m+1)-INDEX(Кс75_n,m))))))',
     note="СП 256 табл. 7.5 по числу силовых линий щита (доля сантехники 100–85 %; таблица — лист «Объект»). Применяется в режиме «нормы 2026», если Кс линии не задан.")
scol("Ко", "c", '=IF(Методика="нормы 2026",IF(AND([@[Руст быт, кВт]]>0,[@[Рр сил, кВт]]>0),Ко_сил,1),Ко_сил)',
     note="0,9 — совмещение бытовой и силовой нагрузки (СП 256 формула 6). В режиме «нормы 2026» — только если на щите есть обе.")
scol("Рр, кВт", "c", "=k_пк*[@[Руст быт, кВт]]*[@[Кс быт]]+[@Ко]*[@[Рр сил, кВт]]+[@[Руст пост, кВт]]+[@[Р отопления, кВт]]+[@[Рр ВНГ, кВт]]",
     note="Рр = kп.к·Руст быт·Кс + Ко·Рр сил + Руст пост + Р отоп + 0,7·Р ВНГ (СП 256 п. 7.1.10, Изм. 9).")
scol("Кс итог", "c", '=IF([@[Руст, кВт]]=0,0,[@[Рр, кВт]]/[@[Руст, кВт]])')
scol("cos φ", "c", '=IFERROR(' + SUB.format(c="_Pcos собств") + "/" + SUB.format(c="_P собств") + ',0.95)')
for k in (1, 2, 3):
    scol(f"P L{k}, кВт", "h",
         f'=SUMIFS(Линии[Р, кВт],Линии[Щит],[@Щит],Линии[Фаза],"L1,L2,L3",Линии[Группа Кс (итог)],"<>Нет")/3+'
         f'SUMIFS(Линии[Р, кВт],Линии[Щит],[@Щит],Линии[Фаза],"L{k}",Линии[Группа Кс (итог)],"<>Нет")')
scol("Iр, А", "c",
     '=IFERROR(MAX(MAX([@[P L1, кВт]],[@[P L2, кВт]],[@[P L3, кВт]])*[@[Кс итог]]/((Uн_лин*1000/SQRT(3))/1000*[@[cos φ]]),'
     '[@[Рр, кВт]]/(SQRT(3)*Uн_лин*[@[cos φ]])),0)', note="По наиболее нагруженной фазе.")
# ΔU и Z по пути
scol("_ΔU фидера", "h", '=IF([@[_строка фидера]]="","",INDEX(Линии[ΔU, %],[@[_строка фидера]]))')
scol("_ΔU пит. линии", "h",
     '=IF(OR(N([@[Питающий кабель: S, мм²]])=0,N([@[Длина, м]])=0),0,(IF([@Материал]="алюминий",0.036,0.0225)*N([@[Длина, м]])/'
     '(N([@[Питающий кабель: S, мм²]])*MAX(1,N([@[Кабелей параллельно]])))*[@[cos φ]]+0.00008*N([@[Длина, м]])/MAX(1,N([@[Кабелей параллельно]]))*'
     'SIN(ACOS([@[cos φ]])))*[@[Iр, А]]/(Uн_лин*1000/SQRT(3))*100)')
scol("_ΔU источника", "h", '=IF([@[Питается от (итог)]]<>"",0,IF(LEFT(UPPER(TRIM([@[Питание: от (ТП / сеть)]])),2)="ТП",0,N([@[ΔU до питающего, %]])+[@[_ΔU пит. линии]]))')
scol("ΔU до шин, %", "c", f'=SUMPRODUCT({ONPATH},IFERROR(--[_ΔU фидера],0)+IFERROR(--[_ΔU источника],0))',
     note="ΔU от ВРУ до шин щита = ΔU источника + Σ ΔU фидеров на пути.")
scol("_петля пит. кабеля", "h",
     '=IF(N([@[Питающий кабель: S, мм²]])>0,2*IF([@Материал]="алюминий",ρкз_Al,ρкз_Cu)*N([@[Длина, м]])/(N([@[Питающий кабель: S, мм²]])*MAX(1,N([@[Кабелей параллельно]]))),0)')
scol("_Zсети", "h",
     '=IF([@[Питается от (итог)]]<>"","",IF(N([@[Iкз min от сети, А]])>0,c_min*(Uн_лин*1000/SQRT(3))/[@[Iкз min от сети, А]],'
     'IF(AND(LEFT(UPPER(TRIM([@[Питание: от (ТП / сеть)]])),2)="ТП",N([@[Sтр, кВА]])>0),0.055*Uн_лин^2/([@[Sтр, кВА]]/1000)*IF([@[Схема тр-ра]]="Y/Yн",3,1),"")))')
scol("_Zисточника", "h", '=IF([@[_Zсети]]="","",[@[_Zсети]]+[@[_петля пит. кабеля]])')
scol("_Zфидера", "h", '=IF([@[_строка фидера]]="","",INDEX(Линии[_Zкабеля],[@[_строка фидера]]))')
scol("Zпетли на шинах, Ом", "c",
     f'=IF(XLOOKUP([@Корень],[Щит],[_Zисточника],"")="","",SUMPRODUCT({ONPATH},IFERROR(--[_Zисточника],0)+IFERROR(--[_Zфидера],0))+SUMPRODUCT({ONPATH})*R_пер)',
     note="Zшин = Zсети + Σ(петля фидеров + Rпер) на пути (ГОСТ 28249).")
scol("Iкз min на шинах, А", "c", '=IF(ISNUMBER([@[Zпетли на шинах, Ом]]),c_min*(Uн_лин*1000/SQRT(3))/[@[Zпетли на шинах, Ом]],"")')
scol("_Zmax сети", "h",
     '=IF([@[Питается от (итог)]]<>"","",IF(N([@[Iкз max от сети, кА]])>0,1.05*(Uн_лин*1000/SQRT(3))/([@[Iкз max от сети, кА]]*1000),'
     'IF(AND(LEFT(UPPER(TRIM([@[Питание: от (ТП / сеть)]])),2)="ТП",N([@[Sтр, кВА]])>0),0.055*Uн_лин^2/([@[Sтр, кВА]]/1000),"")))')
scol("_Zmax источника", "h",
     '=IF([@[_Zmax сети]]="","",[@[_Zmax сети]]+IF(N([@[Питающий кабель: S, мм²]])>0,IF([@Материал]="алюминий",0.028,0.0175)*N([@[Длина, м]])/'
     '(N([@[Питающий кабель: S, мм²]])*MAX(1,N([@[Кабелей параллельно]]))),0))')
scol("_Zmax фидера", "h", '=IF([@[_строка фидера]]="","",INDEX(Линии[_Zmax кабеля],[@[_строка фидера]]))')
scol("Zmax на шинах, Ом", "h",
     f'=IF(XLOOKUP([@Корень],[Щит],[_Zmax источника],"")="","",SUMPRODUCT({ONPATH},IFERROR(--[_Zmax источника],0)+IFERROR(--[_Zmax фидера],0)))')
scol("Iкз max на шинах, кА", "c", '=IF(ISNUMBER([@[Zmax на шинах, Ом]]),1.05*(Uн_лин*1000/SQRT(3))/[@[Zmax на шинах, Ом]]/1000,"")')
# категории по пути от ГРЩ (наследование: один ввод — не выше II)
scol("Категория шины «Ввод»", "c",
     '=LET(w,TEXTSPLIT([@Путь],"|"),names,FILTER(w,w<>""),REDUCE("",names,LAMBDA(acc,nm,LET(inp,N(XLOOKUP(nm,[Щит],[Вводов],1)),'
     'avr,LOWER(TRIM(XLOOKUP(nm,[Щит],[АВР],"нет")&"")),own,IF(inp>=2,IF(avr="да","I","II"),"III"),fr,XLOOKUP(nm,[Щит],[_строка фидера],""),'
     'bt,IF(fr="","",INDEX(Линии[Тип шины],fr)&""),man,TRIM(XLOOKUP(nm,[Щит],[Кат. питающей линии (ручн.)],"")&""),'
     'up,IF(acc="",man,IF(OR(bt="ИБП",bt="ДГУ"),"I особая",IF(bt="Одиночная","III",acc))),'
     'IF(own<>"III",own,IF(OR(up="I особая",up="I",up="II"),"II","III"))))))',
     note="1 ввод — III (или II, если питающая линия I/II: один кабель не даёт I, ПУЭ 1.2.19); 2+ ввода без АВР — II; с АВР — I.")
scol("_ранг собств", "h", '=MAXIFS(Линии[_ранг для фидера],Линии[Щит],[@Щит])')
scol("Высшая категория нагрузки", "c", '=IFERROR(INDEX({"III","II","I","I особая"},MAX(1,MAXIFS([_ранг собств],[Путь],"*|"&[@Щит]&"|*"))),"III")')
scol("Категория фидера", "c", '=IF([@[_строка фидера]]="","",INDEX(Линии[Категория (итог)],[@[_строка фидера]]))')
scol("Iпуск max, А", "h", '=MAXIFS(Линии[Iпуск, А],Линии[Щит],[@Щит])')
scol("Источник фидера", "h", '=IF([@[_строка фидера]]="","",INDEX(Линии[Источник для режимов],[@[_строка фидера]]))')
scol("Пожарный фидер", "h", '=IF([@[_строка фидера]]="","",INDEX(Линии[Пожарная (итог)],[@[_строка фидера]]))')
scol("Не обеспечено, линий", "c", '=COUNTIFS(Линии[Щит],[@Щит],Линии[Категория: проверка],"НЕ ОБЕСПЕЧЕНА")')
scol("Замечания щита", "c",
     '=TEXTJOIN("; ",TRUE,IF(AND([@[Питается от (итог)]]="",[@[Zпетли на шинах, Ом]]=""),"нет данных сети для ТКЗ: введите Iкз от сетевой организации или трансформатор",""),'
     'IF(AND([@[Категория фидера]]<>"",IFERROR(MATCH([@[Категория фидера]],{"III","II","I","I особая"},0),1)<IFERROR(MATCH([@[Высшая категория нагрузки]],{"III","II","I","I особая"},0),1)),'
     '"фидер "&[@[Категория фидера]]&" ниже нагрузки "&[@[Высшая категория нагрузки]],""),'
     'IF([@[Не обеспечено, линий]]>0,"не обеспечена категория у "&[@[Не обеспечено, линий]]&" линий",""),'
     'IF(AND([@[Питается от]]<>"",[@[_строка фидера]]=""),"нет линии-фидера в «Линиях»: номер линии = имя щита или столбец «Питает щит»",""),'
     'IF(AND(ISNUMBER([@[ΔU до шин, %]]),[@[ΔU до шин, %]]>4),"ΔU до шин "&FIXED([@[ΔU до шин, %]],2)&" % > 4 %",""))')


def F2(rng, f):
    """записать формулу; при ошибке - в журнал (сборка продолжается)."""
    try:
        rng.Formula2 = f
    except Exception as e:
        log("ОШИБКА формулы %s!%s: %s" % (rng.Worksheet.Name, rng.Address, f[:250]))


# ------------------------------------------------------------------------------------------------
# Лист «Нагрузки»: режимы по источникам, пожар, категории, ИБП для выбранного щита (вместе с нижестоящими)
# ------------------------------------------------------------------------------------------------
UPS_SET = [("Допустимая загрузка ИБП", 0.75, "", "ИБП_загрузка"), ("cos φ на выходе ИБП", 0.9, "", "ИБП_cos_вых"),
           ("Автономия, мин", 15, "мин", "ИБП_t"), ("Напряжение АКБ, В", 240, "В", "ИБП_U"), ("КПД инвертора", 0.95, "", "ИБП_КПД"),
           ("Коэффициент разряда / старения АКБ", 0.8, "", "ИБП_k"), ("Заряд АКБ, доля мощности", 0.1, "", "ИБП_заряд"),
           ("cos φ на входе ИБП", 0.95, "", "ИБП_cos_вх")]
UPS_SIZES = [6, 10, 15, 20, 30, 40, 60, 80, 100, 120, 160, 200]

RASCHET = ('LAMBDA(m,ap,aq,LET(byt,SUM(m*Линии[_Руст быт]),sil,SUM(m*Линии[_Рр сил]),post,SUM(m*Линии[_Руст пост]),'
           'otop,SUM(m*Линии[_Р отопления]),vng,SUM(m*Линии[_Рр ВНГ]),pall,SUM(m*Линии[_Руст (без Нет)]),'
           'pw,SUM(m*Линии[_P для cos]),pc,SUM(m*Линии[_P·cos]),kc,LET(x,byt,mm,MATCH(x,Кс_Р,1),IF(x>=INDEX(Кс_Р,ROWS(Кс_Р)),INDEX(Кс_К,ROWS(Кс_К)),'
           'INDEX(Кс_К,mm)+(x-INDEX(Кс_Р,mm))*(INDEX(Кс_К,mm+1)-INDEX(Кс_К,mm))/(INDEX(Кс_Р,mm+1)-INDEX(Кс_Р,mm)))),'
           'ko,IF(Методика="нормы 2026",IF(AND(byt>0,sil>0),Ко_сил,1),Ко_сил),pr0,k_пк*byt*kc+ko*sil+post+otop+vng,'
           'c0,IF(pw=0,0.95,pc/pw),q0,pr0*TAN(ACOS(c0)),pr,pr0+ap,q,q0+aq,sr,SQRT(pr^2+q^2),'
           'VSTACK(pall,IF(pall=0,0,pr/pall),IF(sr=0,0,pr/sr),IF(pr=0,0,q/pr),pr,q,sr,sr/(SQRT(3)*Uн_лин))))')
KSBYT = ('=LAMBDA(x,LET(m,MATCH(x,Кс_Р,1),IF(x>=INDEX(Кс_Р,ROWS(Кс_Р)),INDEX(Кс_К,ROWS(Кс_К)),'
         'INDEX(Кс_К,m)+(x-INDEX(Кс_Р,m))*(INDEX(Кс_К,m+1)-INDEX(Кс_К,m))/(INDEX(Кс_Р,m+1)-INDEX(Кс_Р,m)))))')
ROWS8 = ["Руст, кВт", "Кс (Рр / Руст)", "cos φ", "tg φ", "Рр, кВт", "Qр, квар", "Sр, кВА", "Iр, А"]
SRC = ["Ввод 1", "Ввод 2", "Ввод 3", "ИБП", "ДГУ"]


def build_loads(wb, xl, root, sep):
    ws = wb.Worksheets.Add(After=wb.Worksheets("Щиты"))
    ws.Name = "Нагрузки"
    def R(f):
        """формула с вызовами РАСЧЁТ(...) -> LET с локальной LAMBDA (COM не даёт создать именованную LAMBDA)."""
        assert f.startswith("=")
        return "=LET(РАСЧЁТ," + RASCHET + "," + f[1:] + ")"

    wo = wb.Worksheets("Объект")
    # настройки ИБП на «Объекте»
    wo.Range("B14").Value = "ИБП (для щитов с шиной ИБП)"
    wo.Range("B14").Font.Bold = True
    names = {}
    for i, (a, v, u, nm) in enumerate(UPS_SET):
        r = 15 + i
        wo.Range(f"B{r}").Value = a
        wo.Range(f"C{r}").Value = v
        wo.Range(f"D{r}").Value = u
        wo.Range(f"C{r}").Interior.Color = C_IN
        names[nm] = f"=Объект!$C${r}"
    r0 = 15 + len(UPS_SET) + 1
    wo.Range(f"B{r0}").Value = "Ряд типоразмеров ИБП, кВА"
    wo.Range(f"C{r0}:{letters(2 + len(UPS_SIZES))}{r0}").Value = [UPS_SIZES]
    wo.Range(f"C{r0}:{letters(2 + len(UPS_SIZES))}{r0}").Interior.Color = C_IN
    names["ИБП_ряд"] = f"=Объект!$C${r0}:${letters(2 + len(UPS_SIZES))}${r0}"
    set_names(wb, names)

    ws.Range("B1").Value = "РЕЖИМЫ: нагрузка выбранного щита (вместе со всеми нижестоящими) по вводам в разных режимах"
    ws.Range("E3").Value = ("Как читать: столбцы — источники (Ввод 1/2/3, ИБП, ДГУ), блоки — режимы. На «Нормальный» и «Аварийный» "
                            "проверяются вводы, аппараты и кабели (берётся больший). Пожарные — отдельно, в максимум не входят. Блок 7 — кВт по категориям для ТУ.")
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    ws.Range("B3").Value = "Щит:"
    ws.Range("C3").Value = root
    ws.Range("C3").Interior.Color = C_IN
    v = ws.Range("C3").Validation
    v.Delete()
    v.Add(3, 1, 1, "=Список_щитов")
    F2(ws.Range("D3"), '="Вводов: "&XLOOKUP($C$3,Щиты[Щит],Щиты[Вводов],"")&", АВР: "&XLOOKUP($C$3,Щиты[Щит],Щиты[АВР],"")&", ИБП питается от: "&XLOOKUP($C$3,Щиты[Щит],Щиты[ИБП от],"Ввод 1")')
    # служебные массивы (столбцы AA:AE, скрыты)
    ws.Range("AA2:AE2").Value = [["в составе", "нижний щит", "источник", "пожарная", "категория"]]
    F2(ws.Range("AA3"), '=--(ISNUMBER(SEARCH("|"&$C$3&"|",Линии[_путь щита]))*(Линии[Номер линии]<>""))')
    F2(ws.Range("AB3"), '=IFERROR(TEXTBEFORE(TEXTAFTER(Линии[_путь щита],"|"&$C$3&"|"),"|"),"")')
    F2(ws.Range("AC3"), '=IF(AB3#="",Линии[Источник для режимов],XLOOKUP(AB3#,Щиты[Щит],Щиты[Источник фидера],"Ввод 1"))')
    F2(ws.Range("AD3"), '=IF(AB3#="",Линии[Пожарная (итог)],IF(XLOOKUP(AB3#,Щиты[Щит],Щиты[Пожарный фидер],"нет")="да","да",Линии[Пожарная (итог)]))')
    F2(ws.Range("AE3"), '=Линии[Категория (итог)]')
    ws.Columns("AA:AE").Hidden = True
    # вход ИБП (нагрузка на тот ввод, от которого питается ИБП)
    ws.Range("B5").Value = "Вход ИБП, кВт / квар:"
    F2(ws.Range("C5"), R('=LET(r,РАСЧЁТ(AA3#*(AC3#="ИБП")*(AD3#="нет"),0,0),INDEX(r,5)/ИБП_КПД*(1+ИБП_заряд))'))
    F2(ws.Range("D5"), '=C5*TAN(ACOS(ИБП_cos_вх))')
    F2(ws.Range("E5"), '=XLOOKUP($C$3,Щиты[Щит],Щиты[ИБП от],"Ввод 1")&""')
    ws.Range("E5").Font.Color = 8421504

    def block(r, title, cols):
        ws.Range(f"B{r}").Value = title
        ws.Range(f"B{r}").Font.Bold = True
        heads = ["Показатель"] + [c[0] for c in cols]
        ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 1 + len(heads))).Value = [heads]
        ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 1 + len(heads))).Interior.Color = C_HEAD
        ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 1 + len(heads))).Font.Bold = True
        ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + 1, 1 + len(heads))).HorizontalAlignment = -4152
        ws.Range(ws.Cells(r + 2, 2), ws.Cells(r + 9, 2)).Value = [[x] for x in ROWS8]
        for j, (_, f) in enumerate(cols):
            F2(ws.Cells(r + 2, 3 + j), R(f))
        ws.Range(ws.Cells(r + 2, 3), ws.Cells(r + 9, 2 + len(cols))).NumberFormatLocal = "0,00"
        return r + 11

    def mk(cond, add=None):
        ap, aq = ("0", "0") if add is None else add
        return f"=РАСЧЁТ(AA3#*{cond},{ap},{aq})"

    nf = '(AD3#="нет")'
    up = lambda srcs: ('IF(OR(' + ",".join(f'$E$5="{x}"' for x in srcs) + '),$C$5,0)', 'IF(OR(' + ",".join(f'$E$5="{x}"' for x in srcs) + '),$D$5,0)')
    vv = '((AC3#="Ввод 1")+(AC3#="Ввод 2")+(AC3#="Ввод 3"))'
    r = 7
    r = block(r, "1. НОРМАЛЬНЫЙ РЕЖИМ (секционный выключатель отключён, пожарные не входят; ИБП — нагрузка на выходе, справочно)",
              [(x, mk(f'(AC3#="{x}")*{nf}', up([x]) if x.startswith("Ввод") else None)) for x in SRC] +
              [("Итого (вводы)", mk(f"{vv}*{nf}", up(["Ввод 1", "Ввод 2", "Ввод 3"])))])
    r = block(r, "2. АВАРИЙНЫЙ: потерян Ввод 1 (его нагрузка и вход ИБП — на Ввод 2)",
              [("Ввод 2", mk(f'((AC3#="Ввод 1")+(AC3#="Ввод 2"))*{nf}', up(["Ввод 1", "Ввод 2"]))),
               ("Ввод 3", mk(f'(AC3#="Ввод 3")*{nf}', up(["Ввод 3"])))])
    r = block(r, "3. АВАРИЙНЫЙ: потерян Ввод 2 (его нагрузка и вход ИБП — на Ввод 1)",
              [("Ввод 1", mk(f'((AC3#="Ввод 1")+(AC3#="Ввод 2"))*{nf}', up(["Ввод 1", "Ввод 2"]))),
               ("Ввод 3", mk(f'(AC3#="Ввод 3")*{nf}', up(["Ввод 3"])))])
    r = block(r, "4. ПОТЕРЯНЫ ВСЕ ВВОДЫ: работают только ИБП и ДГУ",
              [("ИБП", mk(f'(AC3#="ИБП")*{nf}')), ("ДГУ", mk(f'(AC3#="ДГУ")*{nf}'))])
    r = block(r, "5. ПОЖАР: пожарные нагрузки (в расчётный максимум не входят, СП 256 п. 7.1.9; Кс пожарных = 1)",
              [(x, mk(f'(AC3#="{x}")*(AD3#="да")')) for x in SRC] + [("Итого", mk('(AD3#="да")'))])
    r = block(r, "6. НОРМАЛЬНЫЙ + ПОЖАРНЫЕ (проверка вводов при пожаре)",
              [("Итого (вводы)", mk(f"{vv}", up(["Ввод 1", "Ввод 2", "Ввод 3"])))])
    # категории
    ws.Range(f"B{r}").Value = "7. ПО КАТЕГОРИЯМ НАДЁЖНОСТИ (для ТУ и пояснительной записки; ИБП — нагрузка на выходе)"
    ws.Range(f"B{r}").Font.Bold = True
    heads = ["Категория", "Линий", "Руст, кВт", "Рр, кВт", "Qр, квар", "Sр, кВА", "Пожарные: Руст, кВт", "Пожарные: Рр, кВт"]
    ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 9)).Value = [heads]
    ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 9)).Interior.Color = C_HEAD
    ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 9)).Font.Bold = True
    for i, cat in enumerate(["I особая", "I", "II", "III"]):
        rr = r + 2 + i
        ws.Cells(rr, 2).Value = cat
        F2(ws.Cells(rr, 3), f'=SUM(AA3#*(AE3#="{cat}"))')
        F2(ws.Cells(rr, 4), R(f'=LET(x,РАСЧЁТ(AA3#*(AE3#="{cat}")*{nf},0,0),HSTACK(INDEX(x,1),INDEX(x,5),INDEX(x,6),INDEX(x,7)))'))
        F2(ws.Cells(rr, 8), R(f'=LET(x,РАСЧЁТ(AA3#*(AE3#="{cat}")*(AD3#="да"),0,0),HSTACK(INDEX(x,1),INDEX(x,5)))'))
    rr = r + 6
    ws.Cells(rr, 2).Value = "Итого"
    F2(ws.Cells(rr, 3), '=SUM(AA3#)')
    F2(ws.Cells(rr, 4), R(f'=LET(x,РАСЧЁТ(AA3#*{nf},0,0),HSTACK(INDEX(x,1),INDEX(x,5),INDEX(x,6),INDEX(x,7)))'))
    F2(ws.Cells(rr, 8), R('=LET(x,РАСЧЁТ(AA3#*(AD3#="да"),0,0),HSTACK(INDEX(x,1),INDEX(x,5)))'))
    ws.Range(ws.Cells(r + 2, 4), ws.Cells(rr, 9)).NumberFormatLocal = "0,00"
    ws.Cells(rr + 1, 2).Value = "Кс быт применяется к каждому набору линий отдельно (табл. 7.2 по сумме Руст быт набора), поэтому сумма строк может отличаться от «Итого»."
    ws.Cells(rr + 1, 2).Font.Color = 8421504
    r = rr + 3
    # ИБП
    ws.Range(f"B{r}").Value = "8. ИБП ВЫБРАННОГО ЩИТА (нагрузка шин ИБП, без пожарных)"
    ws.Range(f"B{r}").Font.Bold = True
    ups = [("P на выходе, кВт", R('=INDEX(РАСЧЁТ(AA3#*(AC3#="ИБП")*(AD3#="нет"),0,0),5)')),
           ("Требуемая S по загрузке, кВА", f'=C{r + 1}/ИБП_cos_вых/ИБП_загрузка'),
           ("Типоразмер ИБП, кВА", f'=IFERROR(MIN(FILTER(ИБП_ряд,ИБП_ряд>=C{r + 2})),"больше ряда")'),
           ("Загрузка принятого ИБП", f'=IFERROR(C{r + 1}/ИБП_cos_вых/C{r + 3},"")'),
           ("Ёмкость АКБ, А·ч", f'=C{r + 1}*1000*ИБП_t/60/(ИБП_U*ИБП_КПД*ИБП_k)'),
           ("Вход ИБП, кВт (на ввод)", "=C5")]
    for i, (a, f) in enumerate(ups):
        ws.Cells(r + 1 + i, 2).Value = a
        F2(ws.Cells(r + 1 + i, 3), f)
        ws.Cells(r + 1 + i, 3).NumberFormatLocal = "0,00"
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C:I").ColumnWidth = 13
    return ws


# ------------------------------------------------------------------------------------------------
# «Расчёт_нагрузок»: таблица для пояснительной записки; «Кабельный_журнал»
# ------------------------------------------------------------------------------------------------
def masks(bcell, modecell=None):
    """LET-переменные: m - линии щита и нижестоящих (без пожарных), по режиму «Весь щит / Секция N»."""
    sec = "" if modecell is None else f',IF({modecell}="Весь щит",1,src=("Ввод "&RIGHT({modecell},1)))'
    return ('щит_,' + bcell + ',pth,Линии[_путь щита],m0,--ISNUMBER(SEARCH("|"&щит_&"|",pth))*(Линии[Номер линии]<>""),'
            'ch,IFERROR(TEXTBEFORE(TEXTAFTER(pth,"|"&щит_&"|"),"|"),""),'
            'src,IF(ch="",Линии[Источник для режимов],XLOOKUP(ch,Щиты[Щит],Щиты[Источник фидера],"Ввод 1")),'
            'fire,IF(ch="",Линии[Пожарная (итог)],IF(XLOOKUP(ch,Щиты[Щит],Щиты[Пожарный фидер],"нет")="да","да",Линии[Пожарная (итог)])),'
            'm,m0*(fire="нет")' + ('*' + sec[1:] if sec else '') + ',')


def build_pz(wb, root, sep):
    ws = wb.Worksheets.Add(After=wb.Worksheets("Нагрузки"))
    ws.Name = "Расчёт_нагрузок"
    ws.Range("A1").Value = "ТАБЛИЦА НАГРУЗОК ДЛЯ ПОЯСНИТЕЛЬНОЙ ЗАПИСКИ"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    for r, a, v in ((3, "Объект:", "(название и адрес объекта)"), (4, "Щит (ВРУ):", root), (5, "Режим:", "Весь щит")):
        ws.Range(f"B{r}").Value = a
        ws.Range(f"C{r}").Value = v
        ws.Range(f"C{r}").Interior.Color = C_IN
    v = ws.Range("C4").Validation
    v.Delete()
    v.Add(3, 1, 1, "=Список_щитов")
    v = ws.Range("C5").Validation
    v.Delete()
    v.Add(3, 1, 1, sep.join(["Весь щит", "Секция 1", "Секция 2", "Секция 3"]))
    # основные показатели (для текста записки)
    rows = [("Напряжение сети", '="~"&FIXED(Uн_лин*1000,0)&"/"&FIXED(Uн_лин*1000/SQRT(3),0)&" В, 50 Гц"'),
            ("Система заземления", "TN-C-S"),
            ("Категория надёжности электроснабжения", '=XLOOKUP($C$4,Щиты[Щит],Щиты[Категория шины «Ввод»],"")'),
            ("Установленная мощность Руст, кВт", "=FIXED(C16,2)"),
            ("Расчётная мощность Рр, кВт / Sр, кВА / cos φ", '=FIXED(G16,2)&" / "&FIXED(I16,2)&" / "&FIXED(IF(I16=0,0,G16/I16),2)'),
            ("Расчётный ток Iр, А", '=FIXED(J16,1)'),
            ("ΔU от ВРУ до шин щита, %", '=IFERROR(FIXED(XLOOKUP($C$4,Щиты[Щит],Щиты[ΔU до шин, %]),2),"—")'),
            ("Iкз min (1ф) / Iкз max (3ф) на шинах", '=IFERROR(FIXED(XLOOKUP($C$4,Щиты[Щит],Щиты[Iкз min на шинах, А]),0)&" А / "&FIXED(XLOOKUP($C$4,Щиты[Щит],Щиты[Iкз max на шинах, кА]),1)&" кА","нет данных сети")'),
            ("Компенсация реактивной мощности", '=IF(N(G16)=0,"",IF(G16*(IF(G16=0,0,H16/G16)-0.35)<50,"не требуется (< 50 квар на ввод, СП 256 п. 7.3)",ROUND(G16*(H16/G16-0.35),1)&" квар"))')]
    ws.Range("E3").Value = "ОСНОВНЫЕ ПОКАЗАТЕЛИ"
    ws.Range("E3").Font.Bold = True
    for i, (a, f) in enumerate(rows):
        ws.Cells(4 + i, 5).Value = a
        if str(f).startswith("="):
            F2(ws.Cells(4 + i, 7), f)
        else:
            ws.Cells(4 + i, 7).Value = f
    # итоговые строки (над таблицей, чтобы таблица могла расти вниз)
    head = ["№", "Наименование электроприёмника (вид нагрузки)", "Руст, кВт", "Кс", "cos φ", "tg φ", "Рр, кВт", "Qр, квар", "Sр, кВА", "Iр, А"]
    ws.Range("A14:J14").Value = [head]
    ws.Range("A14:J14").Font.Bold = True
    ws.Range("A14:J14").Interior.Color = C_HEAD
    ws.Range("A14:J14").WrapText = True
    base = '=LET(' + masks("$C$4", "$C$5")
    comp = ('byt,SUM(m*Линии[_Руст быт]),sil,SUM(m*Линии[_Рр сил]),post,SUM(m*Линии[_Руст пост]),otop,SUM(m*Линии[_Р отопления]),'
            'vng,SUM(m*Линии[_Рр ВНГ]),pall,SUM(m*Линии[_Руст (без Нет)]),pw,SUM(m*Линии[_P для cos]),pc,SUM(m*Линии[_P·cos]),'
            'kc,LET(x,byt,mm,MATCH(x,Кс_Р,1),IF(x>=INDEX(Кс_Р,ROWS(Кс_Р)),INDEX(Кс_К,ROWS(Кс_К)),INDEX(Кс_К,mm)+(x-INDEX(Кс_Р,mm))*(INDEX(Кс_К,mm+1)-INDEX(Кс_К,mm))/(INDEX(Кс_Р,mm+1)-INDEX(Кс_Р,mm)))),'
            'ko,IF(Методика="нормы 2026",IF(AND(byt>0,sil>0),Ко_сил,1),Ко_сил),cs_,IF(pw=0,0.95,pc/pw),')
    # строка 13: сумма строк таблицы; строка 14: расчётная нагрузка щита (с Ко и kп.к)
    ws.Range("B15").Value = "Итого (сумма строк)"
    F2(ws.Range("C15"), '=SUM(C17#)' if False else '=IFERROR(SUM(INDEX(A17#,0,3)),0)')
    F2(ws.Range("G15"), '=IFERROR(SUM(INDEX(A17#,0,7)),0)')
    F2(ws.Range("H15"), '=IFERROR(SUM(INDEX(A17#,0,8)),0)')
    F2(ws.Range("I15"), '=SQRT(G15^2+H15^2)')
    F2(ws.Range("J15"), '=I15/(SQRT(3)*Uн_лин)')
    ws.Range("B16").Value = "Расчётная нагрузка щита: Рр = kп.к·Руст быт·Кс + Ко·Рр сил + Руст пост + Р отоп + 0,7·Р ВНГ (СП 256)"
    ws.Range("B16").WrapText = True
    F2(ws.Range("C16"), base + comp + 'pr,k_пк*byt*kc+ko*sil+post+otop+vng,qv_,pr*TAN(ACOS(cs_)),HSTACK(pall,IF(pall=0,0,pr/pall),cs_,TAN(ACOS(cs_)),pr,qv_,SQRT(pr^2+qv_^2),SQRT(pr^2+qv_^2)/(SQRT(3)*Uн_лин)))')
    ws.Range("A15:J16").Font.Bold = True
    # таблица по видам нагрузки (Кс быт — по всему щиту, как в v3.12: строки складываются в итог)
    F2(ws.Range("A17"), (base + comp +
        'kinds,UNIQUE(FILTER(Линии[Вид нагрузки (итог)],m,"")),'
        'rows,REDUCE("",kinds,LAMBDA(acc,k,LET(mk,m*(Линии[Вид нагрузки (итог)]=k),bb_,SUM(mk*Линии[_Руст быт]),ss_,SUM(mk*Линии[_Рр сил]),'
        'pp_,SUM(mk*Линии[_Руст пост]),oo_,SUM(mk*Линии[_Р отопления]),vv_,SUM(mk*Линии[_Рр ВНГ]),aa_,SUM(mk*Линии[_Руст (без Нет)]),'
        'ww_,SUM(mk*Линии[_P для cos]),cc_,SUM(mk*Линии[_P·cos]),cs,IF(ww_=0,0.95,cc_/ww_),prr_,k_пк*bb_*kc+ss_+pp_+oo_+vv_,qq_,prr_*TAN(ACOS(cs)),'
        'VSTACK(acc,HSTACK(k,aa_,IF(aa_=0,0,prr_/aa_),cs,TAN(ACOS(cs)),prr_,qq_,SQRT(prr_^2+qq_^2),SQRT(prr_^2+qq_^2)/(SQRT(3)*Uн_лин)))))),'
        'b0,DROP(rows,1),body,FILTER(b0,(CHOOSECOLS(b0,2)<>0)+(CHOOSECOLS(b0,6)<>0),b0),HSTACK(SEQUENCE(ROWS(body)),body))'))
    ws.Range("C15:J200").NumberFormatLocal = "0,00"
    ws.Range("A15:A200").NumberFormatLocal = "0"
    for c_, w_ in (("A", 5), ("B", 48), ("C", 11), ("D", 8), ("E", 8), ("F", 8), ("G", 11), ("H", 11), ("I", 11), ("J", 11)):
        ws.Columns(c_).ColumnWidth = w_
    ws.Range("B18").Value = None
    ws.Range("L3").Value = ("Как считается: строка = вид нагрузки (столбец «Вид нагрузки» в «Линиях», пусто — по коду). Кс бытовой нагрузки — "
                            "по таблице СП 256 по сумме Руст быт ВСЕГО щита, поэтому строки складываются в «Итого». Строка «Расчётная нагрузка» "
                            "учитывает Ко = 0,9 и kп.к. Пожарные нагрузки не входят (СП 256 п. 7.1.9). Для записки: выберите щит и режим, "
                            "скопируйте A14:J… и вставьте как значения.")
    ws.Range("L3").WrapText = True
    ws.Columns("L").ColumnWidth = 60
    return ws


# ------------------------------------------------------------------------------------------------
# Доработки по замечаниям: гайды на листах, итоги щитов, журнал, бирки, «Однолинейка», «Таблицы_норм»
# ------------------------------------------------------------------------------------------------
def guide_block(ws, r0, title, lines, width_cols=8):
    """блок пояснений на листе: заголовок + строки текста (объединённые по ширине)."""
    ws.Cells(r0, 2).Value = title
    ws.Cells(r0, 2).Font.Bold = True
    ws.Cells(r0, 2).Font.Size = 12
    ws.Range(ws.Cells(r0, 2), ws.Cells(r0, 1 + width_cols)).Interior.Color = 14857357
    for i, t in enumerate(lines):
        rng = ws.Range(ws.Cells(r0 + 1 + i, 2), ws.Cells(r0 + 1 + i, 1 + width_cols))
        rng.Merge()
        rng.WrapText = True
        rng.Value = t
        rng.VerticalAlignment = -4160
        ws.Rows(r0 + 1 + i).RowHeight = max(15, 15 * -(-len(t) // 150))
    return r0 + 1 + len(lines)


OBJ_GUIDE = [
    "Лист задаёт настройки ВСЕГО объекта. Обычно меняют только C3 (название), C8 (методика) и C9 (kп.к). Остальное — по необходимости.",
    "C3 Объект — название и адрес. Идёт в таблицу для пояснительной записки.",
    "C4–C7 Ток КЗ: ρ меди/алюминия при нагреве, переходные сопротивления на каждый щит, c_min. Влияют на Iкз min в конце линий и проверку отключения ≤ 0,4 с. Значения по ГОСТ 28249 — не меняйте без причины.",
    "C8 Методика: «нормы 2026» — расчёт по СП 256 с Изм. 9 (cos φ силовых 0,8, Кс силовых по табл. 7.5, Ко только при бытовой + силовой нагрузке, отопление и ВНГ отдельно). «v3.12» — как старая книга, только для сверки.",
    "C9 kп.к — региональный коэффициент к бытовой нагрузке (СП 256 табл. 7.5а). Частный дом — 1. Многоквартирный дом — значение региона (ЦФО 0,81).",
    "F2:G14 Кс силовых по числу электроприёмников (СП 256 табл. 7.5, доля сантехники 100–85 %). Применяется к силовым линиям без ручного Кс в режиме «нормы 2026».",
    "B15:C22 ИБП — допустимая загрузка, cos φ, автономия, напряжение АКБ, КПД, разряд, заряд. Влияют на подбор ИБП и на нагрузку входа ИБП на вводе (лист «Режимы»).",
    "Где ещё настройки: Uн, общий К прокладки, таблица Кс быт, мест на лист схемы — лист «Таблицы_норм» (BF4, BQ29, BQ30, BP4:BQ11).",
]
SHIELD_GUIDE = [
    "ЩИТЫ: одна строка — один щит объекта. Жёлтые столбцы — ввод, серые — результат. Свёрнутые столбцы (кнопка «+» над таблицей) — подробности расчёта.",
    "Для КАЖДОГО щита: Щит (имя, как в «Линиях»), Вводов (1/2/3), АВР (да/нет), ИБП от (Ввод 1/2/3), Мест под вводной узел (обычно 15).",
    "«Питается от» обычно НЕ заполняйте: книга сама находит линию-фидер в «Линиях» (номер линии = имя щита, или столбец «Питает щит»).",
    "Только для ВРУ / ГРЩ (щит без вышестоящего): данные сети — Iкз min и Iкз max от сетевой организации, или «ТП» и мощность трансформатора; питающий кабель и аппарат. Без них ток КЗ не считается ни в одном щите.",
    "Результат: Руст, Рр, Iр, cos φ — щит вместе со всеми нижестоящими; ΔU до шин и Iкз — от ВРУ по фидерам; «Категория шины «Ввод»» — какую категорию даёт схема; «Замечания щита» — что исправить.",
]
BUS_GUIDE = [
    "ШИНЫ: нужен, только если тип шины НЕ определяется сам. Обычно лист пустой.",
    "Как определяется сам: в «Линиях» в столбце «Шина» (в строке с автоматом) пишется имя шины. Имя содержит «ИБП» → шина ИБП (категория I особая); «ДГУ» → шина ДГУ; иначе — основная шина секции линии (Ввод 1/2/3 по столбцу «Секция»).",
    "Когда заполнять: имя шины не говорит о её типе. Пример: шина «ЩГП» питается от ИБП → строка: Шина = ЩГП, Щит = ЩР-1 (или пусто — для всех щитов), Тип = ИБП. Пример: «Шина вентиляции» без резерва → Тип = Одиночная (категория III).",
    "Типы: Ввод 1 / Ввод 2 / Ввод 3 — источник по секции; ИБП, ДГУ — отдельный источник (I особая); Одиночная — без резерва (III). Тип влияет на категорию линий и на режимы (лист «Режимы»).",
]


def put_guides(wb):
    """гайды на листах «Объект» (снизу), «Щиты» и «Шины» (сверху, над таблицей)."""
    wo = wb.Worksheets("Объект")
    guide_block(wo, 28, "КАК НАСТРАИВАТЬ ЛИСТ «ОБЪЕКТ»", OBJ_GUIDE, 3)
    for nm, lines, nrow in (("Щиты", SHIELD_GUIDE, 7), ("Шины", BUS_GUIDE, 6)):
        ws = wb.Worksheets(nm)
        ws.Rows(f"1:{nrow}").Insert()
        guide_block(ws, 1, "КАК ЗАПОЛНЯТЬ ЛИСТ «" + nm.upper() + "»", lines, 12)


def build_totals(wb, root):
    """«Итоги_щитов»: карточка выбранного щита (как блок «Нагрузка_щитов» v3.12) + таблица по всем щитам."""
    ws = wb.Worksheets.Add(After=wb.Worksheets("Щиты"))
    ws.Name = "Итоги_щитов"
    ws.Range("B1").Value = "ИТОГИ ПО ЩИТАМ (данные под однолинейку)"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    ws.Range("B3").Value = "Щит:"
    ws.Range("C3").Value = root
    ws.Range("C3").Interior.Color = C_IN
    v = ws.Range("C3").Validation
    v.Delete()
    v.Add(3, 1, 1, "=Список_щитов")
    X = lambda h: f'XLOOKUP($C$3,Щиты[Щит],Щиты[{h}],"")'
    card = [("Питается от", X("Питается от (итог)"), ""), ("Руст", X("Руст, кВт"), "кВт"), ("  в т. ч. бытовая", X("Руст быт, кВт"), "кВт"),
            ("Кс быт (табл. 7.2)", X("Кс быт"), ""), ("Рр силовой (Σ P·Кс)", X("Рр сил, кВт"), "кВт"), ("Руст постоянной", X("Руст пост, кВт"), "кВт"),
            ("Р отопления (Кс 1)", X("Р отопления, кВт"), "кВт"), ("Рр ВНГ (×0,7)", X("Рр ВНГ, кВт"), "кВт"), ("Ко", X("Ко"), ""),
            ("Рр", X("Рр, кВт"), "кВт"), ("Кс итог = Рр / Руст", X("Кс итог"), ""), ("cos φ", X("cos φ"), ""),
            ("Qр = Рр·tg φ", '=IFERROR(C13*TAN(ACOS(C15)),"")', "квар"), ("Sр", '=IFERROR(SQRT(C13^2+C16^2),"")', "кВА"), ("Iр (по наиб. нагруженной фазе)", X("Iр, А"), "А"),
            ("P L1 / L2 / L3", '=IFERROR(FIXED(' + X("P L1, кВт") + ',2)&" / "&FIXED(' + X("P L2, кВт") + ',2)&" / "&FIXED(' + X("P L3, кВт") + ',2),"")', "кВт"),
            ("L1р / L2р / L3р (P·Кс итог)", '=IFERROR(FIXED(' + X("P L1, кВт") + '*C14,2)&" / "&FIXED(' + X("P L2, кВт") + '*C14,2)&" / "&FIXED(' + X("P L3, кВт") + '*C14,2),"")', "кВт"),
            ("Перекос фаз (max−min)/max", '=IFERROR(LET(a,' + X("P L1, кВт") + ',b,' + X("P L2, кВт") + ',cc,' + X("P L3, кВт") + ',(MAX(a,b,cc)-MIN(a,b,cc))/MAX(a,b,cc)*100),"")', "%"),
            ("Наибольший пусковой ток", X("Iпуск max, А"), "А"),
            ("Компенсация реактивной мощности", '=IFERROR(IF((TAN(ACOS(C15))-0.35)*C13<50,"не требуется (< 50 квар, СП 256 п. 7.3)",FIXED((TAN(ACOS(C15))-0.35)*C13,1)&" квар"),"")', ""),
            ("ΔU от ВРУ до шин", X("ΔU до шин, %"), "%"), ("Iкз min (1ф) на шинах", X("Iкз min на шинах, А"), "А"), ("Iкз max (3ф) на шинах", X("Iкз max на шинах, кА"), "кА"),
            ("Категория шины «Ввод»", X("Категория шины «Ввод»"), "")]
    for i, (a, f, u) in enumerate(card):
        r = 4 + i
        ws.Cells(r, 2).Value = a
        F2(ws.Cells(r, 3), "=" + f if not f.startswith("=") else f)
        ws.Cells(r, 4).Value = u
        ws.Cells(r, 3).HorizontalAlignment = -4152
    ws.Range("C4:C30").NumberFormatLocal = "0,00"
    ws.Range("B4:D27").Borders.LineStyle = 1
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 6
    # таблица по всем щитам
    heads = ["Щит", "Питается от", "Руст, кВт", "Кс итог", "Рр, кВт", "cos φ", "Iр, А", "P L1, кВт", "P L2, кВт", "P L3, кВт",
             "ΔU до шин, %", "Iкз min, А", "Iкз max, кА", "Категория"]
    ws.Range(ws.Cells(3, 6), ws.Cells(3, 5 + len(heads))).Value = [heads]
    ws.Range(ws.Cells(3, 6), ws.Cells(3, 5 + len(heads))).Font.Bold = True
    ws.Range(ws.Cells(3, 6), ws.Cells(3, 5 + len(heads))).Interior.Color = C_HEAD
    ws.Range(ws.Cells(3, 6), ws.Cells(3, 5 + len(heads))).WrapText = True
    F2(ws.Range("F4"), '=HSTACK(Щиты[Щит],Щиты[Питается от (итог)],Щиты[Руст, кВт],Щиты[Кс итог],Щиты[Рр, кВт],Щиты[cos φ],Щиты[Iр, А],'
                       'Щиты[P L1, кВт],Щиты[P L2, кВт],Щиты[P L3, кВт],Щиты[ΔU до шин, %],Щиты[Iкз min на шинах, А],Щиты[Iкз max на шинах, кА],Щиты[Категория шины «Ввод»])')
    ws.Range("H4:R300").NumberFormatLocal = "0,00"
    for c in range(6, 6 + len(heads)):
        ws.Columns(c).ColumnWidth = 11
    ws.Range("B29").Value = "Карточка слева — для таблицы под однолинейкой выбранного щита (скопируйте B4:D27 как значения). Справа — все щиты объекта."
    ws.Range("B29").Font.Color = 8421504
    return ws


def build_journal(wb):
    """кабельный журнал в раскладке v3.12: маркировка, трасса, кабель по проекту / проложено, испытания, дата."""
    ws = wb.Worksheets.Add(After=wb.Worksheets("Расчёт_нагрузок"))
    ws.Name = "Кабельный_журнал"
    ws.Range("A1").Value = "КАБЕЛЬНЫЙ ЖУРНАЛ. Для AutoCAD — макрос «Выгрузить журнал» (лёгкий файл для Data Link)."
    ws.Range("A1").Font.Bold = True
    hdr = [("A2:A4", "Маркировка кабеля"), ("B2:C2", "Трасса"), ("B3:B4", "Начало"), ("C3:C4", "Конец"),
           ("D2:H2", "Кабель"), ("D3:H3", "По проекту"), ("I3:K3", "Проложено"), ("I2:K2", "Кабель"),
           ("L2:L4", "Испытания"), ("M2:M4", "Дата прокладки")]
    for a, t in hdr:
        rng = ws.Range(a)
        rng.Merge()
        rng.Value = t
    ws.Range("D4:K4").Value = [["Количество кабелей и сечение", "Лоток", "Труба гофрированная ПВХ", "Труба гофрированная ПНД",
                                "Длина трассы, м", "Марка", "Количество кабелей и сечение", "Длина, м"]]
    h = ws.Range("A2:M4")
    h.Font.Bold = True
    h.WrapText = True
    h.HorizontalAlignment = -4108
    h.VerticalAlignment = -4108
    h.Interior.Color = C_HEAD
    h.Borders.LineStyle = 1
    F2(ws.Range("A5"), ('=LET(m,(Линии[Номер линии]<>"")*(Линии[_кабель]<>""),n,SUM(m),cs,FILTER(Линии[_кабель с сечением],m),'
                        'tr,FILTER(IFERROR(--Линии[Лоток],0),m),pv,FILTER(IFERROR(--Линии[_гофра ПВХ],0),m),pn,FILTER(IFERROR(--Линии[_гофра ПНД],0),m),'
                        'dpv,IFERROR(VLOOKUP(cs,Кабель_диаметр,4,FALSE),""),dpn,IFERROR(VLOOKUP(cs,Кабель_диаметр,5,FALSE),""),'
                        'IF(n=0,"нет линий",HSTACK(FILTER(Линии[Номер линии],m),FILTER(Линии[Щит],m),FILTER(Линии[Наименование (итог)],m),cs,'
                        'IF(tr=0,"","L="&tr),IF(pv=0,"",dpv&" L="&pv),IF(pn=0,"",dpn&" L="&pn),FILTER(Линии[Длина, м],m))))'))
    ws.Range("O2").Value = "Запас длины:"
    ws.Range("P2").Value = 0.2
    ws.Range("P2").Interior.Color = C_IN
    wb.Names.Add("Запас_кабеля", "=Кабельный_журнал!$P$2")
    ws.Range("O4:Q4").Value = [["Кабель (марка, сечение)", "Длина трассы, м", "С запасом, м"]]
    ws.Range("O4:Q4").Font.Bold = True
    ws.Range("O4:Q4").Interior.Color = C_HEAD
    F2(ws.Range("O5"), ('=LET(k,Линии[_кабель с сечением],m,(Линии[Номер линии]<>"")*(Линии[_кабель]<>""),u,UNIQUE(FILTER(k,m,"")),'
                        'l,MAP(u,LAMBDA(x,SUM(m*(k=x)*IFERROR(--Линии[Длина, м],0)))),HSTACK(u,l,ROUNDUP(l*(1+Запас_кабеля),0)))'))
    for c_, w_ in (("A", 16), ("B", 12), ("C", 40), ("D", 26), ("E", 9), ("F", 14), ("G", 14), ("H", 9), ("I", 12), ("J", 14), ("K", 9),
                   ("L", 11), ("M", 11), ("O", 28), ("P", 12), ("Q", 12)):
        ws.Columns(c_).ColumnWidth = w_
    return ws


def build_tags(wb):
    """«Бирки»: строка на каждый провод (участок цепи) — откуда / куда; общие участки (QF → БП) — один раз."""
    try:
        ws = wb.Worksheets("Бирки")
        ws.Cells.Clear()
    except Exception:
        ws = wb.Worksheets.Add(After=wb.Worksheets("Спецификация"))
        ws.Name = "Бирки"
    ws.Range("A1").Value = "БИРКИ: строка на каждый провод — от автомата до устройства, от устройства до клеммы, от клеммы — кабель к потребителю"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:F3").Value = [["№", "Щит", "Откуда", "Куда", "Линия", "Наименование"]]
    ws.Range("A3:F3").Font.Bold = True
    ws.Range("A3:F3").Interior.Color = C_HEAD
    F2(ws.Range("A4"), ('=LET(kk,TOCOL(HSTACK(Линии[_бирка 1],Линии[_бирка 2],Линии[_бирка 3],Линии[_бирка 4],Линии[_бирка 5],Линии[_бирка 6]),1),ii,TOCOL(HSTACK(Линии[_бирка 1 линия],Линии[_бирка 2 линия],Линии[_бирка 3 линия],Линии[_бирка 4 линия],Линии[_бирка 5 линия],Линии[_бирка 6 линия]),1),'
                        'IF(ROWS(kk)=0,"нет",LET(u,UNIQUE(kk),inf,XLOOKUP(u,kk,ii),'
                        'HSTACK(SEQUENCE(ROWS(u)),TEXTBEFORE(u,"¦"),TEXTBEFORE(TEXTAFTER(u,"¦"),"¦"),TEXTAFTER(u,"¦",2),TEXTBEFORE(inf,"¦"),TEXTAFTER(inf,"¦")))))'))
    for c_, w_ in (("A", 5), ("B", 12), ("C", 16), ("D", 46), ("E", 16), ("F", 50)):
        ws.Columns(c_).ColumnWidth = w_
    return ws


def tidy_reference(wb):
    """«Справочная» -> «Таблицы_норм»: старые коды/УГО/этажи/кабели (теперь на листе «Коды») убраны."""
    try:
        ws = wb.Worksheets("Справочная")
    except Exception:
        return
    for a in ("I2:J25", "L2:N60", "P2:Z60", "AD2:AG40", "AU2:AY80", "AZ2:AZ60", "BD2:BD60", "BT2:BV40"):
        try:
            ws.Range(a).ClearContents()
        except Exception:
            for c in ws.Range(a).Cells:
                try:
                    if c.MergeCells:
                        continue
                    c.ClearContents()
                except Exception:
                    pass
    ws.Range("A1").Value = ("ТАБЛИЦЫ НОРМ: допустимые токи кабелей (BX:CA), сопротивления (B:D), номиналы автоматов (F:G), Кс быт (BP:BQ), "
                            "настройки сети (BQ29:BQ31), мест на лист схемы (BF4), УГО → блоки AutoCAD (BA:BC). "
                            "Коды потребителей, УГО, этажи, кабели — на листе «Коды».")
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = 255
    ws.Name = "Таблицы_норм"


def single_line_view(wb):
    """«Однолинейка»: щиты разделены цветом и линией, вводной узел подсвечен, служебные столбцы свёрнуты."""
    ws = wb.Worksheets("Однолинейка")
    ws.Range("F2").Value = "Дифзащита, мА"
    rng = ws.Range("A3:CA3000")
    cf = rng.FormatConditions
    cf.Delete()
    c1 = cf.Add(2, None, '=$CA3="вводной узел"')
    c1.Interior.Color = 16247773
    c2 = cf.Add(2, None, '=НЕЧЁТ($BZ3)')
    c2.Interior.Color = 15921906
    c3 = cf.Add(2, None, '=И($A3<>"";$A3<>$A2)')
    c3.Borders(-4160).LineStyle = 1
    c3.Borders(-4160).Color = 0
    ws.Range("BZ2").Value = "_№ щита"
    ws.Range("CA2").Value = "_зона"
    ws.Range("AF:CA").Columns.OutlineLevel = 2
    try:
        ws.Outline.ShowLevels(0, 1)
    except Exception:
        pass


def spec_by_board(wb, sep):
    """спецификация по щиту: «Спецификация» D1 (все / щит) — на виду; «Каталог» N считает только выбранный щит."""
    sp = wb.Worksheets("Спецификация")
    sp.Range("C1").Value = "Спецификация для щита (все — весь объект):"
    sp.Range("C1").Font.Bold = True
    sp.Range("C1").HorizontalAlignment = -4152
    sp.Range("D1").Value = "все"
    sp.Range("D1").Interior.Color = C_IN
    sp.Range("D1").Font.Bold = True
    v = sp.Range("D1").Validation
    v.Delete()
    v.Add(3, 1, 1, "=Список_щитов")
    v.ShowError = False
    sp.Range("L1").Formula2 = '=IF(D1="все","*",D1)'
    sp.Range("L1").Font.Color = 16777215
    set_names(wb, {"Щит_спец": "=Спецификация!$L$1"})
    try:
        sa = wb.Worksheets("Спецификация_авто")
        sa.Visible = -1
        F2(sa.Range("K1"), '="Щит: "&Спецификация!$D$1&" (выбор — «Спецификация» D1)"')
        sa.Range("K1").Font.Bold = True
    except Exception:
        pass
    kt = wb.Worksheets("Каталог")
    O = "Однолинейка!"
    rng = lambda c: f"{O}${c}$3:${c}$3000"
    f = ('=IF(OR(A4="Автомат",A4="Дифавтомат"),IF(W4="","",COUNTIFS(' + rng("H") + ',U4,' + rng("A") + ',Щит_спец)),'
         'IF(A4="Аппарат",IF(C4="","",COUNTIFS(' + rng("I") + ',C4,' + rng("A") + ',Щит_спец)+COUNTIFS(' + rng("J") + ',C4,' + rng("A") + ',Щит_спец)'
         '+COUNTIFS(' + rng("K") + ',C4,' + rng("A") + ',Щит_спец)),IF(A4="Модуль",IF(C4="","",COUNTIFS(' + rng("AW") + ',"?*",' + rng("AI") + ',C4,'
         + rng("A") + ',Щит_спец)),IF(A4="Клемма",IF(C4="","",IF(C4=$S$7,SUMIFS(' + rng("BW") + ',' + rng("A") + ',Щит_спец),IF(C4=$S$8,SUMIFS('
         + rng("BX") + ',' + rng("A") + ',Щит_спец),""))),""))))')
    kt.Range("N4:N500").Formula2 = f
    kt.Range("N3").Value = "По однолинейке (авто; щит — «Спецификация» D1)"
    # дифзащита в «Каталоге»: «да» = 30 мА, число — свой ток утечки
    kt.Range("U4:U500").Formula2 = '=IF(OR(A4="Автомат",A4="Дифавтомат"),I4&" С"&J4&"А"&IF(K4="да"," 30мА",IF(ISNUMBER(K4)," "&K4&"мА","")),"")'
    kt.Range("K3").Value = "Дифзащита, мА (да = 30)"


# порядок «Щитов»: ввод → главные результаты → детали (свёрнуты)
_MAIN = ["Питается от (итог)", "Замечания щита", "Линия фидера", "Руст, кВт", "Рр, кВт", "Iр, А", "cos φ", "ΔU до шин, %",
         "Iкз min на шинах, А", "Iкз max на шинах, кА", "Категория шины «Ввод»"]
for _c in S:
    if _c["kind"] == "c" and _c["head"] not in _MAIN:
        _c["kind"] = "h"
_S_in = [c for c in S if c["kind"] == "in"]
_S_main = [next(c for c in S if c["head"] == h) for h in _MAIN]
_S_rest = [c for c in S if c["kind"] != "in" and c["head"] not in _MAIN]
for _c in _S_main:
    _c["grp"] = None
_S_main[0]["grp"] = "РЕЗУЛЬТАТ (считается сам)"
S[:] = _S_in + _S_main + _S_rest


# ------------------------------------------------------------------------------------------------
def letters(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


_dup = [h for h, n in __import__("collections").Counter(c["head"] for c in L).items() if n > 1] +        [h for h, n in __import__("collections").Counter(c["head"] for c in S).items() if n > 1]
assert not _dup, "повтор заголовков: %s" % _dup
LETTER = {c["key"]: letters(i + 1) for i, c in enumerate(L)}
HEAD = {c["key"]: c["head"] for c in L}


def esc(h):
    return re.sub(r"([\[\]#'])", r"'\1", h)


def expand(f):
    f = f.replace("{NB}", "OR({_n}3=1,{board}3<>{board}2)")
    f = re.sub(r"\{P:([A-Za-z_][A-Za-z0-9_]*)\}", lambda m: 'IF({_n}3=1,"",{%s}2)' % m.group(1), f)
    f = re.sub(r"\{T:([A-Za-z_][A-Za-z0-9_]*)\}", lambda m: "Линии[%s]" % esc(HEAD[m.group(1)]), f)
    return re.sub(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", lambda m: LETTER[m.group(1)], f)


# ------------------------------------------------------------------------------------------------
def prepare_rows(data):
    """строки таблиц из JSON extract_v3."""
    lines = {(ln["file"], ln.get("board"), ln["line"]): ln for ln in data["lines"]}
    board_file = {}
    for p in data["places"]:
        board_file.setdefault(p["board"], p["file"])
    for b in data["boards"]:
        board_file.setdefault(b["name"], b["file"])
    file_board = {}
    for b, f in board_file.items():
        file_board.setdefault(f, b)
    parent, feeds = {}, {}
    notes = []
    # фидеры из «Нижестоящие_щиты»
    for f in data["feeders"]:
        child = file_board.get(f.get("file"))
        if not child:
            notes.append(f"фидер {f['line']}: нет файла {f.get('file')}")
            continue
        if child in parent:
            notes.append(f"фидер {f['line']}: щит {child} уже питается линией {feeds[child][1]} — второй фидер считается обычной линией")
            continue
        parent[child] = file_board.get(f["from_file"])
        feeds[child] = (f["from_file"], f["line"])
    # линия с номером = имя щита из другого файла
    names = set(board_file)
    for p in data["places"]:
        ln = p.get("line")
        if ln in names and ln not in parent and board_file[ln] != p["file"]:
            parent[ln] = p["board"]
            feeds[ln] = (p["file"], ln)
            notes.append(f"щит {ln} подключён к {p['board']} по номеру линии")
    feed_of = {v: k for k, v in feeds.items()}           # (файл, линия) -> щит
    order = sorted(board_file, key=lambda b: (0 if not parent.get(b) else 1, b))
    if not order:                      # пустой шаблон: один щит без линий
        order = ["ЩР"]
    rows = []
    for b in order:
        pl = [p for p in data["places"] if p["board"] == b]
        pl.sort(key=lambda p: (p.get("pos") is None, p.get("pos") or 0))
        for p in pl:
            r = {"board": b, "pos": "-" if p.get("unplaced") else p.get("pos"), "line": p.get("line")}
            for k_src, k_dst in (("bus", "bus"), ("qf_txt", "qf_txt"), ("rcd", "rcd_in"), ("dev1", "dev1"), ("dev2", "dev2"),
                                 ("dev3", "dev3"), ("module", "module"), ("ch", "ch_in"), ("xt", "xt"), ("mode", "mode")):
                if p.get(k_src) is not None:
                    r[k_dst] = p[k_src]
            ph = p.get("phase")
            if ph is None and p.get("phase_calc"):
                pc = str(p["phase_calc"])
                ph = "1,2,3" if pc == "L1,L2,L3" else pc.replace("L", "")
            if ph is not None:
                r["phase_in"] = ph
            ln = lines.get((p["file"], b, p.get("line"))) or lines.get((p["file"], None, p.get("line")))
            if ln:
                for k_src, k_dst in (("name", "name_m"), ("p", "p_in"), ("u", "u_in"), ("n", "n"), ("ph", "ph_in"), ("l_plan", "lplan_in"),
                                     ("l_tray", "tray"), ("l_pvc", "pvc_in"), ("l_pnd", "pnd_in"), ("cable", "cable_in"), ("cores", "cores_in"),
                                     ("s", "s_in"), ("qf", "qf_in"), ("floor", "floor"), ("dwg_pvc", "dwg_pvc"), ("dwg_pnd", "dwg_pnd"),
                                     ("dwg_handles", "dwg_h"), ("dwg_note", "dwg_note"), ("dwg_label", "dwg_label"), ("kc_group", "grp_in"),
                                     ("kc", "kc_in"), ("section", "sec_in"), ("cat_pz", "catpz_in"), ("k_lay", "klay_in"), ("cat", "cat_in"),
                                     ("fire", "fire_in")):
                    if ln.get(k_src) is not None:
                        r[k_dst] = ln[k_src]
                r["_ref"] = ln.get("_ref")
            child = feed_of.get((p["file"], p.get("line"))) or p.get("feeds")
            if child and child != p.get("line"):          # номер линии = имя щита -> фидер определится сам
                r["feeds"] = child
            rows.append(r)
    boards = []
    bmeta = {b["name"]: b for b in data["boards"]}
    for b in order:
        m = bmeta.get(b, {})
        sch, man = m.get("scheme", {}), m.get("manual", {})
        drawn = [p.get("pos") for p in data["places"] if p["board"] == b and p.get("mode") != "вручную" and isinstance(p.get("pos"), (int, float))]
        nin = m.get("nin", int(min(drawn)) - 1 if drawn else 15)
        boards.append({"Щит": b, "Питается от": None, "Мест под вводной узел": nin, "Вводов": sch.get("inputs"), "АВР": sch.get("avr"),
                       "ИБП от": sch.get("ups_from"),
                       "Питание: от (ТП / сеть)": man.get("from") if not parent.get(b) else None,
                       "ΔU до питающего, %": man.get("du_up"), "Iкз min от сети, А": man.get("ikz_min_up"),
                       "Iкз max от сети, кА": man.get("ikz_max_up"), "Sтр, кВА": man.get("tr_s"), "Схема тр-ра": man.get("tr_scheme"),
                       "Питающий кабель: S, мм²": man.get("s"), "Жил": man.get("cores"), "Кабелей параллельно": man.get("n_par"),
                       "Длина, м": man.get("l"), "Материал": man.get("mat"), "In питающего аппарата, А": man.get("qf_in"),
                       "Хар-ка аппарата": man.get("qf_char"), "_ref": m.get("_ref")})
    return rows, boards, notes


def as_text(v):
    """строка из цифр («00», «0») должна остаться текстом — иначе Excel превратит её в число."""
    if isinstance(v, str) and re.fullmatch(r"\s*[+-]?\d+([.,]\d+)?\s*", v):
        return "'" + v
    return v


def make_table(ws, name, heads, rows_values, first_row=2):
    n = max(1, len(rows_values))
    ncol = len(heads)
    last = letters(ncol)
    ws.Range(f"A{first_row}:{last}{first_row}").Value = [heads]
    if rows_values:
        ws.Range(f"A{first_row + 1}:{last}{first_row + n}").Value = rows_values
    lo = ws.ListObjects.Add(1, ws.Range(f"A{first_row}:{last}{first_row + n}"), None, 1)
    lo.Name = name
    lo.TableStyle = "TableStyleLight1"
    return lo


def set_names(wb, names):
    for nm, ref in names.items():
        try:
            wb.Names(nm).Delete()
        except Exception:
            pass
        wb.Names.Add(nm, ref)


def build(data, base, out, mode="нормы 2026"):
    rows, boards, notes = prepare_rows(data)
    for n_ in notes:
        log("  ", n_)
    shutil.copyfile(base, out)
    with ExcelSession() as xl:
        t0 = time.time()
        wb = xl.Workbooks.Open(os.path.abspath(out), 0, False)
        xl.Calculation = -4135
        xl.ScreenUpdating = False
        for ws in wb.Worksheets:
            try:
                ws.Unprotect(PW)
            except Exception:
                pass
        for nm in DELETE:
            try:
                wb.Worksheets(nm).Delete()
            except Exception:
                pass
        # внешние связи (от «Нижестоящие_щиты» и т.п.) - разорвать
        for src in (wb.LinkSources(1) or []):
            try:
                wb.BreakLink(src, 1)
            except Exception:
                pass
        # имена с #REF! - удалить
        for i in range(wb.Names.Count, 0, -1):
            nmo = wb.Names(i)
            try:
                if "#REF" in nmo.RefersTo:
                    nmo.Delete()
            except Exception:
                pass
        log("основа подготовлена: %.0f с" % (time.time() - t0))

        # --- «Объект»
        wo = wb.Worksheets.Add(Before=wb.Worksheets(1))
        wo.Name = "Объект"
        obj = [("Объект", "", "название и адрес"),
               ("ρ меди для Iкз min, Ом·мм²/м", data["kz"].get("rho_cu") or 0.0263, "≈ 1,5·ρ20 (ГОСТ 28249 п. 2.4.2)"),
               ("ρ алюминия для Iкз min, Ом·мм²/м", data["kz"].get("rho_al") or 0.042, "≈ 1,5·ρ20"),
               ("Переходные сопротивления на щит, Ом", data["kz"].get("r_contact") or 0.015, "контакты и аппараты на каждый щит"),
               ("c_min для Iкз min", data["kz"].get("c_min") or 0.95, "МЭК 60909-0")]
        wo.Range("B1").Value = "ОБЪЕКТ: настройки"
        wo.Range("B1").Font.Bold = True
        wo.Range("B1").Font.Size = 14
        for i, (a, b_, c_) in enumerate(obj):
            r = 3 + i
            wo.Range(f"B{r}").Value = a
            wo.Range(f"C{r}").Value = b_
            wo.Range(f"D{r}").Value = c_
            wo.Range(f"C{r}").Interior.Color = C_IN
        wo.Columns("B").ColumnWidth = 42
        wo.Columns("C").ColumnWidth = 18
        wo.Columns("D").ColumnWidth = 60
        sep = xl.International[4]
        for r, a, b_, c_ in ((8, "Методика расчёта", mode, "v3.12 — как старая книга (для сверки); нормы 2026 — исправления по СП 256 с Изм. 9 (см. «Методика»)"),
                             (9, "kп.к — региональный коэффициент", 1, "СП 256 табл. 7.5а (ЦФО 0,81, СЗФО 0,91…); для частного дома по умолчанию 1")):
            wo.Range(f"B{r}").Value = a
            wo.Range(f"C{r}").Value = b_
            wo.Range(f"D{r}").Value = c_
            wo.Range(f"C{r}").Interior.Color = C_IN
        v = wo.Range("C8").Validation
        v.Delete()
        v.Add(3, 1, 1, sep.join(["v3.12", "нормы 2026"]))
        set_names(wb, {"ρкз_Cu": "=Объект!$C$4", "ρкз_Al": "=Объект!$C$5", "R_пер": "=Объект!$C$6", "c_min": "=Объект!$C$7",
                       "Методика": "=Объект!$C$8", "k_пк": "=Объект!$C$9"})
        wo.Range("B11").Value = "Сеть (Uн), К прокладки, Кс быт, Ко — «Справочная» BP:BQ."
        # табл. 7.5 СП 256 (доля сантехники 100–85 %)
        wo.Range("F2:G2").Value = [["Число силовых ЭП", "Кс (СП 256 табл. 7.5)"]]
        k75 = [(1, 1), (2, 1), (3, 0.9), (5, 0.8), (8, 0.75), (10, 0.7), (15, 0.65), (20, 0.65), (30, 0.6), (50, 0.55), (100, 0.55), (200, 0.5)]
        wo.Range("F3:G%d" % (2 + len(k75))).Value = k75
        wo.Range("F3:G%d" % (2 + len(k75))).Interior.Color = C_IN
        wo.Range("F2:G2").Font.Bold = True
        wo.Columns("F").ColumnWidth = 16
        wo.Columns("G").ColumnWidth = 20
        set_names(wb, {"Кс75_n": "=Объект!$F$3:$F$%d" % (2 + len(k75)), "Кс75_k": "=Объект!$G$3:$G$%d" % (2 + len(k75))})

        # --- «Коды»: справочники объекта (таблицы Excel, растут без ограничения)
        wk = wb.Worksheets.Add(After=wo)
        wk.Name = "Коды"
        wk.Range("A1").Value = "КОДЫ ПОТРЕБИТЕЛЕЙ, УГО, ЭТАЖИ, КАБЕЛИ — новая строка дописывается сразу под таблицей, таблица растёт сама"
        wk.Range("A1").Font.Bold = True
        ref = data.get("ref", {})
        blocks = [("Коды", "A", ["Код", "Наименование", "Опуск, м", "1 - этаж/2 - шлейф", "Пуск: 1 нет/2 двигатель/3 свет", "Марка кабеля",
                                 "Жил", "Напряжение", "Фаз", "Сечение", "Вид нагрузки", "Группа Кс", "Пожарная"], ref.get("codes", {})),
                  ("УГО", "O", ["Код (УГО)", "УГО", "Дифзащита, мА (да = 30)", "Мин. автомат, А", "3L клемм на линию"], ref.get("ugo", {})),
                  ("Этажи", "U", ["Этаж", "Название этажа", "Подъём, м", "По другим этажам, м"], ref.get("floors", {})),
                  ("Кабели", "Z", ["Обозначение", "Кабель", "Сечение в марке (1 да / 2 нет)"], ref.get("cables", {}))]
        for tname, c0, heads, vals in blocks:
            i0 = wk.Range(c0 + "1").Column
            rng_h = wk.Range(wk.Cells(3, i0), wk.Cells(3, i0 + len(heads) - 1))
            rng_h.Value = [heads]
            body = [[as_text(x) for x in list(v) + [None] * (len(heads) - len(v))] for v in vals.values()] or [[None] * len(heads)]
            wk.Range(wk.Cells(4, i0), wk.Cells(3 + len(body), i0 + len(heads) - 1)).Value = body
            lo_ = wk.ListObjects.Add(1, wk.Range(wk.Cells(3, i0), wk.Cells(3 + len(body), i0 + len(heads) - 1)), None, 1)
            lo_.Name = tname
            lo_.TableStyle = "TableStyleLight9"
            lo_.HeaderRowRange.WrapText = True
        wk.Rows(3).RowHeight = 45
        wk.Columns.ColumnWidth = 12
        wk.Columns("B").ColumnWidth = 30
        set_names(wb, {"Потребители_обозначение": "=Коды[[Код]:[Сечение]]", "УГО_спр": "=УГО", "Этаж_обозначение": "=Этажи",
                       "Кабель_обозначение": "=Кабели", "Коды_код": "=Коды[Код]", "Коды_группа": "=Коды[Группа Кс]",
                       "Коды_пожарная": "=Коды[Пожарная]", "Коды_вид": "=Коды[Вид нагрузки]"})
        log("коды: %d, УГО: %d, этажей: %d, кабелей: %d" % tuple(len(b[3]) for b in blocks))

        # --- «Щиты» и «Линии»: шапки + данные
        wsb = wb.Worksheets.Add(After=wk)
        wsb.Name = "Щиты"
        wsl = wb.Worksheets.Add(After=wb.Worksheets("Щиты"))
        wsl.Name = "Линии"
        sheads = [c["head"] for c in S]
        bvals = [[b.get(h) if c["kind"] == "in" else None for c, h in zip(S, sheads)] for b in boards]
        lob = make_table(wsb, "Щиты", sheads, bvals)
        wsh = wb.Worksheets.Add(After=wsb)
        wsh.Name = "Шины"
        wsh.Range("A1").Value = "ШИНЫ: тип шины вручную. Без строки здесь — по имени (ИБП / ДГУ) или по секции линии (Ввод 1 / 2 / 3)."
        wsh.Range("A1").Font.Bold = True
        bus_rows = [[b.get("bus"), b.get("board"), b.get("type")] for b in data.get("buses", [])] or [[None, None, None]]
        lsh = make_table(wsh, "Шины", ["Шина", "Щит", "Тип"], bus_rows)
        vb = lsh.ListColumns("Тип").DataBodyRange.Validation
        vb.Delete()
        vb.Add(3, 1, 1, sep.join(["Ввод 1", "Ввод 2", "Ввод 3", "ИБП", "ДГУ", "Одиночная"]))
        wsh.Columns("A:C").ColumnWidth = 18
        lheads = [c["head"] for c in L]
        lvals = [[r.get(c["key"]) if c["kind"] == "in" else None for c in L] for r in rows]
        lol = make_table(wsl, "Линии", lheads, lvals)
        log("таблицы: щитов %d, строк %d" % (len(boards), len(rows)))
        # формулы
        for c in L:
            if c["f"]:
                try:
                    lol.ListColumns(c["head"]).DataBodyRange.Formula2 = expand(c["f"])
                except Exception as e:
                    log("ОШИБКА формулы Линии[%s]: %s\n   %s" % (c["head"], e, expand(c["f"])[:300]))
        for c in S:
            if c["f"]:
                try:
                    lob.ListColumns(c["head"]).DataBodyRange.Formula2 = c["f"]
                except Exception as e:
                    log("ОШИБКА формулы Щиты[%s]: %s\n   %s" % (c["head"], e, c["f"][:300]))
        set_names(wb, {"Список_щитов": "=Щиты[Щит]"})
        root = next((b["Щит"] for b in boards if not b.get("Питается от")), boards[0]["Щит"])
        build_loads(wb, xl, root, sep)
        build_pz(wb, root, sep)
        build_journal(wb)
        spec_by_board(wb, sep)
        build_totals(wb, root)
        build_tags(wb)
        single_line_view(wb)
        # --- «Проверка»: все замечания объекта
        wp = wb.Worksheets.Add(Before=wb.Worksheets(1))
        wp.Name = "Проверка"
        wp.Range("A1").Value = "ПРОВЕРКА ОБЪЕКТА"
        wp.Range("A1").Font.Bold = True
        wp.Range("A1").Font.Size = 14
        wp.Range("A2").Formula2 = '="Замечаний: щиты — "&COUNTIF(Щиты[Замечания щита],"?*")&", линии — "&COUNTIF(Линии[Замечания],"?*")&". Методика: "&Методика'
        wp.Range("A4:B4").Value = [["Щит", "Замечания щита"]]
        wp.Range("A5").Formula2 = '=FILTER(HSTACK(Щиты[Щит],Щиты[Замечания щита]),Щиты[Замечания щита]<>"","замечаний нет")'
        wp.Range("D4:I4").Value = [["Щит", "№ места", "Номер линии", "Наименование", "Замечания", "Перейти"]]
        wp.Range("D5").Formula2 = ('=FILTER(HSTACK(Линии[Щит],Линии[№ места],Линии[Номер линии],Линии[Наименование],Линии[Замечания],'
                                   'HYPERLINK("#\'Линии\'!C"&(Линии[_n]+2),"→")),Линии[Замечания]<>"","замечаний нет")')
        for a in ("A4:B4", "D4:I4"):
            wp.Range(a).Font.Bold = True
            wp.Range(a).Interior.Color = C_HEAD
        for c_, w_ in (("A", 12), ("B", 70), ("D", 12), ("E", 8), ("F", 16), ("G", 40), ("H", 70), ("I", 8)):
            wp.Columns(c_).ColumnWidth = w_
        wp.Columns("B").WrapText = True
        wp.Columns("H").WrapText = True

        # оформление
        for ws_, spec, lo in ((wsl, L, lol), (wsb, S, lob)):
            for i, c in enumerate(spec):
                hc = lo.HeaderRowRange.Cells(1, i + 1)
                hc.WrapText = True
                if c["kind"] == "in":
                    hc.Interior.Color = C_IN
                else:
                    hc.Interior.Color = C_CALC
                if c.get("note"):
                    try:
                        hc.AddComment(c["note"])
                    except Exception:
                        pass
            ws_.Rows(2).RowHeight = 60
            ws_.Columns.ColumnWidth = 11
        # группы столбцов «Линий» (заголовок группы в строке 1) и скрытие служебных
        last_grp = None
        for i, c in enumerate(L):
            if c.get("grp") and c["grp"] != last_grp:
                cell = wsl.Cells(1, i + 1)
                cell.Value = c["grp"]
                cell.Font.Bold = True
                cell.Interior.Color = C_GRP
                last_grp = c["grp"]
            if c["kind"] == "h":
                wsl.Columns(i + 1).OutlineLevel = 2
        last_grp = None
        for i, c in enumerate(S):
            if c.get("grp") and c["grp"] != last_grp:
                cell = wsb.Cells(1, i + 1)
                cell.Value = c["grp"]
                cell.Font.Bold = True
                cell.Interior.Color = C_GRP
                last_grp = c["grp"]
            if c["kind"] == "h":
                wsb.Columns(i + 1).OutlineLevel = 2
        # формат чисел у результатов: два знака
        for lo, spec in ((lob, S), (lol, L)):
            for c in spec:
                if c["kind"] != "in" and any(u in c["head"] for u in ("кВт", ", А", "%", "Ом", "кА", "квар", "кВА", "cos φ", "Кс", "Ко", "Iкз / In", "Длина")):
                    try:
                        lo.ListColumns(c["head"]).DataBodyRange.NumberFormatLocal = "0,00"
                    except Exception:
                        pass
        # наименование: формула по коду в каждой строке, ручные - поверх
        nm = lol.ListColumns(HEAD["name_m"]).DataBodyRange
        nm.Formula2 = expand('={name_code}3')
        for i, r in enumerate(rows):
            if r.get("name_m"):
                nm.Cells(i + 1, 1).Value = r["name_m"]
        nm.Interior.Color = 15921906
        vm = lol.ListColumns(HEAD["mode"]).DataBodyRange.Validation
        vm.Delete()
        vm.Add(3, 1, 1, sep.join(["вручную", "рисовать"]))
        vm.ShowError = False
        # условное форматирование «Линий» (формулы - в записи русского Excel)
        lastr = R0 + max(1, len(rows)) + 300
        body = wsl.Range(f"A{R0}:{letters(len(L))}{lastr}")
        cf = body.FormatConditions
        cf.Delete()
        Lb, Lz, Ls, Ln = LETTER["bno"], LETTER["zone"], LETTER["sheetno"], LETTER["_n"]
        c1 = cf.Add(2, None, f'=${Lz}{R0}="вводной узел"')
        c1.Interior.Color = 16247773
        c2 = cf.Add(2, None, f'=НЕЧЁТ(${Lb}{R0})')
        c2.Interior.Color = 15921906
        c3 = cf.Add(2, None, f'=И(${Ln}{R0}>1;$A{R0}<>$A{R0 - 1})')
        c3.Borders(-4160).LineStyle = 1
        c3.Borders(-4160).Color = 0
        c4 = cf.Add(2, None, f'=И($A{R0}=$A{R0 - 1};${Ls}{R0}<>${Ls}{R0 - 1};${Ls}{R0}<>"")')
        c4.Borders(-4160).LineStyle = -4115
        c4.Borders(-4160).Color = 8421504
        try:
            wsl.Outline.ShowLevels(0, 1)
            wsb.Outline.ShowLevels(0, 1)
        except Exception:
            pass
        wsl.Activate()
        xl.ActiveWindow.SplitRow = 2
        xl.ActiveWindow.SplitColumn = 3
        xl.ActiveWindow.FreezePanes = True

        # --- «Однолинейка»: интерфейс (спилл столбцов «Линий»)
        w1 = wb.Worksheets("Однолинейка")
        w1.Range("A3:CA1000").ClearContents()
        for c_, k in ONE_MAP.items():
            w1.Range(f"{c_}3").Formula2 = '=LET(v,FILTER(Линии[%s],Линии[_№ места (итог)]<>""),IF(v="","",v))' % esc(HEAD[k])
        w1.Range("AE1").Formula2 = '="Ошибок: "&COUNTIF($BU$3:$BU$3000,2)&"   Предупреждений: "&COUNTIF($BU$3:$BU$3000,1)'

        xl.Calculation = -4105
        t1 = time.time()
        xl.CalculateFull()
        log("пересчёт: %.1f с" % (time.time() - t1))
        # ошибки формул
        for ws_ in (wsl, wsb, w1):
            try:
                e = ws_.UsedRange.SpecialCells(-4123, 16)
                log("ошибки формул на «%s»: %d (%s)" % (ws_.Name, e.Count, e.Address[:200]))
            except Exception:
                log("ошибки формул на «%s»: 0" % ws_.Name)
        res = compare(wsl, lol, wsb, lob, rows, boards)
        # гайд, методика, порядок листов
        import guide_v4
        import guide_v312
        import methodology_sheet
        guide_v312.G = guide_v4.G
        guide_v312.build(wb)
        methodology_sheet.render(wb)
        put_guides(wb)
        tidy_reference(wb)
        try:
            wb.Worksheets("Нагрузки").Name = "Режимы"
        except Exception:
            pass
        order = ["Гайд", "Проверка", "Объект", "Коды", "Щиты", "Итоги_щитов", "Шины", "Линии", "Режимы", "Расчёт_нагрузок", "Кабельный_журнал",
                 "Однолинейка", "Каталог", "Спецификация", "Спецификация_авто", "Бирки", "Методика", "Таблицы_норм", "Справочная_потребители"]
        for nm in reversed(order):
            try:
                wb.Worksheets(nm).Move(Before=wb.Worksheets(1))
            except Exception:
                pass
        for nm in ("Справочная_потребители",):
            try:
                wb.Worksheets(nm).Visible = 0
            except Exception:
                pass
        for nm, color in (("Гайд", 5287936), ("Проверка", 255), ("Объект", 13431551), ("Коды", 15773696), ("Щиты", 13431551),
                          ("Шины", 13431551), ("Линии", 13431551), ("Режимы", 3243501), ("Итоги_щитов", 3243501), ("Расчёт_нагрузок", 3243501),
                          ("Кабельный_журнал", 3243501), ("Однолинейка", 8421504), ("Методика", 5287936)):
            try:
                wb.Worksheets(nm).Tab.Color = color
            except Exception:
                pass
        wb.Worksheets("Гайд").Activate()
        wb.Save()
        log("сохранено: %s (%.0f с)" % (out, time.time() - t0))
        wb.Close(False)
    return res


def compare(wsl, lol, wsb, lob, rows, boards):
    """сверка с результатами v3.12 (rows[i]['_ref'], boards[i]['_ref'])."""
    def colvals(lo, head):
        v = lo.ListColumns(head).DataBodyRange.Value
        return [x[0] for x in v] if isinstance(v, tuple) else [v]
    got = {k: colvals(lol, HEAD[k]) for k in ("p", "i", "qf_line", "s", "l", "du", "du_vru", "grp", "cos")}
    keys = (("p", "p"), ("i", "i"), ("qf", "qf_line"), ("s", "s"), ("l", "l"), ("du", "du"), ("du_vru", "du_vru"), ("grp", "grp"), ("cos", "cos"))
    stats, diffs = {}, []
    for i, r in enumerate(rows):
        ref = r.get("_ref") or {}
        for kr, kg in keys:
            a, b = ref.get(kr), got[kg][i]
            if a in (None, "") and b in (None, ""):
                continue
            ok = (abs(float(a) - float(b)) <= max(0.011, 0.001 * abs(float(a)))) if isinstance(a, (int, float)) and isinstance(b, (int, float)) else (str(a) == str(b))
            st = stats.setdefault(kr, [0, 0])
            st[0 if ok else 1] += 1
            if not ok and len(diffs) < 40:
                diffs.append(f"{r['board']} {r.get('line')}: {kr} v3.12={a} v4={b}")
    log("СВЕРКА линий с v3.12 (совпало / не совпало):", {k: tuple(v) for k, v in stats.items()})
    for d in diffs:
        log("   ", d)
    bgot = {h: colvals(lob, h) for h in ("Руст, кВт", "Руст быт, кВт", "Рр сил, кВт", "Рр, кВт", "Iр, А", "cos φ", "ΔU до шин, %")}
    bkeys = (("p_inst", "Руст, кВт"), ("p_byt", "Руст быт, кВт"), ("p_sil", "Рр сил, кВт"), ("pr", "Рр, кВт"), ("ir", "Iр, А"), ("cos", "cos φ"), ("du_bus", "ΔU до шин, %"))
    for i, b in enumerate(boards):
        ref = b.get("_ref") or {}
        bad = []
        for kr, h in bkeys:
            a, g = ref.get(kr), bgot[h][i]
            if isinstance(a, (int, float)) and isinstance(g, (int, float)):
                if abs(a - g) > max(0.011, 0.001 * abs(a)):
                    bad.append(f"{kr}: {a:.3f} → {g:.3f}")
            elif a not in (None, "") or g not in (None, "", 0):
                if str(a) != str(g):
                    bad.append(f"{kr}: {a} → {g}")
        log("  щит %-12s %s" % (b["Щит"], "OK" if not bad else "; ".join(bad)))
    return stats


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--data", required=True)
    ap.add_argument("--base", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--mode", default="нормы 2026", help="v3.12 или «нормы 2026»")
    a = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    with open(a.data, encoding="utf-8") as fh:
        data = json.load(fh)
    build(data, a.base, a.out, a.mode)


if __name__ == "__main__":
    main()
