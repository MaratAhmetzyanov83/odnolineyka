# -*- coding: utf-8 -*-
"""build_v5.py - книга однолинейки v5 (один файл на объект, .xlsm с макросами). СОБРАН make_build_v5.py — правьте исходники.

    py build_v5.py --data <obj.json> --base "<книга v3.12>" --out "<Однолинейка_v5_объект.xlsm>" [--defaults]

Отличия от v4: в «Линиях» короткая строка (серое — выбор таблицы, синее — своё), маркер группы «гр», автоподбор автомата и
сечения, фазы — автобалансировка по автоматам, кнопки щитов с карточкой, лист «Экспликация», макросы книги.
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
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "2026-10-03_v4", "scripts"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

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
# Таблица «Линии» v5: (ключ, заголовок, вид, формула).
#   in - ввод (виден);  a - ввод с выбором таблицы (виден; формула =Линии[@[_авто …]] серым, своё значение — синим);
#   c  - результат (виден);  x - ввод «Ещё» (свёрнут, кнопка «+»);  xa - ввод «Ещё» с выбором таблицы;  xc - результат в «Ещё»;  h - служебный (скрыт).
# Плейсхолдеры формул: {k}3 - этот ряд, {k}2 - ряд выше, {k}4 - ряд ниже; {P:k} - значение ряда выше (пусто в первом ряду);
# {NB} - первый ряд щита; {T:k} - Линии[заголовок k]. Ссылки на соседние ряды чинит макрос книги после вставки/удаления строк.
# ------------------------------------------------------------------------------------------------
U_F = "(Uн_лин/SQRT(3))"      # кВ
POT = "Потребители_обозначение"
L = []


def col(key, head, kind, f=None, note=None, grp=None, auto=None):
    L.append(dict(key=key, head=head, kind=kind, f=f, note=note, grp=grp, auto=auto))


# «на автомате строки выше» (группа): «гр» в столбце «Группа»; «=», «^», «↑» тоже понимаются.
# Маркер — в отдельном столбце вводa: иначе автомат фидера → ток щита → фазы → снова автомат (цикл).
MARK = 'OR(TRIM({grp_mark}R&"")="гр",TRIM({grp_mark}R&"")="=",TRIM({grp_mark}R&"")="^",TRIM({grp_mark}R&"")="↑")'
MK3, MK4 = MARK.replace("}R", "}3"), MARK.replace("}R", "}4")
NXC = 'AND({board}4={board}3,{mode}4<>"вручную",OR({line}4="",{mk}4=1))'     # ряд ниже — в той же группе
GRPOK = 'AND(NOT({NB}),{mode}3<>"вручную",{P:gsrc}="ok")'                   # ряд в группе автомата выше
# щит с одним вводом или с АВР: каждый ввод несёт весь щит; иначе — только свою шину (секцию)
ONE = 'OR(N(XLOOKUP({board}3,Щиты[Щит],Щиты[Вводов],1))<=1,LOWER(TRIM(XLOOKUP({board}3,Щиты[Щит],Щиты[АВР],"")&""))="да")'

# --- ввод: видимые столбцы, в порядке мысли инженера
col("board", "Щит", "in", note="Имя щита (как на листе «Щиты»). Строки щита — подряд, в порядке мест на схеме. Кнопки щитов над таблицей показывают один щит.")
col("line", "Номер линии", "in", note="Код.Этаж.Помещение.Группа (пример: R.1.05.3). Код — лист «Коды». Номер приходит из чертежа (DLN).")
col("name_in", "Наименование", "a", auto="name_auto",
    note="Серое — по коду и помещению (как собирать — лист «Объект»). Впишите своё — станет синим. Сотрёте — вернётся серое.")
col("n", "Кол-во потребителей, шт", "in", note="Число розеток, светильников… Приходит из DLN. Нужно для мощности (кВт на шт по коду) и для длины опусков.")
col("p_in", "Р, кВт", "a", auto="p_auto",
    note="Серое — кВт на шт по коду × кол-во (розетки 0,06 кВт на шт). Фидер — Руст нижнего щита. Известную мощность впишите сами.")
col("grp_mark", "Группа", "in",
    note="Пусто — у линии свой автомат. «гр» — линия на автомате строки выше (несколько линий на одном автомате). «узо» — линия со своим автоматом на шине УЗО: выше строка «УЗО.1» с «УЗО 2P 40А 30мА» в «Аппаратах». Строки группы и блока УЗО — подряд.")
col("qf_in", "Автомат", "a", auto="qf_auto",
    note="Серое — подбор: полюса, характеристика, номинал по току группы и минимуму кода, ток утечки. В строке «гр» — «↑» и автомат группы. "
         "Свой автомат: 25, C25 или 2P C25А 30мА — таблица проверит.")
col("ctl_in", "Управление", "a", auto="ctl_auto",
    note="Серое — по пометке в выноске (Relay, DIM, DALI) или по коду. Пусто — без модуля. Раскладка по модулям — следующий этап.")
col("dev_in", "Аппараты", "a", auto="dev_auto", note="Аппараты после автомата через «;»: БП, КМ, МК, Т. Пример: КМ NCH8-40/40; БП 24V 75W. Серое — по коду.")
col("s_in", "Сечение, мм²", "a", auto="s_auto",
    note="Серое — подбор: наименьшее сечение не меньше минимума кода, при котором Iдоп·К ≥ In автомата (ПУЭ 1.3.10), ΔU от ВРУ в норме "
         "(СП 256 п. 8.23) и Iкз в конце отключает автомат за 0,4 с (ПУЭ 1.7.79). Своё — синим, таблица проверит.")
col("lplan_in", "Длина линии (по плану), м", "a", auto="lplan_auto",
    note="Серое — длина из чертежа (DLN: ПВХ + ПНД). Своя — синим. Итог с опусками и подъёмом — столбец «Длина, м».")
col("phase_in", "Фаза", "a", auto="phase_auto",
    note="Серое — автобалансировка: каждому автомату (группе) — наименее загруженная фаза. Линии группы — на фазе своего автомата. Своя: 1, 2, 3 или 1,2,3.")
col("xt_in", "Клемма", "a", auto="xt_auto", note="Серое — следующая XT по порядку. Своя — синим. «нет» — клемма без номера.")

# --- результаты (видны)
col("issues", "Замечания", "c",
    '=TEXTJOIN("; ",TRUE,{chk}3,{ae}3,IF(LEFT({icu_chk}3,3)="Icu",{icu_chk}3,""),IF(AND({line}3<>"",{s}3="не подобрано"),"сечение не подобрано: проверьте длину и автомат или впишите сечение",""))',
    note="Что не так в строке. Пусто — в норме. Сводка по объекту — лист «Проверка».")
col("i", "Iр, А", "c",
    '=IF({line}3="","",IF({feeder}3<>"",IFERROR(XLOOKUP({feeder}3,Щиты[Щит],Щиты[Iр, А]),""),IF({inp}3=1,IF(' + ONE + ',IFERROR(XLOOKUP({board}3,Щиты[Щит],Щиты[Iр, А]),""),IFERROR(N({p}3)*XLOOKUP({board}3,Щиты[Щит],Щиты[Кс итог])/(SQRT(3)*Uн_лин*XLOOKUP({board}3,Щиты[Щит],Щиты[cos φ])),"")),'
    'IF({dc}3=1,IFERROR({p}3*1000/{u}3,""),IFERROR(IFS({ph}3=1,{p}3/(' + U_F + '*{cos}3),{ph}3=3,{p}3/(SQRT(3)*Uн_лин*{cos}3)),"")))))',
    note="1ф: P/(Uф·cos φ); 3ф: P/(√3·Uн·cos φ); до 50 В (LED 24 В) — постоянный ток P/U. Фидер — Iр нижнего щита.")
col("l", "Длина, м", "c",
    '=IF({line}3="","",IFERROR(IF(VLOOKUP({code}3,' + POT + ',4,FALSE)=1,VLOOKUP({code}3,' + POT + ',3,FALSE)*(MAX(1,N({n}3))*2-1)+{lplan}3+'
    'VLOOKUP({floorcode}3,Этаж_обозначение,3,FALSE)+VLOOKUP({floorcode}3,Этаж_обозначение,4,FALSE),VLOOKUP({code}3,' + POT + ',3,FALSE)*(MAX(1,N({n}3))*2-1)+{lplan}3),{lplan}3))',
    note="Опуск·(2n−1) + длина по плану + подъём на этаж; n пусто — 1.")
col("du_vru", "ΔU от ВРУ, %", "c", '=IF(ISNUMBER({du}3),{du}3+IF({dc}3=1,0,{du_bus}3),"")',
    note="Потери линии + потери до шин щита. Норма — СП 256 п. 8.23: до светильников 3 %, до прочих 4 % (+0,005 %/м сверх 100 м). Линии до 50 В (LED 24 В) — потери от БП, допустимо — «Объект» (по умолчанию 5 %).")
col("cab_show", "Кабель", "c", '=IF(OR({line}3="",{cab}3=""),"",IF({s}3="не подобрано",{cab}3&" — сечение не подобрано",{cab_sec}3))')
col("place_lbl", "Лист / место", "c",
    '=IF({pos_eff}3="","",IF({zone}3="вводной узел","узел ","")&"л."&(INT(({pos_eff}3-1)/Столбцов_на_лист)+1)&" / "&(MOD({pos_eff}3-1,Столбцов_на_лист)+1))',
    note="Лист схемы A3 и место на нём. «узел» — вводной узел щита (рисуете сами).")

# --- «Ещё»: редкие результаты (xc) и вводы (свёрнуты кнопкой «+»)
col("ut_grp", "Ток утечки, мА", "xc",
    '=IF(OR({start}3=0,{feeder}3<>"",{inp}3=1),"",IF(AND({uzh}3<>"",{uzh}3={_n}3),ROUND(N({uz_ut}3),1),ROUND(N({ut}3)+N({nx_ut}3),1)))',
    note="ПУЭ 7.1.83: 0,4 мА на 1 А тока нагрузки + 10 мкА на 1 м линии; у группы — сумма по линиям. С дифзащитой должно быть не больше 1/3 уставки (30 мА → 10 мА), иначе — замечание.")
col("bp_need", "БП: нужно, Вт", "xc",
    '=IF(OR({start}3=0,N({dc}3)+N({nx_dc}3)=0),"",ROUNDUP(Запас_БП*(N({preq}3)+N({nx_preq}3)),0))',
    note="Строка автомата группы с линиями до 50 В: мощность БП, которая нужна = запас («Объект», 1,3) × Σ (Uвых·I) по линиям группы. "
         "Сравнивается с мощностью из «Аппаратов» (БП 24V 480W).")
col("bp_uo", "Uвых БП, В", "xc",
    '=IF(OR({start}3=0,N({dc}3)+N({nx_dc}3)=0),"",ROUND(MAX(N({uout}3),N({nx_uout}3)),1))',
    note="Напряжение на выходе БП, при котором у самой нагруженной линии в конце остаётся номинал (U + потери). Если больше, чем даёт БП, — увеличьте сечение.")
col("lmax", "Lмакс по КЗ, м", "xc",
    '=IFERROR(IF(OR({line}3="",{inp}3=1,{dc}3=1,NOT(ISNUMBER({s}3)),N({in_grp}3)=0),"",MAX(0,ROUNDDOWN((c_min*(Uн_лин*1000/SQRT(3))/({mult}3*N({in_grp}3))'
    '-XLOOKUP({board}3,Щиты[Щит],Щиты[Zпетли на шинах, Ом]))*{s}3/(2*ρкз_Cu),0))),"")',
    note="Наибольшая длина линии при этом сечении и автомате, когда Iкз в конце ещё отключает автомат за 0,4 с (Iкз ≥ кратность·In, ПУЭ 1.7.79). Сравните с «Длина, м»: видно, какой запас.")
col("feeds", "Питает щит", "x", note="Обычно пусто: если номер линии = имя щита, фидер определяется сам. Пример, когда нужно: линия ЩСПЗ.1 питает щит ЩСПЗ.")
col("rcd_in", "Утечка, мА", "xa", auto="rcd_auto",
    note="Серое — по коду (лист «Коды»). Своё: 10, 30, 100, 300 или «нет». В группе берётся самая строгая уставка. Вид аппарата (дифавтомат / УЗО на группу) — в «Кодах».")
col("ph_in", "Кол-во фаз", "xa", auto="ph_auto", note="1 или 3. Серое — по коду (фидер — 3).")
col("u_in", "Напряжение", "xa", auto="u_auto", note="230 или 400. Серое — по коду.")
col("grp_in", "Группа Кс", "xa", auto="grp_auto",
    note="Быт — табл. 7.2 по сумме; Сил — Кс линии или табл. 7.5; Пост — 1; Отоп — 1; ВНГ — 0,7 (СП 256, Изм. 9); Нет — не учитывается. Серое — по коду.")
col("kc_in", "Кс (для Сил)", "x", note="Кс силовой линии (СП 256 табл. 7.5). Пусто — по таблице 7.5 (режим «нормы 2026») или 1.")
col("cos_in", "cos φ (ручн.)", "x", note="Пусто — 0,95; двигатели — по таблице; силовые в режиме «нормы 2026» — 0,8.")
col("sec_in", "Секция (1/2/3)", "x", note="Секция шин щита. Пусто — 1.")
col("bus", "Шина", "x", note="Пусто — основная шина. Пишется в строке с автоматом. Имя с «ИБП» или «ДГУ» — шина этого источника.")
col("cat_in", "Категория ввода (строки «Ввод»)", "x", note="Только для строк «Ввод»: I особая / I / II / III. Пусто — по схеме щита (вводы, АВР) и шине (ИБП, ДГУ — I особая).")
col("fire_in", "Пожарная (да / нет)", "x", note="Пусто — по коду. Пожарные линии — кабель FRLS, в расчётный максимум не входят.")
col("catpz_in", "Вид нагрузки", "x", note="Строка таблицы нагрузок для пояснительной записки. Пусто — по коду.")
col("mode", "Режим", "x", note="Пусто — как обычно. «вручную» — строка только в таблице, в схеме не рисуется. «рисовать» — нарисовать строку вводного узла.")
col("module", "Модуль", "xa", auto="mod_auto", note="Серое — модуль по умолчанию для управления строки («Объект», таблица «Модули_умолч»): подряд идущие линии с одним управлением идут на каналы одного модуля, новый — когда каналы кончились; «-» — модуль закончился. Своё: модель в первой строке модуля (следующие строки — на его каналы), «-» — конец модуля.")
col("ch_in", "Канал", "x", note="Пусто — авто; № вручную; = параллельно; + в тот же кабель; р — резерв; б/к — без клеммы.")
col("pos", "№ места", "x", note="Пусто — следующее место после вводного узла. Число — своё место (пропуск = пустая колонка). «-» — строки нет на схеме.")
col("brand_in", "Марка кабеля", "xa", auto="brand_auto", note="Обозначение марки (лист «Коды», таблица «Кабели»): мс, мфр… Серое — по коду; пожарная линия — марка для пожарных (лист «Объект»).")
col("cores_in", "Кол-во жил", "xa", auto="cores_auto", note="Серое — по коду, иначе 1ф — 3, 3ф — 5.")
col("klay_in", "К прокл. линии", "x", note="Поправочный коэффициент прокладки. Пусто — общий (лист «Таблицы_норм»).")
col("tray", "Лоток", "x", note="Длина по лотку, м.")
col("pvc_in", "Гофра ПВХ", "x", note="Длина в гофре ПВХ, м. Пусто — из чертежа.")
col("pnd_in", "Гофра ПНД", "x", note="Длина в гофре ПНД, м. Пусто — из чертежа.")
col("floor", "Потолок - 1/Пол - 2", "x")
col("note", "Примечание", "x")
col("dwg_pvc", "DWG: гофра ПВХ по плану, м", "x")
col("dwg_pnd", "DWG: гофра ПНД по плану, м", "x")
col("dwg_h", "DWG: хэндлы полилиний", "x")
col("dwg_note", "DWG: примечание", "x")
col("dwg_label", "Доп информация из выноски", "x")

# --- служебные: разбор номера, наименование, место
col("_n", "_n", "h", "=ROW()-ROW(Линии[#Headers])")
col("sortkey", "_ключ упорядочивания", "h")       # пишет и стирает макрос «Упорядочить щит»
col("code", "Код", "h", '=IF({line}3="","",IFERROR(LEFT({line}3,SEARCH(".",{line}3)-1),{line}3))')
col("rest1", "_часть1", "h", '=IFERROR(RIGHT({line}3,LEN({line}3)-SEARCH(".",{line}3)),"")')
col("floorcode", "Этаж/Шлейф", "h", '=IFERROR(LEFT({rest1}3,SEARCH(".",{rest1}3)-1),{rest1}3)')
col("rest2", "_часть2", "h", '=IFERROR(RIGHT({rest1}3,LEN({rest1}3)-SEARCH(".",{rest1}3)),"")')
col("room", "Помещение", "h", '=IFERROR(LEFT({rest2}3,SEARCH(".",{rest2}3)-1),{rest2}3)')
col("feeder", "_фидер", "h", '=IF({feeds}3<>"",{feeds}3,IF(AND({line}3<>"",{line}3<>{board}3,ISNUMBER(XMATCH({line}3,Щиты[Щит]))),{line}3,""))')
col("name_auto", "_авто наименование", "h",
    '=IF({line}3="","",IF({feeder}3<>"","Щит "&{feeder}3,LET(nm,IFERROR(XLOOKUP({code}3,Коды_код,Коды_имя)&"",""),'
    'md,IFERROR(N(XLOOKUP({code}3,Коды_код,Коды_режим)),1),'
    'ex,IF(Имя_режим="по экспликации",IFERROR(XLOOKUP({floorcode}3&"|"&{room}3,Экспликация[Этаж]&"|"&Экспликация[Помещение],Экспликация[Название])&"",""),""),'
    'rm,IF({room}3="","",IF(ex<>"",". "&ex&IF(Имя_пом="нет"," ("&{room}3&")"," (пом. "&{room}3&")"),IF(Имя_пом="нет"," ",", пом. ")&{room}3)),'
    'IF(nm="","",IF(md=2,nm&" шлейф "&{floorcode}3,nm&rm)))))')
col("name", "Наименование (итог)", "h", '=IF(TRIM({name_in}3&"")="",{name_auto}3,{name_in}3)')
col("bno", "_№ щита", "h", '=IFERROR(XMATCH({board}3,Щиты[Щит]),0)')
col("nin", "_мест узла", "h", '=LET(k,IFERROR(XLOOKUP({board}3,Щиты[Щит],Щиты[Мест под вводной узел]),""),IF(k="",Столбцов_на_лист,N(k)))')
col("pos_eff", "_№ места (итог)", "h",
    '=IF({pos}3="-","",IF(ISNUMBER({pos}3),{pos}3,LET(prev,IF({NB},0,N({P:pos_eff})),nx,prev+1,IF(AND({mode}3<>"вручную",nx<={nin}3),{nin}3+1,nx))))')
col("zone", "_зона", "h", '=IF({pos_eff}3="","",IF({pos_eff}3<={nin}3,"вводной узел","линии"))')

# --- мощность, фазы, марка, cos φ
col("kw_per", "_кВт на шт", "h", '=IFERROR(N(XLOOKUP({code}3,Коды_код,Коды_квт)),0)')
col("inp", "_ввод щита", "h", '=IF(AND({line}3<>"",{code}3="Ввод"),1,0)')     # ввод щита: Р, Iр и cos φ своего щита
col("p_auto", "_авто Р", "h",
    '=IF({line}3="","",IF({feeder}3<>"",IFERROR(XLOOKUP({feeder}3,Щиты[Щит],Щиты[Руст, кВт]),0),IF({inp}3=1,IF(' + ONE + ',IFERROR(XLOOKUP({board}3,Щиты[Щит],Щиты[Руст, кВт]),0),SUMIFS({T:p_inp},{T:board},{board}3,{T:src},{src}3)),'
    'IF({kw_per}3>0,ROUND({kw_per}3*MAX(1,N({n}3)),3),""))))')
col("p_fin", "_Р (ввод итог)", "h", '=IF(TRIM({p_in}3&"")="",{p_auto}3,{p_in}3)')
col("p", "_Р линии, кВт", "h", '=IF({line}3="","",IF({feeder}3<>"",IFERROR(XLOOKUP({feeder}3,Щиты[Щит],Щиты[Руст, кВт]),0),N({p_fin}3)))')
col("u_auto", "_авто U", "h", '=IF({line}3="","",IFERROR(IF(VLOOKUP({code}3,' + POT + ',8,FALSE)=0,"",VLOOKUP({code}3,' + POT + ',8,FALSE)),""))')
col("u", "U, В", "h", '=IF(TRIM({u_in}3&"")="",{u_auto}3,{u_in}3)')
col("dc", "_до 50 В", "h", '=IF({line}3="",0,IF(AND(N({u}3)>0,N({u}3)<=50),1,0))')     # LED 24 В, 12 В: постоянный ток после БП
col("ph_auto", "_авто фаз", "h",
    '=IF({line}3="","",IF({feeder}3<>"",3,IFERROR(IF(N(VLOOKUP({code}3,' + POT + ',9,FALSE))=0,1,VLOOKUP({code}3,' + POT + ',9,FALSE)),1)))')
col("ph", "Фаз", "h", '=IF({line}3="","",IF(TRIM({ph_in}3&"")="",{ph_auto}3,N({ph_in}3)))')
col("brand_auto", "_авто марка", "h",
    '=IF({line}3="","",LET(bc,IFERROR(IF(VLOOKUP({code}3,' + POT + ',6,FALSE)=0,"",VLOOKUP({code}3,' + POT + ',6,FALSE)&""),""),'
    'IF(AND({fire}3="да",OR(bc="",bc="мс")),Марка_пож,IF(AND(bc="",{feeder}3<>""),"мс",bc))))')
col("cable", "Марка (итог)", "h", '=IF(TRIM({brand_in}3&"")="",{brand_auto}3,{brand_in}3)')
col("cores_auto", "_авто жил", "h",
    '=IF({line}3="","",LET(cc,IFERROR(N(VLOOKUP({code}3,' + POT + ',7,FALSE)),0),IF(cc>0,cc,IF({cable}3="","",IF({dc}3=1,2,IF(N({ph}3)=3,5,3))))))')
col("cores", "Жил", "h", '=IF(TRIM({cores_in}3&"")="",{cores_auto}3,{cores_in}3)')
col("s_code", "_мин. сечение кода", "h", '=IFERROR(N(VLOOKUP({code}3,' + POT + ',10,FALSE)),0)')
col("tstart", "_тип пуска", "h", '=IFERROR(VLOOKUP({code}3,' + POT + ',5,FALSE),0)')
col("cos", "cos φ", "h",
    '=IF({line}3="","",IF({feeder}3<>"",IFERROR(XLOOKUP({feeder}3,Щиты[Щит],Щиты[cos φ]),0.95),IF({inp}3=1,IFERROR(XLOOKUP({board}3,Щиты[Щит],Щиты[cos φ]),0.95),IF(ISNUMBER({cos_in}3),{cos_in}3,'
    'IFERROR(IF({tstart}3=2,IFS({ph}3=3,VLOOKUP({p}3,Справочная_двигатели,8,FALSE),{ph}3=1,VLOOKUP({p}3,Справочная_двигатели,10,TRUE)),'
    'IF(AND(Методика="нормы 2026",{grp}3="Сил"),0.8,0.95)),0.95)))))')

# --- группа автомата: «гр» = на автомате строки выше; суммы по группе снизу вверх (ряд ниже)
col("mk", "_гр (на автомате выше)", "h", '=IF(' + MK3 + ',1,0)')
col("start", "_свой автомат", "h", '=IF(AND({line}3<>"",{mk}3=0),1,0)')
col("uz", "_на шине УЗО", "h", '=IF({line}3="",0,IF(TRIM({grp_mark}3&"")="узо",1,0))')     # «узо»: свой автомат на шине УЗО строки выше
col("i_ac", "_ток на 230 В", "h", '=IF({line}3="","",IF({dc}3=1,IFERROR({p}3/(' + U_F + '*MAX(0.5,N({cos}3))),0),N({i}3)))')
col("nx_i", "_ток ниже в группе", "h", '=IF(' + NXC + ',IF({line}4="",0,N({i_ac}4))+N({nx_i}4),0)')
col("nx_p", "_Р ниже в группе", "h", '=IF(' + NXC + ',IF({line}4="",0,N({p}4))+N({nx_p}4),0)')
col("nx_ph", "_фаз ниже в группе", "h", '=IF(' + NXC + ',MAX(IF({line}4="",0,N({ph}4)),N({nx_ph}4)),0)')
col("nx_bb", "_мин. номинал ниже", "h", '=IF(' + NXC + ',MAX(IF({line}4="",0,N({bb}4)),N({nx_bb}4)),0)')
col("nx_ma", "_мА ниже в группе", "h", '=IF(' + NXC + ',LET(a_,IF({line}4="",0,N({ma}4)),b_,N({nx_ma}4),IF(a_=0,b_,IF(b_=0,a_,MIN(a_,b_)))),0)')
col("am", "_дифзащита по коду, мА", "h",
    '=IF({code}3="",0,LET(v,IFERROR(XLOOKUP({line}3,Коды_код,Коды_мА),IFERROR(XLOOKUP({code}3,Коды_код,Коды_мА),IFERROR(XLOOKUP(LEFT({code}3,2),Коды_код,Коды_мА),""))),'
    'IF(v="да",30,IF(ISNUMBER(v),v,0))))')
col("rcd_auto", "_авто утечка", "h", '=IF({line}3="","",IF(N({am}3)>0,{am}3,""))')
col("rcd_kind", "_вид аппарата утечки", "h", '=LET(k,IFERROR(XLOOKUP({code}3,Коды_код,Коды_утечка)&"",""),IF(k="","дифавтомат",k))')
col("rcd_t", "_УЗО строки", "h", '=IF({start}3=0,"",LET(sg,TRIM(TEXTSPLIT({dev_fin}3&"",";")),f_,FILTER(sg,LEFT(sg,3)="УЗО",""),IFERROR(INDEX(f_,1)&"","")))')
col("uzh", "_блок УЗО", "h",
    '=IF({line}3="",IF(OR({NB},{mode}3="вручную"),"",{P:uzh}),IF({start}3=1,IF({uz}3=1,IF({NB},"",{P:uzh}),IF({rcd_t}3<>"",{_n}3,"")),IF(' + GRPOK + ',{P:uzh},"")))')
col("rcd_a", "_УЗО, А", "h",
    '=IF({rcd_t}3="","",LET(tk,TEXTSPLIT(UPPER({rcd_t}3)," ",,TRUE),a_,FILTER(tk,ISNUMBER(--LEFT(tk,1))*((ISNUMBER(SEARCH("А",tk))+ISNUMBER(SEARCH("A",tk)))>0)'
    '*NOT(ISNUMBER(SEARCH("МА",tk)))*NOT(ISNUMBER(SEARCH("MA",tk))),""),IFERROR(--SUBSTITUTE(SUBSTITUTE(INDEX(a_,1),"А",""),"A",""),"")))')
col("rcd_dn", "_УЗО, мА", "h",
    '=IF({rcd_t}3="","",LET(tk,TEXTSPLIT(UPPER({rcd_t}3)," ",,TRUE),a_,FILTER(tk,ISNUMBER(--LEFT(tk,1))*((ISNUMBER(SEARCH("МА",tk))+ISNUMBER(SEARCH("MA",tk)))>0),""),'
    'IFERROR(--SUBSTITUTE(SUBSTITUTE(INDEX(a_,1),"МА",""),"MA",""),30)))')
col("uz_ma", "_мА блока УЗО", "h", '=IF({uzh}3="","",IFERROR(XLOOKUP({uzh}3,{T:_n},{T:rcd_dn}),30))')
col("uz_1f", "_блок УЗО на 1 фазе", "h", '=IF({uzh}3="",0,IF(ISNUMBER(SEARCH("4P",IFERROR(XLOOKUP({uzh}3,{T:_n},{T:rcd_t})&"",""))),0,1))')
col("uz_i", "_ток блока УЗО", "h", '=IF(AND({uzh}3<>"",{uzh}3={_n}3),SUMIFS({T:i_ac},{T:uzh},{_n}3,{T:board},{board}3),"")')
col("uz_ut", "_утечка блока УЗО", "h", '=IF(AND({uzh}3<>"",{uzh}3={_n}3),SUMIFS({T:ut},{T:uzh},{_n}3,{T:board},{board}3),"")')
col("ma", "_мА линии", "h",
    '=IF({line}3="",0,LET(v,IF(TRIM({rcd_in}3&"")="",{rcd_auto}3,{rcd_in}3),nm,IF(v&""="нет",0,IF(v&""="да",30,IFERROR(--v,0))),IF(OR({rcd_kind}3="УЗО на группу",{uzh}3<>""),0,nm)))')
col("grp_ma", "_мА группы", "h", '=IF({start}3=0,"",LET(a_,N({ma}3),b_,IF({mode}3="вручную",0,N({nx_ma}3)),IF(a_=0,b_,IF(b_=0,a_,MIN(a_,b_)))))')
col("bb", "_мин. по типу", "h",
    '=IF({line}3="",0,IFERROR(--XLOOKUP({line}3,Коды_код,Коды_минавт),IFERROR(--XLOOKUP({code}3,Коды_код,Коды_минавт),IFERROR(--XLOOKUP(LEFT({code}3,2),Коды_код,Коды_минавт),0))))')
col("dev_auto", "_авто аппараты", "h", '=IF({line}3="","",IFERROR(XLOOKUP({code}3,Коды_код,Коды_апп)&"",""))')
col("dev_fin", "_аппараты (итог)", "h", '=IF(TRIM({dev_in}3&"")="",{dev_auto}3,{dev_in}3)')
DSPL = '=IF(TRIM({dev_fin}3&"")="","",IFERROR(TRIM(INDEX(TEXTSPLIT({dev_fin}3&";;;",";"),%d)),""))'
col("dev1", "Аппарат 1", "h", DSPL % 1)
col("dev2", "Аппарат 2", "h", DSPL % 2)
col("dev3", "Аппарат 3", "h", DSPL % 3)
col("ba_own", "_БП после автомата", "h",
    '=IF({dev_fin}3="",0,IF(OR(LEFT({dev1}3,2)="T ",LEFT({dev1}3,2)="Т ",ISNUMBER(SEARCH("БП",{dev_fin}3)),ISNUMBER(SEARCH("; T ",{dev_fin}3)),ISNUMBER(SEARCH("; Т ",{dev_fin}3))),1,0))')
col("ctl_auto", "_авто управление", "h",
    '=IF({line}3="","",LET(t,LOWER({dwg_label}3&""),IF(ISNUMBER(SEARCH("dim",t)),"DIM",IF(ISNUMBER(SEARCH("dali",t)),"DALI",'
    'IF(OR(ISNUMBER(SEARCH("relay",t)),ISNUMBER(SEARCH("реле",t))),"Relay",IF(ISNUMBER(SEARCH("кнопк",t)),"",IFERROR(XLOOKUP({code}3,Коды_код,Коды_упр)&"","")))))))')
col("ctl", "_управление (итог)", "h", '=IF(TRIM({ctl_in}3&"")="",{ctl_auto}3,{ctl_in}3)')
col("q_in", "_авто номинал", "h",
    '=IF({line}3="","",LET(ig,IF(N({uz_i}3)>0,N({uz_i}3),N({i_ac}3)+IF({mode}3="вручную",0,N({nx_i}3))),n0,IFERROR(VLOOKUP(ig,Автомат_номиналы,2,TRUE),0),'
    'mn,MAX({bb}3,IF({mode}3="вручную",0,N({nx_bb}3))),MAX(n0,mn,IF({ba_own}3=1,Мин_автомат_БП,0))))')
col("q_ph", "_фаз группы", "h", '=IF({line}3="","",MAX(N({ph}3),IF({mode}3="вручную",0,N({nx_ph}3))))')
col("q_pl", "_полюса", "h", '=IF({line}3="","",IF(N({q_ph}3)>=3,IF(N({grp_ma}3)>0,"4P","3P"),IF(N({grp_ma}3)>0,"2P","1P")))')
col("q_ch", "_характеристика", "h", '=LET(cc,UPPER(TRIM(IFERROR(XLOOKUP({code}3,Коды_код,Коды_хар)&"",""))),IF(OR(cc="B",cc="В"),"B",IF(cc="D","D","С")))')
col("qf_auto", "_авто автомат", "h",
    '=IF({line}3="","",IF({start}3=0,IF(AND(' + GRPOK + ',{P:h_grp}<>""),"↑ "&{P:h_grp},""),'
    'IF(N({q_in}3)=0,"",{q_pl}3&" "&{q_ch}3&{q_in}3&"А"&IF(N({grp_ma}3)>0," "&{grp_ma}3&"мА",""))))')
col("qf_fin", "_автомат (итог ввода)", "h", '=IF(TRIM({qf_in}3&"")="",{qf_auto}3,{qf_in}3)')
col("h", "Автомат (схема)", "h",
    '=IF({start}3=0,"",LET(v,TRIM({qf_fin}3&""),IF(v="","",IF(ISNUMBER(--v),{q_pl}3&" "&{q_ch}3&(--v)&"А"&IF(N({grp_ma}3)>0," "&{grp_ma}3&"мА",""),'
    'IF(ISNUMBER(SEARCH("P",UPPER(v))),v,{q_pl}3&" "&v)))))')
col("au", "_номинал итог", "h",
    '=IF({h}3="","",LET(t,SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(UPPER({h}3),"С","C"),"А","A"),"В","B"),tk,TEXTSPLIT(t," "),'
    'f,FILTER(tk,ISNUMBER(--MID(tk,2,1))*ISNUMBER(SEARCH(LEFT(tk,1)&"#","B#C#D#"))*(LEN(tk)>1),""),x,INDEX(f,1),'
    'IFERROR(--SUBSTITUTE(MID(x,2,9),"A",""),IFERROR(--SUBSTITUTE(INDEX(FILTER(tk,ISNUMBER(--SUBSTITUTE(tk,"A","")),""),1),"A",""),""))))')
col("gsrc", "_группа от", "h", '=IF({start}3=1,IF({mode}3="вручную","узел","ok"),IF(OR({NB},{mode}3="вручную"),"",{P:gsrc}))')
col("in_grp", "_In гр.", "h", '=IF({start}3=1,{au}3,IF(' + GRPOK + ',{P:in_grp},""))')
col("h_grp", "_автомат гр.", "h", '=IF({start}3=1,{h}3,IF(' + GRPOK + ',{P:h_grp},""))')
col("ma_grp_row", "_мА гр. (строки)", "h", '=IF({start}3=1,{grp_ma}3,IF(' + GRPOK + ',{P:ma_grp_row},""))')
col("rcd_grp", "_дифзащита гр. (строки)", "h", '=IF(OR(N({ma_grp_row}3)>0,{uzh}3<>""),1,0)')
col("ba_grp", "_БП в группе", "h", '=IF({start}3=1,{ba_own}3,IF(' + GRPOK + ',N({P:ba_grp}),0))')
col("z", "_ток группы", "h", '=IF({start}3=1,IF(N({uz_i}3)>0,N({uz_i}3),N({i_ac}3)+IF({mode}3="вручную",0,N({nx_i}3))),"")')

# --- кабель: подбор сечения
col("kgr", "_k жил", "h", '=IF(N({cores}3)=1,1,IF({ph}3=1,1,IF(N({cores}3)>=4,0.92,1)))')
col("klay", "_К прокл", "h", '=IF(ISNUMBER({klay_in}3),{klay_in}3,К_прокл)')
col("tcol", "_столбец Iдоп", "h", '=IF(N({cores}3)=1,1,IF({ph}3=1,2,3))')
col("secmark", "_сечение в марке", "h", '=IFERROR(N(VLOOKUP({cable}3,Кабель_обозначение,3,FALSE)),1)')
col("du_bus", "_ΔU до шин", "h", '=IFERROR(N(XLOOKUP({board}3,Щиты[Щит],Щиты[ΔU до шин, %])),0)')
col("zbus", "_Zпетли шин", "h", '=IFERROR(XLOOKUP({board}3,Щиты[Щит],Щиты[Zпетли на шинах, Ом]),"")')
# для подбора сечения — потери и Z до шин по сечениям «по току» (без циклов: щит → фидеры → щит)
col("du_bus_sel", "_ΔU до шин (подбор)", "h", '=IFERROR(N(XLOOKUP({board}3,Щиты[Щит],Щиты[ΔU до шин (подбор), %])),0)')
col("zbus_sel", "_Zпетли шин (подбор)", "h", '=IFERROR(XLOOKUP({board}3,Щиты[Щит],Щиты[Zпетли на шинах (подбор), Ом]),"")')
col("lplan_auto", "_авто длина по плану", "h", '=IF(COUNT({dwg_pvc}3,{dwg_pnd}3)=0,"",ROUNDUP(SUM({dwg_pvc}3,{dwg_pnd}3),0))')
col("lplan", "_Lплан", "h", '=IF(TRIM({lplan_in}3&"")="",{lplan_auto}3,{lplan_in}3)')
col("b", "_b", "h", '=IF({ph}3=1,2,IF({ph}3=3,1,""))')
col("sin", "_sin", "h", '=IFERROR(ROUND(SIN(ACOS({cos}3)),2),"")')
col("du_max", "Допустимо ΔU, %", "h",
    '=IF({line}3="","",IF({dc}3=1,dU_LED,IF(IFERROR(XLOOKUP({code}3,Коды_код,Коды_уго),"")="Светильник",3,4)+IF(Методика="нормы 2026",MIN(0.5,MAX(0,N({l}3)-100)*0.005),0)))')
col("s_auto", "_авто сечение", "h",
    '=IF(OR({line}3="",{cable}3="",{secmark}3<>1),"",LET(sz,Iдоп_S,t_,CHOOSE({tcol}3,Iдоп_1ж,Iдоп_2ж,Iдоп_3ж)*{kgr}3*{klay}3,'
    'dcx,{dc}3=1,inn,IF(dcx,N({i}3),N({in_grp}3)),smin,MAX(1.5,{s_code}3),L_,N({l}3),'
    'rho,XLOOKUP(sz,INDEX(Кабель_сопротивление,0,1),INDEX(Кабель_сопротивление,0,3),""),'
    'xx,XLOOKUP(sz,INDEX(Кабель_сопротивление,0,1),INDEX(Кабель_сопротивление,0,2),""),'
    'du_,IF(dcx,IFERROR(2*rho*L_/sz*N({i}3)/N({u}3)*100,999),IFERROR(N({b}3)*(rho*L_/sz*N({cos}3)+xx*L_*N({sin}3)/1000)*N({i}3)/(' + U_F + '*1000)*100,999)),'
    'kz,IF(OR(dcx,{rcd_grp}3=1,NOT(ISNUMBER({zbus_sel}3)),L_=0,inn=0),1,IFERROR(--(c_min*(Uн_лин*1000/SQRT(3))/({zbus_sel}3+2*ρкз_Cu*L_/sz)>={mult}3*inn),0)),'
    'ok,IFERROR((sz>=smin)*(t_>=inn)*IF(L_>0,(du_+IF(dcx,0,{du_bus_sel}3)<={du_max}3+0.00001),1)*kz,0),'
    'r_,FILTER(sz,ok,"не подобрано"),INDEX(r_,1)))')
col("s", "Сечение", "h", '=IF({line}3="","",IF(TRIM({s_in}3&"")="",{s_auto}3,{s_in}3))')
col("s_basic", "_сечение по току", "h",
    '=IF(OR({line}3="",{cable}3="",{secmark}3<>1),"",LET(sz,Iдоп_S,t_,CHOOSE({tcol}3,Iдоп_1ж,Iдоп_2ж,Iдоп_3ж)*{kgr}3*{klay}3,'
    'ok,IFERROR((sz>=MAX(1.5,{s_code}3))*(t_>=N({in_grp}3)),0),r_,FILTER(sz,ok,""),INDEX(r_,1)))')
col("du_basic", "_ΔU по току, %", "h",
    '=IFERROR(ROUNDUP(N({b}3)*(XLOOKUP({s_basic}3,INDEX(Кабель_сопротивление,0,1),INDEX(Кабель_сопротивление,0,3))*N({l}3)/{s_basic}3*N({cos}3)+'
    'XLOOKUP({s_basic}3,INDEX(Кабель_сопротивление,0,1),INDEX(Кабель_сопротивление,0,2))*N({l}3)*N({sin}3)/1000)*N({i}3)/(' + U_F + '*1000)*100,2),"")')
col("zcab_basic", "_Zкабеля по току", "h", '=IFERROR(2*ρкз_Cu*N({l}3)/{s_basic}3,"")')
col("r", "_ρ", "h", '=IFERROR(VLOOKUP({s}3,Кабель_сопротивление,3,FALSE),"")')
col("x", "_x", "h", '=IFERROR(VLOOKUP({s}3,Кабель_сопротивление,2,FALSE),"")')
col("pvc", "_гофра ПВХ", "h", '=IF({pvc_in}3<>"",{pvc_in}3,IF(OR({lplan}3="",{floor}3=""),"",IF({floor}3=1,MAX(0,{l}3-ROUNDUP(N({dwg_pnd}3),0)-4),ROUNDUP(N({dwg_pvc}3),0))))')
col("pnd", "_гофра ПНД", "h", '=IF({pnd_in}3<>"",{pnd_in}3,IF(OR({lplan}3="",{floor}3=""),"",IF({floor}3=2,MAX(0,{l}3-ROUNDUP(N({dwg_pvc}3),0)-4),ROUNDUP(N({dwg_pnd}3),0))))')
col("du_v", "_ΔU, В", "h", '=IFERROR(IF({dc}3=1,2*{r}3*{l}3/{s}3*{i}3,{b}3*({r}3*{l}3/{s}3*{cos}3+{x}3*{l}3*{sin}3/1000)*{i}3),"")')
col("du", "ΔU, %", "h", '=IFERROR(ROUNDUP({du_v}3/IF({dc}3=1,{u}3,' + U_F + '*1000)*100,2),"")')
col("cab", "_кабель", "h", '=IFERROR(VLOOKUP({cable}3,Кабель_обозначение,2,FALSE),"")')
col("cab_sec", "_кабель с сечением", "h", '=IFERROR(IF({secmark}3=1,IF(ISNUMBER({s}3),{cab}3&" ("&{cores}3&"х"&{s}3&")",{cab}3),{cab}3),"")')
col("cable_full", "_кабель полный", "h",
    '=IF(OR({line}3="",{cab}3="",AND({p_fin}3<>"",NOT(ISNUMBER({p_fin}3)))),"",IF(N({s}3)=0,{cab}3,{cab_sec}3&IF(N({l}3)=0,"",CHAR(10)&"L="&{l}3&"-"&{du}3&"%")))')
col("zcab", "_Zкабеля", "h", '=IFERROR(2*ρкз_Cu*{l}3/{s}3,"")')
col("zmcab", "_Zmax кабеля", "h", '=IFERROR(0.0175*N({l}3)/{s}3,"")')
col("zloop", "Zпетли в конце, Ом", "h",
    '=IFERROR(IF(OR({line}3="",{dc}3=1,NOT(ISNUMBER({s}3)),N({l}3)=0),"",XLOOKUP({board}3,Щиты[Щит],Щиты[Zпетли на шинах, Ом])+{zcab}3),"")')
col("ikz_min", "Iкз min в конце линии, А", "h", '=IF(ISNUMBER({zloop}3),c_min*(Uн_лин*1000/SQRT(3))/{zloop}3,"")')
col("ikz_max", "Iкз max в конце, кА", "h",
    '=IFERROR(IF(OR({line}3="",NOT(ISNUMBER({s}3))),"",1.05*(Uн_лин*1000/SQRT(3))/(XLOOKUP({board}3,Щиты[Щит],Щиты[Zmax на шинах, Ом])+{zmcab}3)/1000),"")')

# --- нагрузка по группам Кс
col("grp_auto", "_авто группа Кс", "h",
    '=IF({line}3="","",IF(ISNUMBER({kc_in}3),"Сил",IF(IFERROR(XLOOKUP({code}3,Коды_код,Коды_группа)&"","")="","Быт",XLOOKUP({code}3,Коды_код,Коды_группа)&"")))')
col("grp", "Группа Кс (итог)", "h", '=IF({line}3="","",IF({feeder}3<>"","Щит",IF(TRIM({grp_in}3&"")="",{grp_auto}3,{grp_in}3)))')
col("kc", "Кс", "h",
    '=IF({grp}3="Щит","щит",IF({grp}3="","",IF({grp}3="Быт",IFERROR(XLOOKUP({board}3,Щиты[Щит],Щиты[Кс быт]),1),'
    'IF({grp}3="Сил",IF(ISNUMBER({kc_in}3),{kc_in}3,IF(Методика="нормы 2026",IFERROR(XLOOKUP({board}3,Щиты[Щит],Щиты[Кс силовых (табл. 7.5)]),1),1)),'
    'IF(OR({grp}3="Пост",{grp}3="Отоп"),1,IF({grp}3="ВНГ",0.7,0))))))')
col("sec", "Секция", "h", '=IF({grp}3="","",IF(AND({inp}3=1,TRIM({sec_in}3&"")=""),IF(COUNTIFS({T:board},{board}3,{T:sec_in},">=2")>0,MIN(3,MAX(1,IFERROR(--{rest1}3,1))),1),IFERROR(MIN(3,MAX(1,INT(--{sec_in}3))),1)))')
col("c_all", "_Руст (без Нет)", "h", '=IF(OR({grp}3="",{grp}3="Нет",{grp}3="Щит"),0,N({p}3))')
col("c_byt", "_Руст быт", "h", '=IF({grp}3="Быт",N({p}3),0)')
col("c_sil", "_Рр сил", "h", '=IF({grp}3="Сил",N({p}3)*N({kc}3),0)')
col("c_post", "_Руст пост", "h", '=IF({grp}3="Пост",N({p}3),0)')
col("c_otop", "_Р отопления", "h", '=IF({grp}3="Отоп",N({p}3),0)')
col("c_vng", "_Рр ВНГ", "h", '=IF({grp}3="ВНГ",N({p}3)*0.7,0)')
col("p_inp", "_Р для вводов", "h", '=IF(OR({line}3="",{grp}3="Нет"),0,N({p}3))')
col("c_pw", "_P для cos", "h", '=IF(OR({line}3="",{grp}3="Щит",{grp}3="Нет"),0,N({p}3))')
col("c_pcos", "_P·cos", "h", '=IF(OR({line}3="",{grp}3="Щит",{grp}3="Нет"),0,N({p}3)*N({cos}3))')

# --- фазы: автобалансировка по автоматам щита (группа — одна фаза); расклад — Щиты[_фазы авто]: крупные первыми
col("gP", "_Р группы", "h",
    '=IF({start}3=1,IF(AND({uz}3=1,{uz_1f}3=1),0,IF(AND({uzh}3<>"",{uzh}3={_n}3,{uz_1f}3=1),SUMIFS({T:p},{T:uzh},{_n}3,{T:board},{board}3),'
    'N({p}3)+IF({mode}3="вручную",0,N({nx_p}3)))),0)')
col("manph", "_фаза вручную", "h",
    '=IF({line}3="","",IF(ISFORMULA({phase_in}3),"",LET(v,UPPER(SUBSTITUTE(TRIM({phase_in}3&"")," ","")),IF(v="","",'
    'IF(OR(v="1,2,3",v="L1,L2,L3",v="123",v="3Ф"),"L1,L2,L3",IF(OR(v="1",v="2",v="3"),"L"&v,IF(OR(v="L1",v="L2",v="L3"),v,"")))))))')
col("phase_auto", "_авто фаза", "h",
    '=IF({line}3="","",IF({start}3=0,{P:ph_grp_up},IF(AND({uz}3=1,{uz_1f}3=1,{uzh}3<>""),IFERROR(XLOOKUP({uzh}3,{T:_n},{T:phase}),"L1"),IF(OR(N({q_ph}3)>=3,N({ph}3)>=3),"L1,L2,L3",'
    'IFERROR(TEXTBEFORE(TEXTAFTER(XLOOKUP({board}3,Щиты[Щит],Щиты[_фазы авто]),";"&{_n}3&":"),";"),"L1")))))')
col("phase", "_фаза итог", "h",
    '=IF({line}3="","",LET(v,UPPER(SUBSTITUTE(TRIM({phase_in}3&"")," ","")),IF(v="",{phase_auto}3,'
    'IF(OR(v="1,2,3",v="L1,L2,L3",v="123",v="3Ф"),"L1,L2,L3",IF(OR(v="1",v="2",v="3"),"L"&v,IF(OR(v="L1",v="L2",v="L3"),v,IF(v="УПР","",v)))))))')
col("ph_grp_up", "_фаза группы", "h", '=IF({start}3=1,{phase}3,IF(' + GRPOK + ',{P:ph_grp_up},""))')
col("ip1", "_Iпуск двиг.", "h", '=IFERROR(IF({tstart}3=2,IFS({ph}3=3,VLOOKUP({p}3,Справочная_двигатели,3,FALSE),{ph}3=1,VLOOKUP({p}3,Справочная_двигатели,5,TRUE)),0),0)')
col("ip2", "_Iпуск LED", "h", '=IFERROR(IF({tstart}3=3,VLOOKUP({p}3,Справочная_ЛЭД,2,TRUE),0),0)')
col("ipusk", "Iпуск, А", "h", '=IF({line}3="","",IF({ip1}3=0,{ip2}3,{ip1}3))')

# --- категории надёжности и шины (порт v3.12 BP:CA)
RANK = '{"III","II","I","I особая"}'
col("fire", "Пожарная (итог)", "h",
    '=IF({line}3="","",IF(TRIM({fire_in}3&"")<>"",IF(LOWER(TRIM({fire_in}3&""))="да","да","нет"),IF(IFERROR(XLOOKUP({code}3,Коды_код,Коды_пожарная),"")="да","да","нет")))')
col("bus_grp", "_шина группы", "h", '=IF({bus}3<>"",{bus}3,IF({start}3=1,"",IF({ag}3="","",{P:bus_grp})))')
col("bus_name", "Шина (итог)", "h", '=IF({line}3="","",IF({bus_grp}3="","основная",{bus_grp}3))')
col("sec_n", "_секция №", "h", '=IF(AND({inp}3=1,TRIM({sec_in}3&"")=""),IF(COUNTIFS({T:board},{board}3,{T:sec_in},">=2")>0,MIN(3,MAX(1,IFERROR(--{rest1}3,1))),1),IFERROR(MIN(3,MAX(1,INT(--{sec_in}3))),1))')   # ввод N — секция N, если в щите есть линии 2-й секции
col("bus_type", "Тип шины", "h",
    '=IF({line}3="","",LET(t,IFERROR(XLOOKUP(1,(Шины[Шина]={bus_name}3)*(((Шины[Щит]={board}3)+(Шины[Щит]=""))>0),Шины[Тип]),"")&"",'
    'IF(t<>"",t,IF(ISNUMBER(SEARCH("ИБП",{bus_name}3)),"ИБП",IF(ISNUMBER(SEARCH("ДГУ",{bus_name}3)),"ДГУ","Ввод "&{sec_n}3)))))')
col("src", "Источник для режимов", "h", '=IF({line}3="","",IF(OR({bus_type}3="ИБП",{bus_type}3="ДГУ"),{bus_type}3,IF(LEFT({bus_type}3,5)="Ввод ",{bus_type}3,"Ввод "&{sec_n}3)))')
col("inp_cat", "Категория ввода (итог)", "h",
    '=IF({inp}3=0,"",IF(TRIM({cat_in}3&"")<>"",TRIM({cat_in}3&""),IF(OR({bus_type}3="ИБП",{bus_type}3="ДГУ"),"I особая",'
    'XLOOKUP({board}3,Щиты[Щит],Щиты[Категория шины «Ввод»],"III"))))')
col("bpath", "_путь щита", "h", '=IF({board}3="","",XLOOKUP({board}3,Щиты[Щит],Щиты[Путь],""))')
col("kind", "Вид нагрузки (итог)", "h",
    '=IF({line}3="","",IF({catpz_in}3<>"",{catpz_in}3,IF(IFERROR(XLOOKUP({code}3,Коды_код,Коды_вид)&"","")="","Прочее",XLOOKUP({code}3,Коды_код,Коды_вид)&"")))')

# --- схема (интерфейс листа «Однолинейка» для SXDRAW)
col("sheetno", "Лист схемы", "h", '=IF({pos_eff}3="","",INT(({pos_eff}3-1)/Столбцов_на_лист)+1)')
col("q", "Кабель (строкой)", "h", '=IF({line}3="","",SUBSTITUTE({cable_full}3,CHAR(10)," "))')
col("t", "_Р схемы", "h", '=IF({line}3="","",IFERROR(--{p}3,""))')
col("uu", "_Ток схемы", "h", '=IF({line}3="","",IFERROR(ROUNDUP({i}3,2),""))')
col("ak", "_фаз", "h", '=IF({line}3="","",IFERROR(--{ph}3,1))')
col("an", "_в кабель выше", "h", '=IF(OR({ch_in}3="+",AND({line}3<>"",{line}3={P:line})),1,0)')
col("bd", "_draw", "h", '=IF(OR({pos_eff}3="",{mode}3="вручную",AND({zone}3="вводной узел",{mode}3<>"рисовать")),0,1)')
col("af", "_гр.", "h",
    '=IF({pos_eff}3="","",IF(AND({start}3=1,{bd}3=1),COUNTIFS({board}$3:{board}3,{board}3,{start}$3:{start}3,1,{bd}$3:{bd}3,1),'
    'IF(OR({NB},{mode}3="вручную",{bd}3=0),"",{P:af_c})))')
col("af_c", "_гр. (перенос)", "h", '=IF({pos_eff}3="",IF({NB},"",{P:af_c}),{af}3)')
col("ab", "_макс. автомат по кабелю", "h",
    '=IF({line}3="","",IFERROR(LOOKUP(INDEX(CHOOSE({tcol}3,Iдоп_1ж,Iдоп_2ж,Iдоп_3ж),MATCH({s}3,Iдоп_S,0))*{kgr}3*{klay}3,'
    '{6,10,16,20,25,32,40,50,63,80,100,125,160,200,250,315,400,500,630}),""))')
col("aj", "_QF нач.", "h", '=IF(AND({start}3=1,{bd}3=1,{af}3<>""),IF(N({grp_ma}3)>0,"QFD","QF")&{af}3,"")')
col("ag", "_QF гр.", "h", '=IF({aj}3<>"",{aj}3,IF(OR({NB},{mode}3="вручную"),"",{P:ag}))')
col("az", "_аппараты гр.", "h", '=IF({aj}3<>"",TEXTJOIN(" → ",TRUE,{dev1}3,{dev2}3,{dev3}3),IF({ag}3="","",{P:az}))')
col("ba", "_БП/тр-р", "h",
    '=IF({az}3="",0,IF(OR(LEFT({az}3,2)="T ",LEFT({az}3,2)="Т ",ISNUMBER(SEARCH("→ T ",{az}3)),ISNUMBER(SEARCH("→ Т ",{az}3)),'
    'ISNUMBER(SEARCH("БП",{az}3))),1,0))')
col("mdef", "_модуль по умолч.", "h", '=IF(OR({line}3="",{ctl}3=""),"",IFERROR(XLOOKUP({ctl}3,Модули_умолч[Управление],Модули_умолч[Модель])&"",""))')
col("mstep", "_каналов на линию", "h", '=IF({ctl}3="",1,IFERROR(MAX(0,N(XLOOKUP({ctl}3,Модули_умолч[Управление],Модули_умолч[Каналов на линию]))),1))')
col("man_mod", "_модуль вручную", "h", '=IF(ISFORMULA({module}3),"",TRIM({module}3&""))')
col("mst", "_блок модуля", "h",
    '=IF({man_mod}3="-","",IF({man_mod}3<>"","M",IF(AND(NOT({NB}),{P:mst}="M"),"M",IF({mdef}3<>"","A",'
    'IF(AND(NOT({NB}),{P:mst}="A",{line}3="",OR({ch_in}3="р",{ch_in}3="+",{ch_in}3="=")),"A","")))))')
col("mrun", "_модель авто-блока", "h", '=IF({mst}3="A",IF({mdef}3<>"",{mdef}3,{P:mrun}),"")')
col("mch", "_каналов модели", "h", '=IF({mrun}3="",0,IFERROR(N(VLOOKUP({mrun}3,Модули_спр,3,FALSE)),0))')
col("mod_auto", "_авто модуль", "h",
    '=IF({mst}3="A",IF(OR({NB},{P:mst}<>"A",{P:mrun}<>{mrun}3,AND({mch}3>0,{mstep}3>0,N({P:ax})+{mstep}3>{mch}3)),{mrun}3,""),'
    'IF(AND({mst}3="",NOT({NB}),{P:mst}="A",{man_mod}3=""),"-",""))')
col("mod", "_модуль (итог)", "h", '=IF({man_mod}3<>"",{man_mod}3,{mod_auto}3)')
col("av", "_префикс", "h", '=IF(OR({mod}3="",{mod}3="-"),"",IFERROR(VLOOKUP({mod}3,Модули_спр,2,FALSE),"R"))')
col("aw", "_модуль нач.", "h", '=IF({av}3="","",{av}3&"."&COUNTIFS({board}$3:{board}3,{board}3,{av}$3:{av}3,{av}3))')
col("ah", "_модуль гр.", "h", '=IF({mod}3="-","",IF({aw}3<>"",{aw}3,IF({NB},"",{P:ah})))')
col("ai", "_модель гр.", "h", '=IF({ah}3="","",IF(AND({mod}3<>"",{mod}3<>"-"),{mod}3,{P:ai}))')
col("o", "Канал (итог)", "h",
    '=IF({ch_in}3="=",{P:ax},IF(ISNUMBER({ch_in}3),{ch_in}3,IF({ah}3="","",IF(OR({line}3<>"",{ch_in}3="+",{ch_in}3="р"),'
    'IF(OR({aw}3<>"",{P:ax}=""),1,{P:ax}+IF({mstep}3=0,0,1)),""))))')
col("ax", "_посл. канал", "h", '=IF({o}3<>"",MAX({o}3,{o}3+{mstep}3-1),IF(OR({NB},{mod}3="-",{aw}3<>""),"",{P:ax}))')
col("ay", "_клемма", "h", '=IF(OR({ch_in}3="б/к",{mode}3="вручную"),0,IF(OR({cable_full}3<>"",{ch_in}3="+",{ch_in}3="="),1,0))')
col("xt_auto", "_авто клемма", "h", '=IF({ay}3=1,"XT"&COUNTIFS({board}$3:{board}3,{board}3,{ay}$3:{ay}3,1),"")')
col("xt_eff", "Клемма (итог)", "h", '=LET(v,TRIM({xt_in}3&""),IF(v="нет","",IF(v="",{xt_auto}3,v)))')
col("g", "QF", "h", '=IF(OR({line}3<>"",{o}3<>"",{ay}3=1),{ag}3,"")')
col("m", "_Модуль (схема)", "h", '=IF(OR({o}3<>"",AND({mod}3<>"",{mod}3<>"-")),{ah}3,"")')
col("w", "Проверка автомата", "h",
    '=IF({line}3="","",IF(AND({start}3=1,{h}3<>"",{au}3=""),"номинал автомата не распознан",IF({in_grp}3="",IF({mk}3=1,"«гр»: нет автомата выше",""),'
    'IF(AND({start}3=1,N({z}3)>N({in_grp}3)),"ПЕРЕГРУЗ",IF(AND(N({ba_grp}3)=0,ISNUMBER({ab}3),N({ab}3)<N({in_grp}3)),"КАБЕЛЬ НЕ ЗАЩИЩЁН","ok")))))')
col("x_", "Проверка канала", "h",
    '=IF(OR({o}3="",{ah}3=""),"",IF({o}3+MAX(0,{mstep}3-1)>IFERROR(VLOOKUP({ai}3,Модули_спр,3,FALSE),99),"ЛИШНИЙ КАНАЛ",'
    'IF(AND({ch_in}3<>"=",{mstep}3>0,COUNTIFS({T:ah},{ah}3,{T:o},{o}3,{T:ch_in},"<>=",{T:board},{board}3)>1),"ДУБЛЬ КАНАЛА","")))')
col("ikz_ratio", "Iкз / In", "h", '=IFERROR(IF(ISNUMBER({ikz_min}3),{ikz_min}3/N({in_grp}3),""),"")')
col("mult", "_кратность", "h", '=IF(OR(ISNUMBER(SEARCH(" B"," "&{h_grp}3)),ISNUMBER(SEARCH(" В"," "&{h_grp}3))),5,IF(ISNUMBER(SEARCH(" D"," "&{h_grp}3)),20,10))')
col("trip", "Отключение ≤ 0,4 с", "h",
    '=IF(NOT(ISNUMBER({ikz_ratio}3)),"",IF({ikz_ratio}3>={mult}3,"ok",IF(N({rcd_grp}3)=1,"ok (дифзащита "&IF({uzh}3<>"",{uz_ma}3,{ma_grp_row}3)&" мА)","мало: Iкз < "&{mult}3&"·In")))')
col("ut", "_ток утечки линии, мА", "h", '=IF(OR({line}3="",{feeder}3<>"",{inp}3=1),0,0.4*N({i_ac}3)+IF({dc}3=1,0,0.01*N({l}3)))')
col("nx_ut", "_утечка ниже в группе", "h", '=IF(' + NXC + ',IF({line}4="",0,N({ut}4))+N({nx_ut}4),0)')
col("uout", "_Uвых БП для линии, В", "h", '=IF({dc}3=1,N({u}3)+N({du_v}3),0)')
col("preq", "_мощность от БП, Вт", "h", '=IF({dc}3=1,(N({u}3)+N({du_v}3))*N({i}3),0)')
col("nx_preq", "_мощность от БП ниже", "h", '=IF(' + NXC + ',IF({line}4="",0,N({preq}4))+N({nx_preq}4),0)')
col("nx_uout", "_Uвых ниже", "h", '=IF(' + NXC + ',MAX(IF({line}4="",0,N({uout}4)),N({nx_uout}4)),0)')
col("nx_dc", "_линий до 50 В ниже", "h", '=IF(' + NXC + ',IF({line}4="",0,N({dc}4))+N({nx_dc}4),0)')
col("bp_t", "_БП группы", "h", '=IF({start}3=0,"",IFERROR(TRIM(MID({dev_fin}3,SEARCH("БП",{dev_fin}3),40)),""))')
col("bp_w", "_БП, Вт", "h",
    '=IF({bp_t}3="","",LET(tk,TEXTSPLIT(SUBSTITUTE(SUBSTITUTE({bp_t}3,";"," "),"→"," ")," ",,TRUE),'
    'w_,FILTER(tk,ISNUMBER(--LEFT(tk,1))*((ISNUMBER(SEARCH("W",tk))+ISNUMBER(SEARCH("Вт",tk)))>0),""),'
    'IFERROR(--SUBSTITUTE(SUBSTITUTE(UPPER(INDEX(w_,1)),"W",""),"ВТ",""),"")))')
col("bp_v", "_БП, В", "h",
    '=IF({bp_t}3="","",LET(tk,TEXTSPLIT(SUBSTITUTE(SUBSTITUTE({bp_t}3,";"," "),"→"," ")," ",,TRUE),'
    'v_,FILTER(tk,ISNUMBER(--LEFT(tk,1))*((ISNUMBER(SEARCH("V",tk))+(RIGHT(UPPER(tk))="В"))>0),""),'
    'IFERROR(--SUBSTITUTE(SUBSTITUTE(UPPER(INDEX(v_,1)),"V",""),"В",""),"")))')
col("ch_cap", "_ток канала модуля, А", "h",
    '=IF(OR({dc}3=0,{ai}3="",{o}3=""),"",IFERROR(N(XLOOKUP({ai}3,Каталог!$C$4:$C$500,Каталог!$Q$4:$Q$500)),""))')
col("ch_sum", "_ток канала, А", "h", '=IF(N({ch_cap}3)=0,"",SUMIFS({T:i},{T:board},{board}3,{T:ah},{ah}3,{T:o},{o}3))')
col("ae", "_статус", "h",
    '=IF(AND({line}3="",{g}3=""),"",TEXTJOIN("; ",TRUE,IF(AND({line}3<>"",{phase}3=""),"нет фазы",""),'
    'IF(AND({line}3<>"",N({t}3)=0,{grp}3<>"Нет"),"нет мощности",""),IF(AND({xt_eff}3<>"",COUNTIFS({T:board},{board}3,{T:xt_eff},{xt_eff}3)>1),"дубль клеммы "&{xt_eff}3,""),'
    'IF(AND({w}3<>"",{w}3<>"ok"),{w}3,""),IF({x_}3<>"",{x_}3,""),IF(AND(N({cores}3)>=4,N({ph}3)=1),{cores}3&" жил на 1 фазе",""),'
    'IF(AND({code}3="R",{start}3=1,IF({uzh}3<>"",N({uz_ma}3)>30,OR(N({grp_ma}3)=0,N({grp_ma}3)>30))),"розетка без дифзащиты до 30 мА",""),'
'IF(AND({start}3=1,N({grp_ma}3)>0,ISNUMBER({ut_grp}3)),IF({ut_grp}3>N({grp_ma}3)/3,"ток утечки "&FIXED({ut_grp}3,1)&" мА > 1/3 уставки ("&{grp_ma}3&" мА)",""),""),'
    'IF(AND({uz}3=1,{uzh}3=""),"«узо»: нет строки с УЗО выше (УЗО 2P 40А 30мА в «Аппаратах»)",""),'
    'IF(AND({uzh}3<>"",{uzh}3={_n}3),IF({rcd_a}3="","впишите номинал УЗО (УЗО 2P 40А 30мА)",IF(N({rcd_a}3)<N({in_grp}3),"УЗО "&{rcd_a}3&" А < автомата "&{in_grp}3&" А перед ним",'
    'IF(N({rcd_a}3)<N({uz_i}3),"УЗО "&{rcd_a}3&" А < тока блока "&FIXED({uz_i}3,1)&" А",""))),""),'
    'IF(AND({uzh}3<>"",{uzh}3={_n}3,ISNUMBER({uz_ut}3)),IF({uz_ut}3>N({uz_ma}3)/3,"ток утечки блока УЗО "&FIXED({uz_ut}3,1)&" мА > 1/3 уставки ("&{uz_ma}3&" мА)",""),""),'
    'IF(AND({start}3=1,N({dc}3)+N({nx_dc}3)>0,{bp_t}3=""),"линии до 50 В без БП: впишите в «Аппараты» (БП 24V 480W)",""),'
    'IF(AND({start}3=1,ISNUMBER({bp_w}3),ISNUMBER({bp_need}3)),IF({bp_need}3>{bp_w}3,"БП "&{bp_w}3&" Вт < нужно "&{bp_need}3&" Вт (×"&Запас_БП&")",""),""),'
    'IF(AND({start}3=1,ISNUMBER({bp_v}3),N({dc}3)=1),IF({bp_v}3<>N({u}3),"БП на "&{bp_v}3&" В, линия "&{u}3&" В",""),""),'
    'IF(AND(N({ch_cap}3)>0,N({ch_sum}3)>N({ch_cap}3)),{ah}3&" канал "&{o}3&": "&FIXED({ch_sum}3,1)&" А > "&{ch_cap}3&" А",""),'
    'IFERROR(IF(AND(ISNUMBER({du_vru}3),{du_vru}3>{du_max}3),"потери "&FIXED({du_vru}3,1)&" % > "&{du_max}3&" %",""),""),'
    'IFERROR(IF(LEFT({trip}3,4)="мало","Iкз мало — отключение > 0,4 с",""),""),IF(AND({line}3<>"",{pos_eff}3=""),"не на схеме",""),'
    'IF(AND({start}3=0,{line}3<>"",{phase}3<>"",{P:ph_grp_up}<>"",{phase}3<>{P:ph_grp_up}),"фаза не как у автомата группы ("&{P:ph_grp_up}&")","")))')
col("icu", "Icu автомата, кА", "h",
    '=IF(OR({start}3=0,{h}3=""),"",LET(t,IFERROR(XLOOKUP({h}3,Каталог!$U$4:$U$500,Каталог!$G$4:$G$500),""),k,SEARCH("кА",t),'
    'IFERROR(--TRIM(RIGHT(SUBSTITUTE(MID(t,MAX(1,k-4),k-MAX(1,k-4))," ",REPT(" ",9)),5)),"")))')
col("icu_chk", "Проверка Icu", "h",
    '=IF(OR({icu}3="",{start}3=0),"",IF(XLOOKUP({board}3,Щиты[Щит],Щиты[Iкз max на шинах, кА],"")="","",'
    'IF({icu}3<XLOOKUP({board}3,Щиты[Щит],Щиты[Iкз max на шинах, кА]),"Icu "&{icu}3&" кА < Iкз max","ok")))')
col("chk", "_целость", "h",
    '=IF(OR(AND({_n}3>1,N({_n}2)<>{_n}3-1),AND({_n}4<>"",N({_n}4)<>{_n}3+1)),"строки сдвинуты (вставка или удаление) — нажмите «Починить формулы»","")')
col("chain", "_цепь для бирок", "h", '=IF(OR({mode}3="вручную",{pos_eff}3="",AND({line}3="",{xt_eff}3="")),"",TEXTJOIN("|",TRUE,{ag}3,SUBSTITUTE({az}3," → ","|"),IF({ah}3<>"",{ah}3&IF({o}3<>""," к."&{o}3,""),""),{xt_eff}3))')
col("chain2", "_цепь с кабелем", "h", '=IF({chain}3="","",{chain}3&IF({line}3<>"","|кабель "&{cab_sec}3&" → "&{name}3,""))')
for _k in range(1, 7):
    col("tg%d" % _k, "_бирка %d" % _k, "h",
        '=IF({chain2}3="","",LET(nd,TEXTSPLIT({chain2}3,"|"),IF(COLUMNS(nd)>%d,{board}3&"¦"&INDEX(nd,%d)&"¦"&INDEX(nd,%d),"")))' % (_k, _k, _k + 1))
    col("ti%d" % _k, "_бирка %d линия" % _k, "h", '=IF({tg%d}3="","",{line}3&"¦"&{name}3)' % _k)
col("ac", "_модель (схема)", "h", '=IF({m}3="","",{ai}3)')
col("ad", "_цепь", "h",
    '=IF(AND({line}3="",{o}3="",{xt_eff}3=""),"",TEXTJOIN(" → ",TRUE,{ag}3,{az}3,IF({ah}3<>"",{ah}3&IF({o}3<>""," к."&{o}3,""),'
    'IF({o}3<>"","к."&{o}3,"")),{xt_eff}3,IF({an}3=1,"(в кабель выше)","")))')
col("be", "_bus2", "h", '=IF({bus}3="",0,1)')
col("bf", "_rcd", "h", '=IF({aj}3="",0,IF(LEFT({aj}3,3)="QFD",1,0))')
col("bg", "_chtype", "h", '=IF({ch_in}3="=","=",IF({ch_in}3="+","+",IF({ch_in}3="р","R","")))')
DEVT = ('=IF({d}3="","",IF(OR(LEFT({d}3,1)="T",LEFT({d}3,1)="Т"),"T",IF(LEFT(UPPER({d}3),2)="БП","PS",IF(OR(LEFT(UPPER({d}3),2)="МК",'
        'LEFT(UPPER({d}3),2)="MK"),"MK",IF(OR(LEFT(UPPER({d}3),2)="КМ",LEFT(UPPER({d}3),2)="KM"),"KM","X")))))')
col("bh", "_dev1", "h", DEVT.replace("{d}", "{dev1}"))
col("bi", "_dev2", "h", DEVT.replace("{d}", "{dev2}"))
col("bj", "_dev3", "h", DEVT.replace("{d}", "{dev3}"))
col("by", "_УГО (рус)", "h",
    '=IF({line}3="","",IFERROR(XLOOKUP({line}3,Коды_код,Коды_уго),IFERROR(XLOOKUP({code}3,Коды_код,Коды_уго),IFERROR(XLOOKUP(LEFT({code}3,2),Коды_код,Коды_уго),""))))')
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
    '=IF({ay}3<>1,0,LET(k,MAX(IFERROR(--XLOOKUP({line}3,Коды_код,Коды_3L),0),IFERROR(--XLOOKUP({code}3,Коды_код,Коды_3L),0),'
    'IFERROR(--XLOOKUP(LEFT({code}3,2),Коды_код,Коды_3L),0)),IF(k>0,k,IF(IFERROR(--LEFT(SUBSTITUTE({bq}3,"x","х"),FIND("х",SUBSTITUTE({bq}3,"x","х"))-1),3)>3,Каталог!$S$9,1))))')
col("bx", "_доп. проходных", "h", '=IF({ay}3<>1,0,MAX(0,IFERROR(--LEFT(SUBSTITUTE({bq}3,"x","х"),FIND("х",SUBSTITUTE({bq}3,"x","х"))-1),3)-3*{bw}3))')
col("e_draw", "_автомат для схемы", "h", '=IF({start}3=1,"авто","")')
col("f_rcd", "_мА для схемы", "h", '=IF({start}3=1,IF(N({grp_ma}3)>0,{grp_ma}3,""),"")')
col("as_i", "_дифзащита гр. (0/1)", "h", '=IF({start}3=1,IF(N({grp_ma}3)>0,1,0),"")')
col("vis", "_видна", "h", '=SUBTOTAL(103,{board}3)')

# Однолинейка v3.12: столбец -> ключ «Линий» (SXDRAW читает A, B, D, H, I-K, O-V, AE, AG-AJ, AN, AY, BD-BL, BY)
ONE_MAP = {"A": "board", "B": "pos_eff", "C": "sheetno", "D": "bus", "E": "e_draw", "F": "f_rcd", "G": "g", "H": "h", "I": "dev1",
           "J": "dev2", "K": "dev3", "L": "module", "M": "m", "N": "ch_in", "O": "o", "P": "xt_eff", "Q": "q", "R": "line", "S": "name",
           "T": "t", "U": "uu", "V": "phase", "W": "w", "X": "x_", "Y": "mode", "Z": "z", "AA": "q_in", "AB": "ab", "AC": "ac",
           "AD": "ad", "AE": "ae", "AF": "af", "AG": "ag", "AH": "ah", "AI": "ai", "AJ": "aj", "AK": "ak", "AL": "code", "AM": "am",
           "AN": "an", "AO": "q_in", "AP": "ab", "AQ": "au", "AR": "q_ph", "AS": "as_i", "AT": "q_pl", "AU": "au", "AV": "av", "AW": "aw",
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
scol("Питается от", "a", '=[@[_авто питается от]]', note="Серое — по линии-фидеру в «Линиях» (номер линии = имя щита или столбец «Питает щит»). Своё — впишите, если фидера в книге нет. У ВРУ / ГРЩ — пусто.")
scol("Вводов", "in", note="Число вводов 1 / 2 / 3.")
scol("АВР", "in", note="да / нет")
scol("ИБП от", "in", note="Ввод 1 / 2 / 3 — от какого ввода питается ИБП (ЩГП).")
scol("Мест под вводной узел", "in", note="Сколько первых мест схемы (шаг 20 мм) занимает вводной узел — вводы, АВР, ИБП, байпас, учёт; его рисуете сами. Пусто — 17 (весь первый лист A3); 11 — узел на части листа, линии на том же листе; 34 — два листа. Строки узла на схеме не рисуются (режим «рисовать» — нарисовать).")
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
scol("Питается от (итог)", "h", '=IF(TRIM([@[Питается от]]&"")="",[@[_авто питается от]],[@[Питается от]])',
     note="Вышестоящий щит: вручную или по линии-фидеру.")
scol("Уровень", "c", '=LEN([@Путь])-LEN(SUBSTITUTE([@Путь],"|",""))-2')
scol("_строка фидера", "h", '=IFERROR(XMATCH([@Щит],Линии[_фидер]),"")')
scol("_авто питается от", "h", '=IFERROR(INDEX(Линии[Щит],[@[_строка фидера]]),"")')
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
         f'=SUMIFS(Линии[_Р линии, кВт],Линии[Щит],[@Щит],Линии[_фаза итог],"L1,L2,L3",Линии[Группа Кс (итог)],"<>Нет")/3+'
         f'SUMIFS(Линии[_Р линии, кВт],Линии[Щит],[@Щит],Линии[_фаза итог],"L{k}",Линии[Группа Кс (итог)],"<>Нет")')
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
# для подбора сечений в «Линиях»: те же ΔU и Z до шин, но по сечениям фидеров «по току» (иначе цикл щит → фидер → щит)
scol("_ΔU фидера (подбор)", "h", '=IF([@[_строка фидера]]="","",INDEX(Линии[_ΔU по току, %],[@[_строка фидера]]))')
scol("ΔU до шин (подбор), %", "h", f'=SUMPRODUCT({ONPATH},IFERROR(--[_ΔU фидера (подбор)],0)+IFERROR(--[_ΔU источника],0))')
scol("_Zфидера (подбор)", "h", '=IF([@[_строка фидера]]="","",INDEX(Линии[_Zкабеля по току],[@[_строка фидера]]))')
scol("Zпетли на шинах (подбор), Ом", "h",
     f'=IF(XLOOKUP([@Корень],[Щит],[_Zисточника],"")="","",SUMPRODUCT({ONPATH},IFERROR(--[_Zисточника],0)+IFERROR(--[_Zфидера (подбор)],0))+SUMPRODUCT({ONPATH})*R_пер)')
# категории по пути от ГРЩ (наследование: один ввод — не выше II)
scol("Категория шины «Ввод»", "c",
     '=LET(w,TEXTSPLIT([@Путь],"|"),names,FILTER(w,w<>""),REDUCE("",names,LAMBDA(acc,nm,LET(inp,N(XLOOKUP(nm,[Щит],[Вводов],1)),'
     'avr,LOWER(TRIM(XLOOKUP(nm,[Щит],[АВР],"нет")&"")),own,IF(inp>=2,IF(avr="да","I","II"),"III"),fr,XLOOKUP(nm,[Щит],[_строка фидера],""),'
     'bt,IF(fr="","",INDEX(Линии[Тип шины],fr)&""),man,TRIM(XLOOKUP(nm,[Щит],[Кат. питающей линии (ручн.)],"")&""),'
     'up,IF(acc="",man,IF(OR(bt="ИБП",bt="ДГУ"),"I особая",IF(bt="Одиночная","III",acc))),'
     'IF(own<>"III",own,IF(OR(up="I особая",up="I",up="II"),"II","III"))))))',
     note="1 ввод — III (или II, если питающая линия I/II: один кабель не даёт I, ПУЭ 1.2.19); 2+ ввода без АВР — II; с АВР — I.")
scol("Iпуск max, А", "h", '=MAXIFS(Линии[Iпуск, А],Линии[Щит],[@Щит])')
scol("Источник фидера", "h", '=IF([@[_строка фидера]]="","",INDEX(Линии[Источник для режимов],[@[_строка фидера]]))')
scol("Пожарный фидер", "h", '=IF([@[_строка фидера]]="","",INDEX(Линии[Пожарная (итог)],[@[_строка фидера]]))')
scol("Замечания щита", "c",
     '=TEXTJOIN("; ",TRUE,IF(AND([@[Питается от (итог)]]="",[@[Zпетли на шинах, Ом]]=""),"нет данных сети для ТКЗ: введите Iкз от сетевой организации или трансформатор",""),'
     'IF(AND([@[Питается от]]<>"",NOT(ISFORMULA([@[Питается от]])),[@[_строка фидера]]=""),"нет линии-фидера в «Линиях»: номер линии = имя щита или столбец «Питает щит»",""),'
     'IF(AND(ISNUMBER([@[ΔU до шин, %]]),[@[ΔU до шин, %]]>4),"ΔU до шин "&FIXED([@[ΔU до шин, %]],2)&" % > 4 %",""))')


# фазы: автоматы щита без ручной фазы — по убыванию мощности группы на наименее загруженную фазу (ручные и 3ф — заданы).
# Ручная фаза читается через ISFORMULA (цикла нет: серая формула не вычисляется). Результат — «;№ строки:Lx;…».
scol("_фазы авто", "h",
     '=LET(b,[@Щит],nn,Линии[_n],m,(Линии[Щит]=b)*(Линии[_свой автомат]=1),man,Линии[_фаза вручную],'
     'фз_,IFERROR(--Линии[_фаз группы],1),gp,IF(Линии[Группа Кс (итог)]="Нет",0,IFERROR(--Линии[_Р группы],0)),'
     'тр_,SUM(m*(man="L1,L2,L3")*gp)/3+SUM(m*(man="")*(фз_>=3)*gp)/3,л1_,SUM(m*(man="L1")*gp)+тр_,л2_,SUM(m*(man="L2")*gp)+тр_,'
     'л3_,SUM(m*(man="L3")*gp)+тр_,sel,m*(man="")*(фз_<3),IF(SUM(sel)=0,"",LET(ids,FILTER(nn,sel),ld,FILTER(gp,sel),'
     'ord,SORTBY(HSTACK(ids,ld),ld,-1,ids,1),res,REDUCE(HSTACK(л1_,л2_,л3_,""),SEQUENCE(ROWS(ord)),LAMBDA(acc,k,LET(фа,INDEX(acc,1,1),'
     'фб,INDEX(acc,1,2),фв,INDEX(acc,1,3),нг,INDEX(ord,k,2),ф,IF(фа<=MIN(фб,фв),1,IF(фб<=фв,2,3)),'
     'HSTACK(фа+(ф=1)*нг,фб+(ф=2)*нг,фв+(ф=3)*нг,INDEX(acc,1,4)&";"&INDEX(ord,k,1)&":L"&ф)))),INDEX(res,1,4)&";")))',
     note="Расклад фаз автоматов без ручной фазы (крупные первыми).")

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
    wo.Range("B16").Value = "ИБП (для щитов с шиной ИБП)"
    wo.Range("B16").Font.Bold = True
    names = {}
    for i, (a, v, u, nm) in enumerate(UPS_SET):
        r = 17 + i
        wo.Range(f"B{r}").Value = a
        wo.Range(f"C{r}").Value = v
        wo.Range(f"D{r}").Value = u
        wo.Range(f"C{r}").Interior.Color = C_IN
        names[nm] = f"=Объект!$C${r}"
    r0 = 17 + len(UPS_SET) + 1
    wo.Range(f"B{r0}").Value = "Ряд типоразмеров ИБП, кВА"
    wo.Range(f"C{r0}:{letters(2 + len(UPS_SIZES))}{r0}").Value = [UPS_SIZES]
    wo.Range(f"C{r0}:{letters(2 + len(UPS_SIZES))}{r0}").Interior.Color = C_IN
    names["ИБП_ряд"] = f"=Объект!$C${r0}:${letters(2 + len(UPS_SIZES))}${r0}"
    set_names(wb, names)

    ws.Range("B1").Value = "РЕЖИМЫ: нагрузка выбранного щита (вместе со всеми нижестоящими) по вводам в разных режимах"
    ws.Range("E3").Value = ("Как читать: столбцы — источники (Ввод 1/2/3, ИБП, ДГУ), блоки — режимы. На «Нормальный» и «Аварийный» "
                            "проверяются вводы, аппараты и кабели (берётся больший). Пожарные — отдельно, в максимум не входят. Итоги по каждому вводу щита — лист «Щит».")
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
    ws.Range("AA2:AD2").Value = [["в составе", "нижний щит", "источник", "пожарная"]]
    F2(ws.Range("AA3"), '=--(ISNUMBER(SEARCH("|"&$C$3&"|",Линии[_путь щита]))*(Линии[Номер линии]<>""))')
    F2(ws.Range("AB3"), '=IFERROR(TEXTBEFORE(TEXTAFTER(Линии[_путь щита],"|"&$C$3&"|"),"|"),"")')
    F2(ws.Range("AC3"), '=IF(AB3#="",Линии[Источник для режимов],XLOOKUP(AB3#,Щиты[Щит],Щиты[Источник фидера],"Ввод 1"))')
    F2(ws.Range("AD3"), '=IF(AB3#="",Линии[Пожарная (итог)],IF(XLOOKUP(AB3#,Щиты[Щит],Щиты[Пожарный фидер],"нет")="да","да",Линии[Пожарная (итог)]))')
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
    # ИБП
    ws.Range(f"B{r}").Value = "7. ИБП ВЫБРАННОГО ЩИТА (нагрузка шин ИБП, без пожарных)"
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
    "Лист задаёт настройки ВСЕГО объекта. Обычно меняют только C3 (название), C8 (методика), C9 (kп.к) и C10–C12 (наименования, марка для пожарных линий).",
    "C10 Наименование линии: «по коду и № помещения» — «Розеточная группа, пом. 105»; «по экспликации» — название помещения из таблицы «Экспликация» (лист «Коды»). C11 — писать ли «пом.». C12 — марка кабеля для пожарных линий (обозначение из «Коды» → «Кабели»). C13–C14 — линии до 50 В (LED 24 В): запас мощности БП (1,3) и допустимые потери от БП (5 %).",
    "C3 Объект — название и адрес. Идёт в таблицу для пояснительной записки.",
    "C4–C7 Ток КЗ: ρ меди/алюминия при нагреве, переходные сопротивления на каждый щит, c_min. Влияют на Iкз min в конце линий и проверку отключения ≤ 0,4 с. Значения по ГОСТ 28249 — не меняйте без причины.",
    "C8 Методика: «нормы 2026» — расчёт по СП 256 с Изм. 9 (cos φ силовых 0,8, Кс силовых по табл. 7.5, Ко только при бытовой + силовой нагрузке, отопление и ВНГ отдельно). «v3.12» — как старая книга, только для сверки.",
    "C9 kп.к — региональный коэффициент к бытовой нагрузке (СП 256 табл. 7.5а). Частный дом — 1. Многоквартирный дом — значение региона (ЦФО 0,81).",
    "F2:G14 Кс силовых по числу электроприёмников (СП 256 табл. 7.5, доля сантехники 100–85 %). Применяется к силовым линиям без ручного Кс в режиме «нормы 2026».",
    "B15:C22 ИБП — допустимая загрузка, cos φ, автономия, напряжение АКБ, КПД, разряд, заряд. Влияют на подбор ИБП и на нагрузку входа ИБП на вводе (лист «Режимы»).",
    "Где ещё настройки: Uн, общий К прокладки, таблица Кс быт, мест на лист схемы — лист «Таблицы_норм» (BF4, BQ29, BQ30, BP4:BQ11).",
]
SHIELD_GUIDE = [
    "ЩИТЫ: одна строка — один щит объекта. Жёлтые столбцы — ввод, серые — результат. Серый курсив во вводе — выбор книги, синее — ваше. Свёрнутые столбцы (кнопка «+») — подробности расчёта.",
    "Для КАЖДОГО щита: Щит (имя, как в «Линиях»), Вводов (1/2/3), АВР (да/нет), ИБП от (Ввод 1/2/3), Мест под вводной узел (17 — весь первый лист, 11 — узел на части листа).",
    "«Питается от» серым книга находит сама по линии-фидеру в «Линиях» (номер линии = имя щита, или столбец «Питает щит»). Вписывайте, только если фидера в книге нет.",
    "Только для ВРУ / ГРЩ (щит без вышестоящего): данные сети — Iкз min и Iкз max от сетевой организации, или «ТП» и мощность трансформатора; питающий кабель и аппарат. Без них ток КЗ не считается ни в одном щите.",
    "Результат: Руст, Рр, Iр, cos φ — щит вместе со всеми нижестоящими; ΔU до шин и Iкз — от ВРУ по фидерам; «Категория шины «Ввод»» — какую категорию дают вводы щита; «Замечания щита» — что исправить. Итоги по каждому вводу — лист «Щит».",
]
BUS_GUIDE = [
    "ТАБЛИЦА «ШИНЫ» (справа от таблицы щитов): нужна, только если тип шины НЕ определяется сам. Обычно пустая.",
    "Как определяется сам: в «Линиях» в столбце «Шина» (в строке с автоматом) пишется имя шины. Имя содержит «ИБП» → шина ИБП (категория I особая); «ДГУ» → шина ДГУ; иначе — основная шина секции линии (Ввод 1/2/3 по столбцу «Секция»).",
    "Когда заполнять: имя шины не говорит о её типе. Пример: шина «ЩГП» питается от ИБП → строка: Шина = ЩГП, Щит = ЩР-1 (или пусто — для всех щитов), Тип = ИБП. Пример: «Шина вентиляции» без резерва → Тип = Одиночная (категория III).",
    "Типы: Ввод 1 / Ввод 2 / Ввод 3 — источник по секции; ИБП, ДГУ — отдельный источник (I особая); Одиночная — без резерва (III). Тип влияет на категорию ввода и на режимы (лист «Режимы»).",
]


def put_guides(wb):
    """гайды на листах «Объект» (снизу), «Щиты» и «Шины» (сверху, над таблицей)."""
    wo = wb.Worksheets("Объект")
    guide_block(wo, 28, "КАК НАСТРАИВАТЬ ЛИСТ «ОБЪЕКТ»", OBJ_GUIDE, 3)
    ws = wb.Worksheets("Щиты")
    lines = SHIELD_GUIDE + BUS_GUIDE
    ws.Rows(f"1:{len(lines) + 2}").Insert()
    guide_block(ws, 1, "КАК ЗАПОЛНЯТЬ ЛИСТ «ЩИТЫ» (и таблицу «Шины» справа)", lines, 12)


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
_MAIN = ["Замечания щита", "Линия фидера", "Руст, кВт", "Рр, кВт", "Iр, А", "cos φ", "ΔU до шин, %",
         "Iкз min на шинах, А", "Iкз max на шинах, кА", "Категория шины «Ввод»"]
for _c in S:
    if _c["kind"] == "c" and _c["head"] not in _MAIN:
        _c["kind"] = "h"
_S_in = [c for c in S if c["kind"] in ("in", "a")]
_S_main = [next(c for c in S if c["head"] == h) for h in _MAIN]
_S_rest = [c for c in S if c["kind"] not in ("in", "a") and c["head"] not in _MAIN]
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


_dup = [h for h, n in __import__("collections").Counter(c["head"].lower() for c in L).items() if n > 1] +        [h for h, n in __import__("collections").Counter(c["head"] for c in S).items() if n > 1]
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
# Документы v5 (пункт 2): справочник диаметров, кабельный журнал по щиту, спецификация по щиту.
# Вставляется make_build_v5.py перед хвостом сборки; пользуется F2, validation, set_names, make_table, C_IN, C_HEAD.
# ------------------------------------------------------------------------------------------------

# Наружные диаметры, мм, для подбора гофры. Ориентировочно: АВВГнг-LS 1 кВ (тот же размер жилы и изоляции), типовые каталоги;
# FRLS — больше LS на слюдяную ленту. Уточняйте по производителю: лист «Справочная_потребители», столбец U.
CABLE_DIAM_ADD = (
    [("ВВГнг(А)-LS", 1, s, d) for s, d in ((10, 9.2), (16, 10.1), (25, 11.6), (35, 12.6), (50, 14.3), (70, 16.0), (95, 18.0), (120, 19.6),
                                           (150, 21.6), (185, 23.8), (240, 26.8))]
    + [("ВВГнг(А)-LS", 2, s, d) for s, d in ((10, 14.5), (16, 16.4), (25, 19.7), (35, 21.7), (50, 25.5))]
    + [("ВВГнг(А)-LS", 3, s, d) for s, d in ((10, 15.3), (16, 17.3), (25, 20.9), (35, 23.4), (50, 26.6), (70, 26.7), (95, 30.4), (120, 33.0),
                                           (150, 37.2), (185, 40.5), (240, 46.3))]
    + [("ВВГнг(А)-LS", 4, s, d) for s, d in ((10, 16.7), (16, 19.3), (25, 22.9), (35, 25.7), (50, 26.1), (70, 29.7), (95, 34.5), (120, 37.5),
                                           (150, 41.1), (185, 44.9), (240, 51.2))]
    + [("ВВГнг(А)-LS", 5, s, d) for s, d in ((10, 18.6), (16, 21.1), (25, 24.5), (35, 27.2), (50, 32.3), (70, 36.6), (95, 41.9), (120, 46.0),
                                           (150, 51.4), (185, 56.8))]
    + [("ВВГнг(А)-FRLS", 1, s, d) for s, d in ((10, 11.2), (16, 12.1), (25, 13.6), (35, 14.6), (50, 16.3), (70, 18.0), (95, 20.0), (120, 21.6),
                                             (150, 23.6), (185, 25.8), (240, 28.8))]
    + [("ВВГнг(А)-FRLS", 3, s, d) for s, d in ((10, 21.3), (16, 23.3), (25, 26.9), (35, 29.4), (50, 32.6))]
    + [("ВВГнг(А)-FRLS", 5, s, d) for s, d in ((10, 24.6), (16, 27.1), (25, 30.5), (35, 33.2), (50, 38.3), (70, 42.6), (95, 47.9), (120, 52.0))]
    + [("ППГнг(А)-HF", n, s, d) for n, s, d in ((3, 10, 15.3), (3, 16, 17.3), (5, 10, 18.6), (5, 16, 21.1), (5, 25, 24.5), (5, 35, 27.2), (5, 50, 32.3))]
    + [("ВБШвнг(А)-LS", n, s, d) for n, s, d in ((3, 1.5, 13.6), (3, 2.5, 14.6), (3, 4, 15.8), (3, 6, 17.0), (5, 1.5, 15.2), (5, 2.5, 16.6),
                                                (5, 4, 18.2), (5, 6, 19.6), (5, 10, 22.5), (5, 16, 25.0))]
    + [("@МКЭШ", n, s, d) for n, s, d in ((2, 0.5, 6.4), (2, 0.75, 6.9), (2, 1, 7.3), (2, 1.5, 8.2), (3, 0.5, 6.8), (3, 0.75, 7.3), (3, 1, 7.7),
                                         (3, 1.5, 8.7), (4, 0.75, 7.9), (5, 0.75, 8.6))]
)
# слаботочные и кабели с сечением в марке (в «Кабели» — «Сечение в марке» = 2): ключ = марка целиком
CABLE_DIAM_PLAIN = [("FTP-Cat.6-4х2х0.52", 7.2), ("FTP 5e 4x2x0.51", 6.0), ("J-Y(St)Y 2х2х0.8", 5.8), ("КИПЭВ 2х2х0.6", 8.5),
                    ("КСПВГ 4х0.35", 4.0), ("КВБбШвнг(А)-LS 14x2.5", 21.0)]


def _num_ru(x):
    return ("%g" % x).replace(".", ",")


def extend_cable_diameters(wb):
    """дописать в «Кабель_диаметр» (Справочная_потребители Q:X) крупные сечения, FRLS, ВБШв, МКЭШ и слаботочку."""
    ws = wb.Worksheets("Справочная_потребители")
    qv = ws.Range("Q4:Q503").Value
    tv = ws.Range("T4:T503").Value
    have = set(str(r[0]) for r in tv if r[0])
    last = max([i for i, r in enumerate(qv) if r[0] not in (None, "")] + [0]) + 4
    mk = "МКЭШнг(А)-LS"
    try:                                                     # марка МКЭШ — как в «Кабели» (бывает латинская A)
        for r in wb.Worksheets("Коды").ListObjects("Кабели").DataBodyRange.Value:
            if str(r[1] or "").startswith("МКЭШ"):
                mk = str(r[1]).strip()
                break
    except Exception:
        pass
    row, n_add = last + 1, 0
    for brand, n, s, d in CABLE_DIAM_ADD:
        brand = mk if brand == "@МКЭШ" else brand
        key = "%s (%dх%s)" % (brand, n, _num_ru(s))
        if key in have or row > 503:
            continue
        ws.Cells(row, 17).Value = brand
        ws.Cells(row, 18).Value = n
        ws.Cells(row, 19).Value = s
        ws.Cells(row, 20).Formula = '=Q%d&" ("&R%d&"х"&S%d&")"' % (row, row, row)
        ws.Cells(row, 21).Value = d
        row, n_add = row + 1, n_add + 1
    for key, d in CABLE_DIAM_PLAIN:
        if key in have or row > 503:
            continue
        ws.Cells(row, 17).Value = key
        ws.Cells(row, 20).Value = key
        ws.Cells(row, 21).Value = d
        row, n_add = row + 1, n_add + 1
    for r in range(last + 1, row):                          # мин. диаметр трубы и размеры гофры — как в строках шаблона
        ws.Cells(r, 22).Formula = "=IF((U%d*1.2)<14,14,U%d*1.2)" % (r, r)
        ws.Cells(r, 23).Formula = '="Ø"&VLOOKUP(V%d,$Z$4:$AA$10,2,TRUE)' % r
        ws.Cells(r, 24).Formula = '="Ø"&VLOOKUP(V%d,$AD$4:$AE$10,2,TRUE)' % r
    if n_add:
        c = ws.Cells(last + 1, 25)
        c.Value = "↓ добавлено v5: крупные сечения, FRLS, ВБШв, МКЭШ, слаботочка — диаметры ориентировочные, уточните по производителю"
        c.Font.Color = 8421504
    log("справочник диаметров: добавлено %d кабелей" % n_add)


# фильтр строк журнала / спецификации: линии с кабелем; ввод щита — только если фидера щита в книге нет (иначе кабель дважды)
def _jm(sel_all, sel_val):
    return ('(Линии[Номер линии]<>"")*(Линии[_кабель]<>"")*(1-(Линии[_ввод щита]=1)*(IFERROR(XLOOKUP(Линии[Щит],Щиты[Щит],Щиты[_строка фидера]),"")<>""))'
            '*IF(' + sel_all + ',1,Линии[Щит]=' + sel_val + ')')


JSEL = 'sel,TRIM($E$1&""),all,OR(sel="",sel="все"),m,' + _jm("all", "sel") + ','
# гофра по типу и диаметру: tp — «ПВХ|Ø20», ln — длина; u — уникальные; s — сумма; bh — бухта, м
GOFRA = ('cs,Линии[_кабель с сечением],pv,IFERROR(--Линии[_гофра ПВХ],0)*m,pn,IFERROR(--Линии[_гофра ПНД],0)*m,'
         'dpv,IFERROR(VLOOKUP(cs,Кабель_диаметр,4,FALSE),"Ø?"),dpn,IFERROR(VLOOKUP(cs,Кабель_диаметр,5,FALSE),"Ø?"),'
         'tp,VSTACK(IF(pv>0,"ПВХ|"&dpv,""),IF(pn>0,"ПНД|"&dpn,"")),ln,VSTACK(pv,pn),u,UNIQUE(FILTER(tp,tp<>"","")),')


def build_journal_v5(wb, sep):
    """кабельный журнал v5: раскладка v3.12 + числа гофры, выбор щита (E1), сводки: кабели, гофра, лоток, ПТО."""
    ws = wb.Worksheets.Add(After=wb.Worksheets("Расчёт_нагрузок"))
    ws.Name = "Кабельный_журнал"
    ws.Range("A1").Value = "КАБЕЛЬНЫЙ ЖУРНАЛ"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("D1").Value = "Щит:"
    ws.Range("D1").Font.Bold = True
    ws.Range("D1").HorizontalAlignment = -4152
    ws.Range("E1").Value = "все"
    ws.Range("E1").Interior.Color = C_IN
    ws.Range("E1").Font.Bold = True
    v = ws.Range("E1").Validation
    v.Delete()
    v.Add(3, 1, 1, "=Список_щитов")
    v.ShowError = False
    ws.Range("F1").Value = "все — весь объект. Для AutoCAD — кнопка «Выгрузить для AutoCAD» на листе «Щит»."
    ws.Range("F1").Font.Color = 8421504
    ws.Range("K1").Value = "Запас длины:"
    ws.Range("K1").HorizontalAlignment = -4152
    ws.Range("L1").Value = 0.2
    ws.Range("L1").NumberFormat = "0%"
    ws.Range("L1").Interior.Color = C_IN
    try:
        wb.Names("Запас_кабеля").Delete()
    except Exception:
        pass
    wb.Names.Add("Запас_кабеля", "=Кабельный_журнал!$L$1")
    hdr = [("A2:A4", "Маркировка кабеля"), ("B2:C2", "Трасса"), ("B3:B4", "Начало"), ("C3:C4", "Конец"),
           ("D2:H2", "Кабель"), ("D3:H3", "По проекту"), ("I3:K3", "Проложено"), ("I2:K2", "Кабель"),
           ("L2:L4", "Испытания"), ("M2:M4", "Дата прокладки"), ("O2:T2", "Для подсчёта (числа)")]
    for a, t in hdr:
        rng = ws.Range(a)
        rng.Merge()
        rng.Value = t
    ws.Range("D4:K4").Value = [["Количество кабелей и сечение", "Лоток", "Труба гофрированная ПВХ", "Труба гофрированная ПНД",
                                "Длина трассы, м", "Марка", "Количество кабелей и сечение", "Длина, м"]]
    ws.Range("O4").NumberFormat = "@"
    ws.Range("O3:T4").Value = [["Потолок / пол", "Гофра ПВХ", "", "Гофра ПНД", "", "Лоток"], ["1 / 2", "Ø", "м", "Ø", "м", "м"]]
    h = ws.Range("A2:T4")
    h.Font.Bold = True
    h.WrapText = True
    h.HorizontalAlignment = -4108
    h.VerticalAlignment = -4108
    ws.Range("A2:M4").Interior.Color = C_HEAD
    ws.Range("A2:M4").Borders.LineStyle = 1
    ws.Range("O2:T4").Interior.Color = 15132390
    ws.Range("O2:T4").Borders.LineStyle = 1
    ws.Range("N2:N4").UnMerge()
    F2(ws.Range("A5"), ('=LET(' + JSEL + 'n,SUM(m),IF(n=0,"нет линий с кабелем",LET(cs,FILTER(Линии[_кабель с сечением],m),'
                        'tr,FILTER(IFERROR(--Линии[Лоток],0),m),pv,FILTER(IFERROR(--Линии[_гофра ПВХ],0),m),pn,FILTER(IFERROR(--Линии[_гофра ПНД],0),m),'
                        'fl,FILTER(Линии[Потолок - 1/Пол - 2]&"",m),dpv,IFERROR(VLOOKUP(cs,Кабель_диаметр,4,FALSE),"Ø?"),dpn,IFERROR(VLOOKUP(cs,Кабель_диаметр,5,FALSE),"Ø?"),'
                        'HSTACK(FILTER(Линии[Номер линии],m),FILTER(Линии[Щит],m),FILTER(Линии[Наименование (итог)],m),cs,'
                        'IF(tr=0,"","L="&tr),IF(pv=0,"",dpv&" L="&pv),IF(pn=0,"",dpn&" L="&pn),FILTER(Линии[Длина, м],m),EXPAND({""},n,6,""),'
                        'fl,IF(pv=0,"",dpv),IF(pv=0,"",pv),IF(pn=0,"",dpn),IF(pn=0,"",pn),IF(tr=0,"",tr)))))'))
    # сводки (справа): кабели, гофра, лоток, ПТО
    ws.Range("V2").Value = "СВОДКА ПО ЖУРНАЛУ (выбранный щит; запас — L1)"
    ws.Range("V2").Font.Bold = True
    blocks = [("V4:Y4", ["Кабель (марка, сечение)", "Длина трассы, м", "С запасом, м", "В спецификацию, м (вверх до 10)"]),
              ("AA4:AE4", ["Гофра", "Ø", "Длина, м", "С запасом, м", "В спецификацию, м (по бухтам)"]),
              ("AG4:AH4", ["Лоток, м", "С запасом, м"]),
              ("AJ4:AN4", ["Кабель", "Потолок / пол", "Артикул ПТО", "Длина, м", "С запасом, м"])]
    for a, t in blocks:
        ws.Range(a).Value = [t]
        ws.Range(a).Font.Bold = True
        ws.Range(a).WrapText = True
        ws.Range(a).Interior.Color = C_HEAD
    ws.Range("AA3").Value = "ГОФРА (бухты — «Спецификация», таблица «Гофра_изделия»)"
    ws.Range("AJ3").Value = "СВОДКА ПТО (кабель + потолок / пол → артикул ПТО, «Таблицы_норм» AJ:AK)"
    for a in ("AA3", "AJ3"):
        ws.Range(a).Font.Bold = True
    F2(ws.Range("V5"), ('=LET(' + JSEL + 'k,Линии[_кабель с сечением],ln,IFERROR(--Линии[Длина, м],0),uu,UNIQUE(FILTER(k,m,"")),'
                        'IF(INDEX(uu,1)="","—",LET(lsum,MAP(uu,LAMBDA(x,SUM(m*(k=x)*ln))),u,FILTER(uu,lsum>0,""),'
                        'IF(INDEX(u,1)="","— (длины не заданы)",LET(l,FILTER(lsum,lsum>0),br,IFERROR(TEXTBEFORE(u," ("),u),'
                        'co,IFERROR(--TEXTBEFORE(TEXTAFTER(u,"(",-1),"х"),0),se,IFERROR(--TEXTBEFORE(TEXTAFTER(u,"х",-1),")"),0),'
                        'SORTBY(HSTACK(u,l,ROUNDUP(l*(1+Запас_кабеля),0),CEILING.MATH(l*(1+Запас_кабеля),10)),br,1,co,-1,se,-1))))))'))
    F2(ws.Range("AG5"), '=LET(' + JSEL + 'l,SUM(m*IFERROR(--Линии[Лоток],0)),HSTACK(l,ROUNDUP(l*(1+Запас_кабеля),0)))')
    F2(ws.Range("AJ5"), ('=LET(' + JSEL + 'cs,Линии[_кабель с сечением],fl,Линии[Потолок - 1/Пол - 2]&"",ln,IFERROR(--Линии[Длина, м],0),'
                         'key,cs&" "&fl,u,UNIQUE(FILTER(key,m*(fl<>""),"")),'
                         'IF(INDEX(u,1)="","— (нет линий с «Потолок / пол»)",LET(c,TEXTBEFORE(u," ",-1),f,TEXTAFTER(u," ",-1),'
                         'l,MAP(u,LAMBDA(x,SUM(m*(key=x)*ln))),a,IFERROR(XLOOKUP(u,INDEX(Справочная_ПТО,0,1),INDEX(Справочная_ПТО,0,2)),"нет в справочнике"),'
                         'HSTACK(c,IF(f="1","потолок",IF(f="2","пол",f)),a,l,ROUNDUP(l*(1+Запас_кабеля),0)))))'))
    for c_, w_ in (("A", 16), ("B", 12), ("C", 40), ("D", 26), ("E", 9), ("F", 14), ("G", 14), ("H", 9), ("I", 12), ("J", 14), ("K", 9),
                   ("L", 11), ("M", 11), ("N", 2), ("O", 8), ("P", 7), ("Q", 7), ("R", 7), ("S", 7), ("T", 7), ("U", 2),
                   ("V", 30), ("W", 11), ("X", 11), ("Y", 14), ("Z", 2), ("AA", 8), ("AB", 7), ("AC", 10), ("AD", 10), ("AE", 14), ("AF", 2),
                   ("AG", 9), ("AH", 10), ("AI", 2), ("AJ", 30), ("AK", 10), ("AL", 14), ("AM", 10), ("AN", 10)):
        ws.Columns(c_).ColumnWidth = w_
    ws.Rows(4).RowHeight = 45
    ws.Range("A5").Select()
    return ws


# --- спецификация по щиту
JOURNAL_GOFRA_F = ('=LET(' + JSEL + GOFRA + 'IF(INDEX(u,1)="","—",LET(t,TEXTBEFORE(u,"|"),d,TEXTAFTER(u,"|"),'
                   'dn,IFERROR(--SUBSTITUTE(d,"Ø",""),999),s,MAP(u,LAMBDA(x,SUM((tp=x)*ln))),'
                   'bh,IFERROR(XLOOKUP(t&"|"&dn,Гофра_изделия[Тип]&"|"&Гофра_изделия[Ø],Гофра_изделия[Бухта, м]),10),'
                   'SORTBY(HSTACK(t,d,s,ROUNDUP(s*(1+Запас_кабеля),0),CEILING.MATH(s*(1+Запас_кабеля),bh)),t,1,dn,1))))')

SPEC_MANUAL_HEADS = ["Щит", "Раздел (1/2)", "Наименование и техническая характеристика", "Тип, марка", "Код", "Завод", "Ед.", "Кол-во", "Примечание"]
SPEC_NORM_HEADS = ["Наименование и техническая характеристика", "Тип, марка", "Код", "Завод", "Ед.", "База", "На 1 м", "Кратность"]
# нормы — по спецификации ЩР-1.1 проекта «Жуковка»: клипсы 2 шт/м гофры ПВХ; гвозди и баллоны — по 1000 на клипсы; лента — 1 на ~330 м ПНД
SPEC_NORMS = [
    ["Держатель-клипса для монтажного пистолета, д.20мм, серый", "", "51020M", "DKC", "шт.", "ПВХ Ø20", 2, 1],
    ["Держатель с защелкой, д.25мм", "", "51025", "DKC", "шт.", "ПВХ Ø25", 2, 1],
    ["Держатель с защелкой, д.32мм", "", "", "DKC", "шт.", "ПВХ Ø32", 2, 1],
    ["Гвозди кассетные для гвоздезабивного пистолета 3,05х32 мм, упак. 1000 шт.", "", "30532stepEG", "TOUA", "упак.", "ПВХ все", 0.002, 1],
    ["Баллон газовый на 1000 гвоздей", "", "toua-fcg", "TOUA", "шт.", "ПВХ все", 0.002, 1],
    ["Лента монтажная перфорированная, 25х0,8", "", "CM610050", "DKC", "шт.", "ПНД все", 0.003, 1],
]
SPEC_GOFRA_HEADS = ["Тип", "Ø", "Наименование", "Код", "Завод", "Бухта, м"]
# артикулы — только подтверждённые спецификацией проекта; остальные впишите
SPEC_GOFRA = [["ПВХ", 16, "Труба ПВХ гибкая гофр. д.16мм, тяжёлая с протяжкой, 100м, цвет серый", "", "DKC", 100],
              ["ПВХ", 20, "Труба ПВХ гибкая гофр. д.20мм, тяжёлая с протяжкой, 100м, цвет серый", "91520", "DKC", 100],
              ["ПВХ", 25, "Труба ПВХ гибкая гофр. д.25мм, тяжёлая с протяжкой, 50м, цвет серый", "91525", "DKC", 50],
              ["ПВХ", 32, "Труба ПВХ гибкая гофр. д.32мм, тяжёлая с протяжкой, 25м, цвет серый", "", "DKC", 25],
              ["ПВХ", 40, "Труба ПВХ гибкая гофр. д.40мм, тяжёлая с протяжкой, 15м, цвет серый", "", "DKC", 15],
              ["ПВХ", 50, "Труба ПВХ гибкая гофр. д.50мм, тяжёлая с протяжкой, 15м, цвет серый", "", "DKC", 15],
              ["ПНД", 16, "Труба ПНД гибкая гофр. д.16мм, тяжёлая с протяжкой, 100м, цвет оранжевый", "", "DKC", 100],
              ["ПНД", 20, "Труба ПНД гибкая гофр. д.20мм, тяжёлая с протяжкой, 100м, цвет оранжевый", "71520", "DKC", 100],
              ["ПНД", 25, "Труба ПНД гибкая гофр. д.25мм, тяжёлая с протяжкой, 50м, цвет оранжевый", "71525", "DKC", 50],
              ["ПНД", 32, "Труба ПНД гибкая гофр. д.32мм, тяжёлая с протяжкой, 25м, цвет оранжевый", "71532", "DKC", 25],
              ["ПНД", 40, "Труба ПНД гибкая гофр. д.40мм, тяжёлая с протяжкой, 15м, цвет оранжевый", "", "DKC", 15],
              ["ПНД", 50, "Труба ПНД гибкая гофр. д.50мм, тяжёлая с протяжкой, 15м, цвет оранжевый", "", "DKC", 15]]

O_ = "Однолинейка!"
SPEC_F = (
    '=LET(sel,Щит_спец,all,sel="*",bn,IF(all,"(весь объект)",sel),TX,LAMBDA(x,IF(x="","",x)),'
    # раздел 1: вручную (корпус и комплект)
    'mt,Спец_вручную[Наименование и техническая характеристика],mr,IF(Спец_вручную[Раздел (1/2)]&""="2",2,1),'
    'mok,(mt<>"")*IF(all,1,Спец_вручную[Щит]=sel),'
    'mv,HSTACK(TX(mt),TX(Спец_вручную[Тип, марка]),TX(Спец_вручную[Код]),TX(Спец_вручную[Завод]),TX(Спец_вручную[Ед.]),TX(Спец_вручную[Кол-во]),EXPAND({""},ROWS(mt),1,""),TX(Спец_вручную[Примечание])),'
    'mv_a,FILTER(mv,mok*(mr=1),""),mv_b,FILTER(mv,mok*(mr=2),""),'
    # раздел 1: аппараты по «Каталогу» (Спецификация_авто) и то, чего в «Каталоге» нет
    'sa,Спецификация_авто!$C$4:$J$160,ap,FILTER(TX(sa),INDEX(sa,0,1)<>"",""),'
    'mask,IF(all,1,--(' + O_ + '$A$3:$A$3000=sel)),'
    'hb,TOCOL(IF(mask,' + O_ + '$H$3:$H$3000,"")),hx,FILTER(hb,(hb<>"")*(COUNTIF(Каталог!$X$4:$X$500,hb)=0),""),'
    'hu,UNIQUE(hx),hn,MAP(hu,LAMBDA(x,SUM(--(hb=x)))),'
    'nb,IF(INDEX(hu,1)="","",HSTACK(IF(ISNUMBER(SEARCH("мА",hu)),"Автоматический выключатель дифференциального тока ","Автоматический выключатель ")&hu,'
    'hu,EXPAND({""},ROWS(hu),2,""),EXPAND({"шт."},ROWS(hu),1,"шт."),hn,EXPAND({""},ROWS(hu),1,""),EXPAND({"нет в «Каталоге» — впишите артикул"},ROWS(hu),1,"нет в «Каталоге» — впишите артикул"))),'
    'ha,TOCOL(IF(mask,' + O_ + '$I$3:$K$3000,"")),ax,FILTER(ha,(ha<>"")*(COUNTIF(Каталог!$Z$4:$Z$500,ha)=0),""),'
    'au,UNIQUE(ax),an,MAP(au,LAMBDA(x,SUM(--(ha=x)))),'
    'na,IF(INDEX(au,1)="","",HSTACK(au,au,EXPAND({""},ROWS(au),2,""),EXPAND({"шт."},ROWS(au),1,"шт."),an,EXPAND({""},ROWS(au),1,""),'
    'EXPAND({"нет в «Каталоге» — впишите артикул"},ROWS(au),1,"нет в «Каталоге» — впишите артикул"))),'
    # раздел 2: кабели (несколько кабелей в линии «5хВВГнг… (1х185)» — длина × число), гофра, крепёж по нормам
    'm,' + _jm("all", "sel") + ',k,Линии[_кабель с сечением],dg_,ISNUMBER(--LEFT(k,1)),nl_,IFERROR(--TEXTBEFORE(k,"х"),""),'
    'nn,IF(dg_,IF(ISNUMBER(nl_),nl_,IFERROR(--LEFT(k,1),1)),1),bk,IF(dg_,IF(ISNUMBER(nl_),TRIM(TEXTAFTER(k,"х")),TRIM(MID(k,2,999))),k),'
    'lc,IFERROR(--Линии[Длина, м],0)*nn,cuu,UNIQUE(FILTER(bk,m,"")),lsum,IF(INDEX(cuu,1)="",0,MAP(cuu,LAMBDA(x,SUM(m*(bk=x)*lc)))),'
    'cu,FILTER(cuu,lsum>0,""),cq,IF(INDEX(cu,1)="",0,CEILING.MATH(FILTER(lsum,lsum>0,0)*(1+Запас_кабеля),10)),'
    'cbr,IFERROR(TEXTBEFORE(cu," ("),cu),cco,IFERROR(--TEXTBEFORE(TEXTAFTER(cu,"(",-1),"х"),0),cse,IFERROR(--TEXTBEFORE(TEXTAFTER(cu,"х",-1),")"),0),'
    'cb,IF(INDEX(cu,1)="","",SORTBY(HSTACK(IF(LEFT(cu,4)="ПуГВ","Провод ","Кабель ")&cu,cu,EXPAND({""},ROWS(cu),2,""),EXPAND({"м"},ROWS(cu),1,"м"),cq,'
    'EXPAND({""},ROWS(cu),2,"")),cbr,1,cco,-1,cse,-1)),'
    + GOFRA +
    'gt,IF(INDEX(u,1)="","",TEXTBEFORE(u,"|")),gd,IF(INDEX(u,1)="","",TEXTAFTER(u,"|")),gdn,IFERROR(--SUBSTITUTE(gd,"Ø",""),999),'
    'gs,IF(INDEX(u,1)="",0,MAP(u,LAMBDA(x,SUM((tp=x)*ln)))),'
    'gk,gt&"|"&gdn,tk,Гофра_изделия[Тип]&"|"&Гофра_изделия[Ø],'
    'gq,IF(INDEX(u,1)="",0,CEILING.MATH(gs*(1+Запас_кабеля),IFERROR(XLOOKUP(gk,tk,Гофра_изделия[Бухта, м]),10))),'
    'gn,IFERROR(XLOOKUP(gk,tk,Гофра_изделия[Наименование]),"Труба гофрированная "&gt&" "&gd),'
    'gf,IF(INDEX(u,1)="","",SORTBY(HSTACK(gn,gt&" "&gd,IFERROR(XLOOKUP(gk,tk,Гофра_изделия[Код])&"",""),IFERROR(XLOOKUP(gk,tk,Гофра_изделия[Завод])&"",""),'
    'EXPAND({"м"},ROWS(u),1,"м"),gq,EXPAND({""},ROWS(u),2,"")),gt,1,gdn,1)),'
    'nq,MAP(Нормы_расхода[База],Нормы_расхода[На 1 м],Нормы_расхода[Кратность],LAMBDA(b,kk,rr,LET(v,IF(b="ПВХ все",SUM(gq*(gt="ПВХ")),'
    'IF(b="ПНД все",SUM(gq*(gt="ПНД")),IF(b="Кабель все",SUM(cq),SUM(gq*(gt&" "&gd=b))))),IF(OR(N(v)=0,N(kk)=0),0,CEILING.MATH(v*kk,IF(N(rr)>0,rr,1)))))),'
    'nf,FILTER(HSTACK(TX(Нормы_расхода[Наименование и техническая характеристика]),TX(Нормы_расхода[Тип, марка]),TX(Нормы_расхода[Код]),TX(Нормы_расхода[Завод]),'
    'TX(Нормы_расхода[Ед.]),nq,EXPAND({""},ROWS(nq),2,"")),(nq>0)*(Нормы_расхода[Наименование и техническая характеристика]<>""),""),'
    # сборка: пустые части отбрасываются, строки нумеруются по разделам
    'CL,LAMBDA(a,LET(v,IFERROR(a,""),FILTER(v,INDEX(v,0,1)<>"",""))),'
    'sec_a,CL(VSTACK(mv_a,ap,nb,na)),sec_b,CL(VSTACK(cb,gf,nf,mv_b)),'
    'n_1,IF(INDEX(sec_a,1,1)="",0,ROWS(sec_a)),n_2,IF(INDEX(sec_b,1,1)="",0,ROWS(sec_b)),'
    'hd_a,HSTACK("1","Вводно-распределительное оборудование "&bn,EXPAND({""},1,7,"")),'
    'hd_b,HSTACK("2","Кабельные изделия (кабель, провод, гофра, крепёж)",EXPAND({""},1,7,"")),'
    'rw_a,IF(n_1=0,hd_a,VSTACK(hd_a,HSTACK("1."&SEQUENCE(n_1),sec_a))),rw_b,IF(n_2=0,hd_b,VSTACK(hd_b,HSTACK("2."&SEQUENCE(n_2),sec_b))),'
    'IFERROR(VSTACK(rw_a,rw_b),""))'
)


def _table_at(ws, name, heads, rows, first_row, first_col):
    """таблица Excel с левым верхним углом (first_row, first_col); rows — список строк (может быть пустым)."""
    n = max(1, len(rows))
    rng_h = ws.Range(ws.Cells(first_row, first_col), ws.Cells(first_row, first_col + len(heads) - 1))
    rng_h.Value = [heads]
    if rows:
        ws.Range(ws.Cells(first_row + 1, first_col), ws.Cells(first_row + n, first_col + len(heads) - 1)).Value = rows
    lo = ws.ListObjects.Add(1, ws.Range(ws.Cells(first_row, first_col), ws.Cells(first_row + n, first_col + len(heads) - 1)), None, 1)
    lo.Name = name
    return lo


def build_spec_v5(wb, sep):
    """«Спецификация» по щиту (D1): раздел 1 — корпус вручную + аппараты по схеме; раздел 2 — кабели, гофра, крепёж по нормам."""
    sp = wb.Worksheets("Спецификация")
    sp.Range("B3:J400").UnMerge()
    sp.Range("B3:J400").Clear()
    sp.Range("B1").ClearContents()                        # заголовок — подпись C1 «Спецификация для щита»
    sp.Range("E1").Value = "все — весь объект. Корпус и комплект щита — таблица справа; нормы крепежа и бухты гофры — там же."
    sp.Range("E1").Font.Color = 8421504
    # таблицы ввода справа
    sp.Range("M1").Value = "КОРПУС И КОМПЛЕКТ ЩИТА — вручную: строка на позицию; «Щит» — для какого щита; раздел 2 — прочие кабельные изделия"
    sp.Range("M1").Font.Bold = True
    lm = _table_at(sp, "Спец_вручную", SPEC_MANUAL_HEADS, [], 2, 13)
    validation(lm.ListColumns("Щит").DataBodyRange, "=Список_щитов", sep)
    validation(lm.ListColumns("Раздел (1/2)").DataBodyRange, "1|2", sep)
    c0 = 13 + len(SPEC_MANUAL_HEADS) + 1                    # таблица норм — правее
    sp.Cells(1, c0).Value = "НОРМЫ РАСХОДА: количество = база (м, после округления) × на 1 м, вверх до кратности"
    sp.Cells(1, c0).Font.Bold = True
    ln_ = _table_at(sp, "Нормы_расхода", SPEC_NORM_HEADS, SPEC_NORMS, 2, c0)
    validation(ln_.ListColumns("База").DataBodyRange,
               "ПВХ все|ПНД все|Кабель все|ПВХ Ø16|ПВХ Ø20|ПВХ Ø25|ПВХ Ø32|ПВХ Ø40|ПВХ Ø50|ПНД Ø16|ПНД Ø20|ПНД Ø25|ПНД Ø32|ПНД Ø40|ПНД Ø50", sep)
    c1 = c0 + len(SPEC_NORM_HEADS) + 1
    sp.Cells(1, c1).Value = "ГОФРА: изделия и бухты (количество в спецификации — вверх до целой бухты)"
    sp.Cells(1, c1).Font.Bold = True
    _table_at(sp, "Гофра_изделия", SPEC_GOFRA_HEADS, SPEC_GOFRA, 2, c1)
    for lo in sp.ListObjects:
        lo.TableStyle = "TableStyleLight9"
        lo.HeaderRowRange.WrapText = True
    kt = wb.Worksheets("Каталог")                          # ток канала модуля (диммеры, DALI) — для проверки линий до 50 В
    kt.Range("Q3").Value = "Ток канала, А (модули: диммер, DALI)"
    kt.Range("Q3").Font.Bold = True
    kt.Range("Q3").WrapText = True
    kt.Range("Q3").Interior.Color = C_HEAD
    kt.Columns("Q").ColumnWidth = 12
    try:
        kt.Range("Q3").AddComment("Модуль, к каналам которого подключены линии до 50 В (LED 24 В): допустимый ток канала. Пусто — не проверяется. Линии на одном канале («=») суммируются.")
    except Exception:
        pass
    import re as _re
    for r_ in range(4, 501):                               # ток канала из названия модуля («…, 4x5A)» → 5 А)
        if kt.Cells(r_, 1).Value == "Модуль" and kt.Cells(r_, 17).Value in (None, ""):
            mm_ = _re.search(r"(\d+)\s*[xх×]\s*(\d+(?:[.,]\d+)?)\s*[AА]\b", str(kt.Cells(r_, 7).Value or ""))
            if mm_:
                kt.Cells(r_, 17).Value = float(mm_.group(2).replace(",", "."))
    F2(sp.Range("B3"), SPEC_F)
    F2(wb.Worksheets("Кабельный_журнал").Range("AA5"), JOURNAL_GOFRA_F)      # гофра в журнале — когда есть «Гофра_изделия»
    for c_, w_ in (("A", 2), ("B", 7), ("C", 60), ("D", 22), ("E", 14), ("F", 12), ("G", 7), ("H", 8), ("I", 7), ("J", 26), ("K", 2), ("L", 2)):
        sp.Columns(c_).ColumnWidth = w_
    for i, w_ in enumerate([10, 8, 48, 14, 12, 10, 7, 8, 18]):
        sp.Columns(13 + i).ColumnWidth = w_
    for i, w_ in enumerate([44, 12, 12, 10, 7, 12, 8, 9]):
        sp.Columns(c0 + i).ColumnWidth = w_
    for i, w_ in enumerate([6, 5, 52, 10, 8, 8]):
        sp.Columns(c1 + i).ColumnWidth = w_
    try:
        wb.Names("Спецификация").RefersTo = "=Спецификация!$B$3:$J$400"
    except Exception:
        pass
    return sp


# --- пункт 4: «Изменения» (снимок выдачи и сравнение), «Наклейки»
SNAP_FIELDS = ["Щит", "Номер линии", "Наименование", "Р, кВт", "Автомат", "Сечение, мм²", "Длина, м", "Фаза", "Кабель", "_Модуль (схема)",
               "Канал (итог)", "Клемма", "Лист / место", "Лист схемы"]
SNAP_LABELS = {3: "наименование", 4: "Р, кВт", 5: "автомат", 6: "сечение", 7: "длина", 8: "фаза", 9: "кабель", 10: "модуль", 11: "канал",
               12: "клемма", 13: "место на схеме"}


def _snap_rng(c):
    return "_выдача!$%s$5:ИНДЕКС(_выдача!$%s$5:$%s$20004;МАКС(1;Выдача_n))" % (c, c, c)


def _name_local(wb, nm, f_local):
    """имя с формулой: Names.Add разбирает формулу по-русски (функции и «;»), поэтому — RefersToLocal."""
    try:
        wb.Names(nm).Delete()
    except Exception:
        pass
    wb.Names.Add(nm, "=_выдача!$A$1")
    wb.Names(nm).RefersToLocal = f_local


def build_changes_v5(wb, sep):
    """скрытый лист «_выдача» (снимок, пишет макрос SX_Snapshot), лист «Изменения» и подсветка изменённых ячеек в «Линиях»."""
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "_выдача"
    ws.Range("A1:A2").Value = [["Изменение"], ["Дата"]]
    ws.Range("B2").NumberFormatLocal = "ДД.ММ.ГГГГ"
    ws.Range("A4").Resize(1, len(SNAP_FIELDS)).Value = [SNAP_FIELDS]
    ws.Range("O4:P4").Value = [["ключ", "история выдач (№ | дата | линий)"]]
    _name_local(wb, "Выдача_n", "=СЧЁТЗ(_выдача!$B$5:$B$20004)")
    letters_ = "ABCDEFGHIJKLMN"
    for j in range(1, len(SNAP_FIELDS) + 1):
        _name_local(wb, "Выдача_%d" % j, "=" + _snap_rng(letters_[j - 1]))
    _name_local(wb, "Выдача_ключ", "=" + _snap_rng("O"))
    # лист «Изменения»
    wc = wb.Worksheets.Add(After=wb.Worksheets("Бирки"))
    wc.Name = "Изменения"
    wc.Range("B1").Value = "ИЗМЕНЕНИЯ ПОСЛЕ ВЫДАЧИ"
    wc.Range("B1").Font.Bold = True
    wc.Range("B1").Font.Size = 14
    wc.Range("B2").Value = "Выдача №"
    F2(wc.Range("C2"), '=IF(Выдача_n=0,"—",_выдача!$B$1&"")')
    wc.Range("D2").Value = "от"
    F2(wc.Range("E2"), '=IF(Выдача_n=0,"",_выдача!$B$2)')
    wc.Range("E2").NumberFormatLocal = "ДД.ММ.ГГГГ"     # NumberFormat «dd.mm.yyyy» портит стили файла
    wc.Range("G1:L2").Merge()                            # подсказка — блоком с переносом (не уходит вправо)
    wc.Range("G1").Value = ("Снимок — кнопка «Зафиксировать выдачу» (над «Линиями»). После неё здесь видно, что изменилось; в «Линиях» изменённые "
                            "ячейки — оранжевые, новые линии — зелёные.")
    wc.Range("G1").WrapText = True
    wc.Range("G1").VerticalAlignment = -4160
    wc.Range("G1").Font.Color = 8421504
    wc.Range("G1").Font.Size = 9
    wc.Range("B4:G4").Value = [["Щит", "Линия", "Что", "Было", "Стало", "Лист схемы"]]
    wc.Range("J4:L4").Value = [["Щит", "Листы схемы (для таблицы изменений)", "Изменений"]]
    for a in ("B4:G4", "J4:L4"):
        wc.Range(a).Font.Bold = True
        wc.Range(a).Interior.Color = C_HEAD
        wc.Range(a).WrapText = True
    ch = ",".join('CH(Линии[%s],%d,"%s")' % (SNAP_FIELDS[j - 1], j, SNAP_LABELS[j]) for j in range(3, 14))
    f = ('=LET(nn,Выдача_n,IF(nn=0,"Выдача не зафиксирована: кнопка «Зафиксировать выдачу» над «Линиями».",LET('
         'ko,Выдача_ключ,mc,(Линии[Номер линии]<>"")*1,kc,Линии[Щит]&"|"&Линии[Номер линии],io,ISNUMBER(XMATCH(kc,ko)),'
         'ic,ISNUMBER(XMATCH(ko,FILTER(kc,mc,""))),TXT,LAMBDA(x,IF(ISNUMBER(x),FIXED(x,2,TRUE),IFERROR(FIXED(--(x&""),2,TRUE),x&""))),'
         'SN,LAMBDA(j,CHOOSE(j,Выдача_1,Выдача_2,Выдача_3,Выдача_4,Выдача_5,Выдача_6,Выдача_7,Выдача_8,Выдача_9,Выдача_10,Выдача_11,Выдача_12,Выдача_13,Выдача_14)),'
         'CH,LAMBDA(cur,j,nm,LET(old,XLOOKUP(kc,ko,SN(j),""),a,TXT(old),b,TXT(cur),'
         'FILTER(HSTACK(Линии[Щит],Линии[Номер линии],IF(mc,nm,""),a,b,Линии[Лист схемы]),mc*io*(a<>b),""))),'
         'ad,FILTER(HSTACK(Линии[Щит],Линии[Номер линии],IF(mc,"добавлена",""),IF(mc,"",""),Линии[Наименование]&"",Линии[Лист схемы]),mc*NOT(io),""),'
         'rm,FILTER(HSTACK(Выдача_1,Выдача_2,IF(ic,"удалена","удалена"),Выдача_3&"",IF(ic,"",""),Выдача_14),NOT(ic),""),'
         'al,IFERROR(VSTACK(ad,rm,' + ch + '),""),fl,FILTER(al,INDEX(al,0,1)<>"",""),'
         'IF(INDEX(fl,1,1)="","нет изменений после выдачи",SORTBY(fl,IFERROR(XMATCH(INDEX(fl,0,1),Щиты[Щит]),999),1,INDEX(fl,0,2),1)))))')
    F2(wc.Range("B5"), f)
    F2(wc.Range("J5"), ('=IFERROR(LET(d,$B$5#,b,INDEX(d,0,1),s,INDEX(d,0,6),u,UNIQUE(b),HSTACK(u,MAP(u,LAMBDA(x,TEXTJOIN(", ",TRUE,'
                        'SORT(UNIQUE(FILTER(s,(b=x)*(s<>""),"")))))),MAP(u,LAMBDA(x,SUM(--(b=x)))))),"—")'))
    for c_, w_ in (("A", 2), ("B", 12), ("C", 16), ("D", 16), ("E", 34), ("F", 34), ("G", 10), ("H", 2), ("I", 2), ("J", 12), ("K", 30), ("L", 11)):
        wc.Columns(c_).ColumnWidth = w_
    wc.Range("E5:F2000").WrapText = True
    # подсветка в «Линиях»: изменённые ячейки (видимые столбцы) — оранжевые, новые линии — зелёные
    wsl_ = wb.Worksheets("Линии")
    lol_ = wsl_.ListObjects("Линии")
    r0 = lol_.DataBodyRange.Row
    last = r0 + lol_.ListRows.Count + 300
    ll = LETTER["line"]
    key = '$A%d&"|"&$%s%d' % (r0, ll, r0)
    num = {"Р, кВт", "Длина, м"}
    for j, h in enumerate(SNAP_FIELDS[2:13], start=3):
        try:
            ci = lol_.ListColumns(h).Index
        except Exception:
            continue
        if lol_.HeaderRowRange.Cells(1, ci).EntireColumn.Hidden or wsl_.Columns(ci).OutlineLevel > 1:
            continue
        cl = letters(ci)
        cur = "%s%d" % (cl, r0)
        old = "ИНДЕКС(Выдача_%d;ПОИСКПОЗ(%s;Выдача_ключ;0))" % (j, key.replace(",", ";"))
        cmp_ = ("ОКРУГЛ(Ч(%s);3)<>ОКРУГЛ(Ч(%s);3)" % (old, cur)) if h in num else ('%s&""<>%s&""' % (old, cur))
        fml = '=И(Выдача_n>0;$%s%d<>"";ЕСЛИОШИБКА(%s;ЛОЖЬ))' % (ll, r0, cmp_)
        try:
            fc = wsl_.Range("%s%d:%s%d" % (cl, r0, cl, last)).FormatConditions.Add(2, None, fml)
            fc.Interior.Color = 14083324
            fc.SetFirstPriority()
        except Exception as e:
            log("подсветка изменений не задана:", h, e)
    try:
        fml = '=И(Выдача_n>0;$%s%d<>"";ЕНД(ПОИСКПОЗ(%s;Выдача_ключ;0)))' % (ll, r0, key.replace(",", ";"))
        fc = wsl_.Range("%s%d:%s%d" % (ll, r0, ll, last)).FormatConditions.Add(2, None, fml)
        fc.Interior.Color = 14348258
        fc.SetFirstPriority()
    except Exception as e:
        log("подсветка новых линий не задана:", e)
    return wc


def build_labels_sheet(wb):
    """лист «Наклейки»: заполняет макрос SX_Labels (кнопка на листе «Щит») для щита в «Щит»!C3."""
    ws = wb.Worksheets.Add(After=wb.Worksheets("Изменения"))
    ws.Name = "Наклейки"
    ws.Range("A1").Value = "НАКЛЕЙКИ — кнопка «Наклейки (A4)» на листе «Щит» (щит в ячейке C3)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Модулей на рейку:"
    ws.Range("A2").Font.Size = 10
    ws.Range("A2").HorizontalAlignment = -4152
    ws.Range("A2:B2").Merge()
    ws.Range("C2").Value = 24
    ws.Range("C2").Interior.Color = C_IN
    ws.Range("D2").Value = ("Ячейка = 1 модуль DIN, 18 мм на бумаге (ширина подгоняется под принтер при каждом нажатии кнопки). "
                            "Печать: A4 альбомная, масштаб 100 %, нарезать по линиям; длинная рейка — несколькими полосами. "
                            "Ширина аппарата — «Каталог», столбец L.")
    ws.Range("D2:O2").Merge()                            # подсказка — блоком с переносом (не печатается отдельной страницей)
    ws.Range("D2").WrapText = True
    ws.Range("D2").VerticalAlignment = -4108
    ws.Range("D2").Font.Size = 8
    ws.Range("D2").Font.Color = 8421504
    ws.Rows(2).RowHeight = 26
    return ws


# ------------------------------------------------------------------------------------------------
# v5: значения по умолчанию в «Кодах» для шаблона и примера (--defaults): марка, кВт на шт, управление
V5_CODE_DEFAULTS = {
    "R": {"kw": 0.06}, "L": {}, "LED": {"brand": "мс"}, "LD": {"brand": "мс"}, "SH": {"brand": "мс", "ctl": "шторы"},
    "ТП": {"brand": "мс"}, "K": {"brand": "мс"}, "К": {"brand": "мс"}, "ПВ": {"brand": "мс"}, "В": {"brand": "мс"},
    "УВ": {"brand": "мс"}, "ПН": {"brand": "мс"}, "КМН": {"brand": "мс"}, "ПУ": {"brand": "мс"}, "ЭK": {"brand": "мс", "ctl": "клапан"},
    "ЭК": {"brand": "мс", "ctl": "клапан"}, "R.ШТК": {"brand": "мс"}, "R.ЩОПС": {"brand": "мс"}, "R.GSM": {"brand": "мс"},
    "АСУ": {"brand": "мс"}, "PS-KNX": {"brand": "мс"}, "АПС": {"brand": "мс"}, "DALI": {"brand": "мс", "ctl": "DALI"},
    "КУП": {"brand": "ре"}, "PE": {"brand": "ре"}, "НВК": {"brand": "мс"},
}
CODE_HEADS = ["Код", "Наименование", "Опуск, м", "1 - этаж/2 - шлейф", "Пуск: 1 нет/2 двигатель/3 свет", "Марка кабеля",
              "Жил", "Напряжение", "Фаз", "Сечение", "Вид нагрузки", "Группа Кс", "Пожарная",
              "кВт на шт", "Характеристика (B/C/D)", "Вид аппарата утечки", "Управление", "Аппараты",
              "УГО", "Дифзащита, мА (да = 30)", "Мин. автомат, А", "3L клемм на линию"]


def prepare_rows(data):
    """строки таблиц из JSON extract_v3 (v5: маркер группы «гр», один столбец аппаратов, листы вводного узла)."""
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
    names = set(board_file)
    for p in data["places"]:
        ln = p.get("line")
        if ln in names and ln not in parent and board_file[ln] != p["file"]:
            parent[ln] = p["board"]
            feeds[ln] = (p["file"], ln)
            notes.append(f"щит {ln} подключён к {p['board']} по номеру линии")
    feed_of = {v: k for k, v in feeds.items()}
    order = sorted(board_file, key=lambda b: (0 if not parent.get(b) else 1, b))
    if not order:
        order = ["ЩР"]
    per_sheet = int((data.get("settings") or {}).get("places_per_sheet") or 17)
    bmeta = {b["name"]: b for b in data["boards"]}
    rows, boards = [], []
    for b in order:
        m = bmeta.get(b, {})
        pl = [p for p in data["places"] if p["board"] == b]
        pl.sort(key=lambda p: (p.get("pos") is None, p.get("pos") or 0))
        drawn = [p.get("pos") for p in pl if p.get("mode") != "вручную" and isinstance(p.get("pos"), (int, float))
                 and not str(p.get("line") or "").startswith(("Ввод", "На ИБП", "От ИБП"))]     # вводы и ИБП — часть узла
        nin_old = int(m.get("nin", int(min(drawn)) - 1 if drawn else per_sheet))     # мест под вводной узел
        in_grp = False
        for p in pl:
            line = p.get("line")
            r = {"board": b, "line": line}
            manual = p.get("mode") == "вручную"
            pos = p.get("pos")
            if p.get("unplaced"):
                r["pos"] = "-"
            elif isinstance(pos, (int, float)):
                r["pos"] = int(pos)
            ln = lines.get((p["file"], b, line)) or lines.get((p["file"], None, line))
            qf_man = ln.get("qf") if ln else None
            qt = p.get("qf_txt")
            if qt not in (None, ""):                         # свой автомат (первая строка группы)
                in_grp = not manual
                if str(qt).strip() not in ("авто", "+"):
                    r["qf_in"] = qt
                elif qf_man not in (None, ""):
                    r["qf_in"] = qf_man
            elif line:
                if manual or p.get("unplaced") or not in_grp:
                    if qf_man not in (None, ""):
                        r["qf_in"] = qf_man
                else:
                    r["grp_mark"] = "гр"                     # на автомате строки выше
            if p.get("grp") == "узо":
                r["grp_mark"] = "узо"                         # свой автомат на шине УЗО строки выше
            if manual:
                in_grp = False
            if p.get("rcd") is not None:
                r["rcd_in"] = p["rcd"]
            devs = [str(p.get(k)).strip() for k in ("dev1", "dev2", "dev3") if p.get(k) not in (None, "")]
            if devs:
                r["dev_in"] = "; ".join(devs)
            for k_src, k_dst in (("bus", "bus"), ("module", "module"), ("ch", "ch_in"), ("xt", "xt_in"), ("mode", "mode")):
                if p.get(k_src) is not None:
                    r[k_dst] = p[k_src]
            if p.get("phase") is not None:                  # только ручная фаза; остальное — автобалансировка
                r["phase_in"] = p["phase"]
            if ln:
                for k_src, k_dst in (("name", "name_in"), ("p", "p_in"), ("u", "u_in"), ("n", "n"), ("ph", "ph_in"), ("l_plan", "lplan_in"),
                                     ("l_tray", "tray"), ("l_pvc", "pvc_in"), ("l_pnd", "pnd_in"), ("cable", "brand_in"), ("cores", "cores_in"),
                                     ("s", "s_in"), ("floor", "floor"), ("dwg_pvc", "dwg_pvc"), ("dwg_pnd", "dwg_pnd"),
                                     ("dwg_handles", "dwg_h"), ("dwg_note", "dwg_note"), ("dwg_label", "dwg_label"), ("kc_group", "grp_in"),
                                     ("kc", "kc_in"), ("section", "sec_in"), ("cat_pz", "catpz_in"), ("k_lay", "klay_in"), ("cat", "cat_in"),
                                     ("fire", "fire_in"), ("cos", "cos_in"), ("note", "note")):
                    if ln.get(k_src) is not None:
                        r[k_dst] = ln[k_src]
                if isinstance(r.get("p_in"), str):               # «кВт» и прочий текст вместо числа — не данные
                    try:
                        r["p_in"] = float(r["p_in"].replace(",", ".").strip())
                    except ValueError:
                        r.pop("p_in")
                if str(line).startswith(("Ввод", "На ИБП", "От ИБП")) and r.get("p_in") in (0, 0.0):
                    r.pop("p_in")                                # ручной ноль у ввода — не данные: серое = нагрузка ввода
                r["_ref"] = ln.get("_ref")
            child = feed_of.get((p["file"], line)) or p.get("feeds")
            if child and child != line:
                r["feeds"] = child
            rows.append(r)
        sch, man = m.get("scheme", {}), m.get("manual", {})
        boards.append({"Щит": b, "Питается от": None, "Мест под вводной узел": nin_old, "Вводов": sch.get("inputs"), "АВР": sch.get("avr"),
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


def validation(rng, items, sep):
    try:
        v = rng.Validation
        v.Delete()
        v.Add(3, 1, 1, items if items.startswith("=") else sep.join(items.split("|")))
        v.ShowError = False
        v.IgnoreBlank = True
    except Exception as e:
        log("список не задан:", rng.Address, e)


# ------------------------------------------------------------------------------------------------
# макросы книги (ASCII; русский текст — ^XXXX, расшифровка U())
def enc(s):
    return "".join(ch if ord(ch) < 128 and ch not in "^~" else "^%04X" % ord(ch) for ch in s)


def uq(s):
    return 'U("%s")' % enc(s)


VBA_STD = r'''Option Explicit
'==========================================================================
' SX v5 - service macros of the Odnolineyka v5 workbook
'   SX_RepairFormulas - re-apply formulas of the service and result columns
'                       of table "Lines" (after rows are inserted / deleted / sorted)
'   SX_FillAuto       - emptied "auto" (gray) cells get their formula back
'   SX_AutoFillOff / SX_AutoFillRestore - while this workbook is active,
'       Excel option "fill formulas in tables" is off: a formula typed into
'       one cell does not spread to the whole column.
' Russian text is encoded (^XXXX) and decoded by U().
'==========================================================================
Private PrevAutoFill As Variant

Public Function U(ByVal s As String) As String
    Dim i As Long, n As Long, ch As String, r As String
    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "^" Then
            r = r & ChrW(CLng("&H" & Mid$(s, i + 1, 4)))
            i = i + 5
        Else
            r = r & ch
            i = i + 1
        End If
    Loop
    U = r
End Function

Public Function SxTable() As ListObject
    On Error Resume Next
    Set SxTable = ThisWorkbook.Worksheets(@LINES@).ListObjects(@LINES@)
End Function

Public Function SxTpl() As Worksheet
    On Error Resume Next
    Set SxTpl = ThisWorkbook.Worksheets(@TPL@)
End Function

Public Sub SX_AutoFillOff()
    On Error Resume Next
    If IsEmpty(PrevAutoFill) Then PrevAutoFill = Application.AutoCorrect.AutoFillFormulasInLists
    Application.AutoCorrect.AutoFillFormulasInLists = False
End Sub

Public Sub SX_AutoFillRestore()
    On Error Resume Next
    If Not IsEmpty(PrevAutoFill) Then
        Application.AutoCorrect.AutoFillFormulasInLists = PrevAutoFill
        PrevAutoFill = Empty
    End If
End Sub

Public Sub SX_FillAuto(ByVal rng As Range)
    Dim lo As ListObject, ws As Worksheet, c As Range, r As Variant, h As String, k As String
    Set lo = SxTable()
    Set ws = SxTpl()
    If lo Is Nothing Or ws Is Nothing Then Exit Sub
    For Each c In rng.Cells
        If Not c.HasFormula Then
            If Len(CStr(c.Formula)) = 0 Then
                h = CStr(lo.HeaderRowRange.Cells(1, c.Column - lo.Range.Column + 1).Value)
                r = Application.Match(h, ws.Columns(1), 0)
                If Not IsError(r) Then
                    k = CStr(ws.Cells(r, 3).Value)
                    If k = "a" Or k = "xa" Then c.Formula2R1C1 = CStr(ws.Cells(r, 2).Value)
                End If
            End If
        End If
    Next
End Sub

Public Sub SX_RepairFormulas()
    Dim lo As ListObject, ws As Worksheet, r As Long, k As String, f As String, calc As Long, bad As Long
    Set lo = SxTable()
    Set ws = SxTpl()
    If lo Is Nothing Or ws Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = -4135
    On Error Resume Next
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        k = CStr(ws.Cells(r, 3).Value)
        f = CStr(ws.Cells(r, 2).Value)
        If f <> "" Then
            If k = "c" Or k = "h" Then
                lo.ListColumns(CStr(ws.Cells(r, 1).Value)).DataBodyRange.Formula2R1C1 = f
                If Err.Number <> 0 Then
                    bad = bad + 1
                    Err.Clear
                End If
            ElseIf k = "a" Or k = "xa" Then
                SX_FillAuto lo.ListColumns(CStr(ws.Cells(r, 1).Value)).DataBodyRange
                Err.Clear
            End If
        End If
    Next
    On Error GoTo 0
    Application.Calculation = calc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If bad > 0 Then
        MsgBox @REPAIR_BAD@ & bad, vbExclamation
    Else
        Application.StatusBar = @REPAIR_OK@
    End If
End Sub

Private Function SxSheetName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", "?", "*", "[", "]", ":")
        s = Replace(s, ch, "_")
    Next
    If Len(s) > 31 Then s = Left$(s, 31)
    SxSheetName = s
End Function

Private Function SxSpill(ByVal c As Range) As Range
    If c.HasSpill Then
        Set SxSpill = c.SpillingToRange
    Else
        Set SxSpill = c
    End If
End Function

Private Sub SxCopyBlock(ByVal src As Range, ByVal ex As Workbook, ByVal nm As String)
    Dim ws As Worksheet
    Set ws = ex.Worksheets.Add(After:=ex.Worksheets(ex.Worksheets.Count))
    ws.Name = SxSheetName(nm)
    src.Copy
    ws.Range("A1").PasteSpecial -4163
    ws.Range("A1").PasteSpecial -4122
    ws.Range("A1").PasteSpecial 8
    Application.CutCopyMode = False
End Sub

' light file for AutoCAD Data Link: per board - table under the scheme, totals per input, cable journal, specification
Public Sub SX_ExportAutoCAD()
    Dim wb As Workbook, ex As Workbook, lo As ListObject, c As Range, sp As Range, b As String, p As String, msg As String
    Dim wsS As Worksheet, wsJ As Worksheet, wsP As Worksheet, n As Long
    Dim oS As Variant, oJ As Variant, oP As Variant
    Set wb = ThisWorkbook
    On Error GoTo fail
    Set wsS = wb.Worksheets(@SH_BOARD@)
    Set wsJ = wb.Worksheets(@SH_JOURNAL@)
    Set wsP = wb.Worksheets(@SH_SPEC@)
    Set lo = wb.Worksheets(@SHIELDS@).ListObjects(@SHIELDS@)
    p = wb.Path & Application.PathSeparator & Left$(wb.Name, InStrRev(wb.Name, ".") - 1) & @EXP_SUFFIX@
    oS = wsS.Range("C3").Value
    oJ = wsJ.Range("E1").Value
    oP = wsP.Range("D1").Value
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set ex = Workbooks.Add(-4167)
    For Each c In lo.ListColumns(@COL_BOARD@).DataBodyRange.Cells
        b = Trim$(CStr(c.Value))
        If Len(b) > 0 Then
            wsS.Range("C3").Value = b
            wsJ.Range("E1").Value = b
            wsP.Range("D1").Value = b
            Application.Calculate
            Set sp = SxSpill(wsS.Range("B7"))
            SxCopyBlock wsS.Range(wsS.Range("B6"), wsS.Cells(sp.Row + sp.Rows.Count - 1, 10)), ex, b & @SFX_SCHEME@
            Set sp = SxSpill(wsS.Range("M6"))
            SxCopyBlock wsS.Range(wsS.Range("L6"), wsS.Cells(15, sp.Column + sp.Columns.Count - 1)), ex, b & @SFX_TOTALS@
            Set sp = SxSpill(wsJ.Range("A5"))
            SxCopyBlock wsJ.Range(wsJ.Range("A2"), wsJ.Cells(sp.Row + sp.Rows.Count - 1, 13)), ex, b & @SFX_JOURNAL@
            Set sp = SxSpill(wsP.Range("B3"))
            SxCopyBlock wsP.Range(wsP.Range("B2"), wsP.Cells(sp.Row + sp.Rows.Count - 1, 10)), ex, b & @SFX_SPEC@
            n = n + 1
        End If
    Next
    If n > 0 Then ex.Worksheets(1).Delete
    wsS.Range("C3").Value = oS
    wsJ.Range("E1").Value = oJ
    wsP.Range("D1").Value = oP
    Application.Calculate
    ex.SaveAs p, 51
    ex.Close False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Application.Visible Then
        MsgBox @EXP_DONE@ & n & vbCrLf & p, vbInformation
    Else
        Application.StatusBar = @EXP_DONE@ & n & " | " & p
    End If
    Exit Sub
fail:
    msg = Err.Description
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error Resume Next
    wsS.Range("C3").Value = oS
    wsJ.Range("E1").Value = oJ
    wsP.Range("D1").Value = oP
    If Not ex Is Nothing Then ex.Close False
    If Application.Visible Then
        MsgBox @EXP_FAIL@ & msg, vbExclamation
    Else
        Application.StatusBar = @EXP_FAIL@ & msg
    End If
End Sub

' ===== point 4: sort board, freeze phases, snapshot of the issue, labels
Private Function SxCol(ByVal lo As ListObject, ByVal h As String) As Long
    On Error Resume Next
    SxCol = 0
    SxCol = lo.ListColumns(h).Index
    Err.Clear
End Function

Private Function SxBoardOf(ByVal lo As ListObject) As String
    Dim c As Range, r As Long, b As String
    On Error Resume Next
    Set c = ActiveCell
    If Not c Is Nothing Then
        If c.Worksheet.Name = lo.Parent.Name Then
            If Not Application.Intersect(c, lo.DataBodyRange) Is Nothing Then
                r = c.Row - lo.DataBodyRange.Row + 1
                b = CStr(lo.DataBodyRange.Cells(r, SxCol(lo, @C_BOARD@)).Value)
            End If
        End If
    End If
    On Error GoTo 0
    If b = "" And Application.Visible Then b = InputBox(@ASK_BOARD@, @TITLE@)
    SxBoardOf = Trim$(b)
End Function

Private Function SxCtlOrder(ByVal c As String) As String
    c = LCase$(Trim$(c))
    If c = "relay" Then
        SxCtlOrder = "1"
    ElseIf c = "dim" Then
        SxCtlOrder = "2"
    ElseIf c = "dali" Then
        SxCtlOrder = "3"
    ElseIf c = LCase$(@L_SHTORY@) Then
        SxCtlOrder = "4"
    ElseIf c = LCase$(@L_KLAPAN@) Then
        SxCtlOrder = "5"
    ElseIf c = "" Then
        SxCtlOrder = "9"
    Else
        SxCtlOrder = "6"
    End If
End Function

' rows of the board: input node keeps its order; then by control, code, room, line. Groups "gr", RCD blocks and manual modules stay together.
Public Sub SX_SortBoard(Optional ByVal board As String = "")
    Dim lo As ListObject, v As Variant, n As Long, i As Long, j As Long, k As Long, nu As Long, first As Long, t As String
    Dim cB As Long, cL As Long, cMode As Long, cSt As Long, cUz As Long, cMst As Long, cZone As Long, cCtl As Long, cCode As Long, cRoom As Long, cKey As Long
    Dim unitOf() As Long, uKey() As String, ord() As Long, rk() As Long, keys() As Variant, posIn() As Long
    Dim isNode As Boolean, isCont As Boolean, prevNode As Boolean, ln As String
    Set lo = SxTable()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If board = "" Then board = SxBoardOf(lo)
    If board = "" Then Exit Sub
    cB = SxCol(lo, @C_BOARD@): cL = SxCol(lo, @C_LINE@): cMode = SxCol(lo, @C_MODE@): cSt = SxCol(lo, @C_START@)
    cUz = SxCol(lo, @C_UZ@): cMst = SxCol(lo, @C_MST@): cZone = SxCol(lo, @C_ZONE@): cCtl = SxCol(lo, @C_CTL@)
    cCode = SxCol(lo, @C_CODE@): cRoom = SxCol(lo, @C_ROOM@): cKey = SxCol(lo, @C_SORTKEY@)
    If cKey = 0 Or cB = 0 Then Exit Sub
    v = lo.DataBodyRange.Value
    n = UBound(v, 1)
    ReDim unitOf(1 To n): ReDim uKey(1 To n): ReDim posIn(1 To n)
    For i = 1 To n
        If CStr(v(i, cB)) = board Then
            If first = 0 Then first = i
            ln = Trim$(CStr(v(i, cL)))
            isNode = (CStr(v(i, cMode)) = @MANUAL@) Or (CStr(v(i, cZone)) = @NODEZONE@)
            isCont = False
            If nu > 0 And i > first Then
                If ln = "" Then
                    isCont = True
                ElseIf Not isNode And Not prevNode Then
                    If Val(CStr(v(i, cSt))) = 0 Then isCont = True
                    If Val(CStr(v(i, cUz))) = 1 Then isCont = True
                    If CStr(v(i, cMst)) = "M" And CStr(v(i - 1, cMst)) = "M" Then isCont = True
                End If
            End If
            If isCont Then
                unitOf(i) = nu
            Else
                nu = nu + 1
                unitOf(i) = nu
                If isNode Then
                    uKey(nu) = "0" & Chr$(1) & Format$(i, "000000")
                Else
                    uKey(nu) = "1" & Chr$(1) & SxCtlOrder(CStr(v(i, cCtl))) & Chr$(1) & UCase$(CStr(v(i, cCode))) & Chr$(1) & UCase$(CStr(v(i, cRoom))) & Chr$(1) & UCase$(ln) & Chr$(1) & Format$(i, "000000")
                End If
            End If
            If ln <> "" Or isNode Then prevNode = isNode
        End If
    Next
    If nu < 2 Then Exit Sub
    ReDim ord(1 To nu)
    For i = 1 To nu: ord(i) = i: Next
    For i = 2 To nu
        k = ord(i): t = uKey(k): j = i - 1
        Do While j >= 1
            If uKey(ord(j)) <= t Then Exit Do
            ord(j + 1) = ord(j): j = j - 1
        Loop
        ord(j + 1) = k
    Next
    ReDim rk(1 To nu)
    For i = 1 To nu: rk(ord(i)) = i: Next
    ReDim keys(1 To n, 1 To 1)
    For i = 1 To n
        If CStr(v(i, cB)) = board Then
            posIn(unitOf(i)) = posIn(unitOf(i)) + 1
            keys(i, 1) = first + (rk(unitOf(i)) * 1000# + posIn(unitOf(i))) / 10000000#
        Else
            keys(i, 1) = CDbl(i)
        End If
    Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    lo.ListColumns(cKey).DataBodyRange.Value = keys
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns(cKey).DataBodyRange, SortOn:=0, Order:=1
        .Header = 1
        .Apply
    End With
    lo.ListColumns(cKey).DataBodyRange.ClearContents
    Application.EnableEvents = True
    SX_RepairFormulas
    Application.ScreenUpdating = True
    Application.StatusBar = @SORT_DONE@ & board
End Sub

' gray (auto) phases of the board become blue (manual): they no longer change after edits
Public Sub SX_FreezePhases(Optional ByVal board As String = "")
    Dim lo As ListObject, cB As Long, cP As Long, i As Long, c As Range, n As Long
    Set lo = SxTable()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If board = "" Then board = SxBoardOf(lo)
    If board = "" Then Exit Sub
    cB = SxCol(lo, @C_BOARD@): cP = SxCol(lo, @C_PHASE@)
    If cP = 0 Then Exit Sub
    Application.EnableEvents = False
    For i = 1 To lo.ListRows.Count
        If CStr(lo.DataBodyRange.Cells(i, cB).Value) = board Or board = "*" Then
            Set c = lo.DataBodyRange.Cells(i, cP)
            If c.HasFormula Then
                If CStr(c.Value) <> "" Then
                    c.Value = CStr(c.Value)
                    n = n + 1
                End If
            End If
        End If
    Next
    Application.EnableEvents = True
    Application.StatusBar = @PH_DONE@ & n
End Sub

Private Function SxSnapSheet() As Worksheet
    On Error Resume Next
    Set SxSnapSheet = ThisWorkbook.Worksheets(@SNAP_SHEET@)
End Function

' snapshot of the issue: key fields of every line -> hidden sheet; sheet "Changes" compares with it
Public Sub SX_Snapshot(Optional ByVal rev As String = "")
    Dim lo As ListObject, ws As Worksheet, v As Variant, i As Long, j As Long, n As Long, m As Long, outA() As Variant, keysA() As Variant, cols As Variant, idx() As Long
    Set lo = SxTable()
    Set ws = SxSnapSheet()
    If lo Is Nothing Or ws Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If rev = "" Then
        If Application.Visible Then
            rev = InputBox(@ASK_REV@, @TITLE@, IIf(CStr(ws.Range("B1").Value) = "", "0", CStr(Val(ws.Range("B1").Value) + 1)))
        Else
            rev = IIf(CStr(ws.Range("B1").Value) = "", "0", CStr(Val(ws.Range("B1").Value) + 1))
        End If
    End If
    If Trim$(rev) = "" Then Exit Sub
    If Application.Visible Then
        If MsgBox(@ASK_FREEZE@, vbYesNo + vbQuestion, @TITLE@) = vbYes Then SX_FreezePhases "*"
    End If
    cols = Array(@F_BOARD@, @F_LINE@, @F_NAME@, @F_P@, @F_QF@, @F_S@, @F_LEN@, @F_PHASE@, @F_CABLE@, @F_MOD@, @F_CH@, @F_XT@, @F_PLACE@, @F_SHEET@)
    m = UBound(cols) + 1
    ReDim idx(1 To m)
    For j = 1 To m: idx(j) = SxCol(lo, CStr(cols(j - 1))): Next
    v = lo.DataBodyRange.Value
    ReDim outA(1 To UBound(v, 1), 1 To m)
    ReDim keysA(1 To UBound(v, 1), 1 To 1)
    For i = 1 To UBound(v, 1)
        If Trim$(CStr(v(i, idx(2)))) <> "" Then
            n = n + 1
            keysA(n, 1) = CStr(v(i, idx(1))) & "|" & CStr(v(i, idx(2)))
            For j = 1 To m
                If idx(j) > 0 Then
                    If IsError(v(i, idx(j))) Then outA(n, j) = "" Else outA(n, j) = v(i, idx(j))
                End If
            Next
        End If
    Next
    Application.EnableEvents = False
    ws.Range(ws.Cells(5, 1), ws.Cells(20004, 15)).ClearContents
    If n > 0 Then
        ws.Range(ws.Cells(5, 1), ws.Cells(4 + n, m)).Value = outA
        ws.Range(ws.Cells(5, 15), ws.Cells(4 + n, 15)).Value = keysA
    End If
    ws.Range("B1").Value = rev
    ws.Range("B2").Value = Date
    ws.Cells(ws.Rows.Count, 16).End(-4162).Offset(1, 0).Value = rev & " | " & Format$(Date, "dd.mm.yyyy") & " | " & n
    Application.EnableEvents = True
    Application.Calculate
    If Application.Visible Then MsgBox @SNAP_DONE@ & n, vbInformation
End Sub

Private Function SxModWidth(ByVal model As String, ByVal dflt As Double) As Double
    Dim r As Variant, w As Variant
    SxModWidth = dflt
    If Trim$(model) = "" Then Exit Function
    On Error Resume Next
    r = Application.Match(model, ThisWorkbook.Worksheets(@CAT_SHEET@).Range("C4:C500"), 0)
    If Not IsError(r) Then
        w = ThisWorkbook.Worksheets(@CAT_SHEET@).Range("L4:L500").Cells(r, 1).Value
        If IsNumeric(w) Then If w > 0 Then SxModWidth = w
    End If
End Function

Private Function SxPoles(ByVal s As String) As Double
    s = UCase$(s)
    If InStr(s, "4P") > 0 Then
        SxPoles = 4
    ElseIf InStr(s, "3P") > 0 Then
        SxPoles = 3
    ElseIf InStr(s, "2P") > 0 Or InStr(s, "1P+N") > 0 Then
        SxPoles = 2
    Else
        SxPoles = 1
    End If
End Function

' caption row above a label strip (rail number)
Private Sub SxLblCap(ByVal ws As Worksheet, ByVal r As Long, ByVal s As String)
    ws.Cells(r, 1).Value = s
    ws.Cells(r, 1).Font.Size = 6.5
    ws.Cells(r, 1).Font.Color = 8421504
    ws.Rows(r).RowHeight = 12
End Sub

' font for a label: 7 pt; 6 pt when the text would take more than 4 lines, 5.5 pt - more than 5 (about 13 characters per module and line)
Private Function SxLblFont(ByVal s As String, ByVal w As Double) As Double
    Dim p As Variant, i As Long, n As Long, per As Double
    per = 13 * IIf(w < 1, 1, w)
    p = Split(s, vbLf)
    For i = 0 To UBound(p)
        n = n + IIf(Len(CStr(p(i))) = 0, 1, -Int(-Len(CStr(p(i))) / per))
    Next
    If n > 5 Then
        SxLblFont = 5.5
    ElseIf n > 4 Then
        SxLblFont = 6
    Else
        SxLblFont = 7
    End If
End Function

' labels for the board: strip on A4, one cell = 1 DIN module (18 mm on paper); designation + line + name under every device.
' Excel prints columns wider than on screen (printer font metrics): the factor F = printed / screen width is taken
' on one wide column (page layout view Width = printed width; rounding to screen pixels is negligible there);
' widths go in screen-pixel steps, so the strip mixes the two nearest screen widths around 18 mm / F.
' A rail longer than the sheet goes on several strips; a caption above each strip names the rail.
Public Sub SX_Labels(Optional ByVal board As String = "")
    Dim lo As ListObject, ws As Worksheet, v As Variant, i As Long, j As Long, n As Long, k As Long, perRail As Double
    Dim cB As Long, cL As Long, cPos As Long, cSt As Long, cQF As Long, cH As Long, cD(1 To 3) As Long, cMS As Long, cMM As Long, cNm As Long, cMode As Long
    Dim items() As Variant, ni As Long, w As Double, col As Long, row As Long, used As Double, txt As String, cw As Double, a As String
    Dim perLine As Double, lu As Double, rail As Long, tgt As Double, x As Double, cwLo As Double, cwHi As Double
    Dim wLo As Double, wHi As Double, acc As Double, vw As Long, nCols As Long, c As Long
    Dim fk As Double, tn As Double, wN As Double, wP As Double
    Set lo = SxTable()
    If lo Is Nothing Then Exit Sub
    If board = "" Then board = Trim$(CStr(ThisWorkbook.Worksheets(@SH_BOARD@).Range("C3").Value))
    If board = "" Then Exit Sub
    cB = SxCol(lo, @C_BOARD@): cL = SxCol(lo, @C_LINE@): cPos = SxCol(lo, @C_POS@): cSt = SxCol(lo, @C_START@)
    cQF = SxCol(lo, @C_QF@): cH = SxCol(lo, @C_H@): cMS = SxCol(lo, @C_MODSTART@): cMM = SxCol(lo, @C_MODEL@)
    cNm = SxCol(lo, @C_NAME@): cMode = SxCol(lo, @C_MODE@)
    cD(1) = SxCol(lo, @C_DEV1@): cD(2) = SxCol(lo, @C_DEV2@): cD(3) = SxCol(lo, @C_DEV3@)
    v = lo.DataBodyRange.Value
    ReDim items(1 To 3, 1 To 4 * UBound(v, 1) + 1)
    For i = 1 To UBound(v, 1)
        If CStr(v(i, cB)) = board And IsNumeric(v(i, cPos)) And CStr(v(i, cPos)) <> "" Then
            If Val(CStr(v(i, cSt))) = 1 And CStr(v(i, cH)) <> "" Then
                ni = ni + 1
                items(1, ni) = SxPoles(CStr(v(i, cH)))
                a = CStr(v(i, cQF)): If a = "" Then a = CStr(v(i, cH))
                items(2, ni) = a & vbLf & CStr(v(i, cL)) & vbLf & CStr(v(i, cNm))
            End If
            If Val(CStr(v(i, cSt))) = 1 Then
                For k = 1 To 3
                    If cD(k) > 0 Then
                        If Trim$(CStr(v(i, cD(k)))) <> "" Then
                            ni = ni + 1
                            a = Trim$(CStr(v(i, cD(k))))
                            If Left$(UCase$(a), 3) = @L_UZO@ Then items(1, ni) = SxPoles(a) Else items(1, ni) = SxModWidth(a, 1)
                            items(2, ni) = a & vbLf & CStr(v(i, cL))
                        End If
                    End If
                Next
            End If
            If cMS > 0 Then
                If CStr(v(i, cMS)) <> "" Then
                    ni = ni + 1
                    items(1, ni) = SxModWidth(CStr(v(i, cMM)), 4)
                    items(2, ni) = CStr(v(i, cMS)) & vbLf & CStr(v(i, cMM))
                End If
            End If
        End If
    Next
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(@LBL_SHEET@)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    perRail = Val(CStr(ws.Range("C2").Value))
    If perRail <= 0 Then perRail = 24
    ws.Range(ws.Cells(3, 1), ws.Cells(400, 80)).UnMerge
    ws.Range(ws.Cells(3, 1), ws.Cells(400, 80)).Clear
    ws.Range("A1").Value = @LBL_TITLE@ & board
    On Error Resume Next
    With ws.PageSetup
        .PaperSize = 9
        .Orientation = 2
        .Zoom = 100
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .CenterHorizontally = False
    End With
    On Error GoTo 0
    tgt = Application.CentimetersToPoints(1.8)
    perLine = Int((Application.CentimetersToPoints(29.7) - ws.PageSetup.LeftMargin - ws.PageSetup.RightMargin) / tgt)
    If perLine < 1 Then perLine = 1
    nCols = perLine
    If perRail < nCols Then nCols = perRail
    ' width: F on a wide column, then screen widths around tgt / F
    vw = 1
    fk = 1
    On Error Resume Next
    ThisWorkbook.Activate
    ws.Activate
    vw = ActiveWindow.View
    ActiveWindow.View = 1
    ws.Columns(1).ColumnWidth = 100
    wN = ws.Columns(1).Width
    ActiveWindow.View = 3
    wP = ws.Columns(1).Width
    ActiveWindow.View = 1
    On Error GoTo 0
    If wN > 0 And wP > 0 Then fk = wP / wN
    If fk < 0.8 Or fk > 1.25 Then fk = 1
    tn = tgt / fk
    x = 5
    Do While x < 15
        ws.Columns(1).ColumnWidth = x
        w = ws.Columns(1).Width
        If w <= tn Then
            cwLo = x: wLo = w
        Else
            cwHi = x: wHi = w
            Exit Do
        End If
        x = x + 0.02
    Loop
    If cwHi = 0 Then cwHi = cwLo: wHi = wLo
    If cwLo = 0 Then cwLo = cwHi: wLo = wHi
    acc = 0
    For c = 1 To nCols
        If Abs(acc + wLo - c * tn) <= Abs(acc + wHi - c * tn) Then
            ws.Columns(c).ColumnWidth = cwLo: acc = acc + wLo
        Else
            ws.Columns(c).ColumnWidth = cwHi: acc = acc + wHi
        End If
    Next
    On Error Resume Next
    ActiveWindow.View = vw
    On Error GoTo 0
    Application.ScreenUpdating = False
    row = 4: rail = 1: used = 0: lu = 0: col = 1
    SxLblCap ws, row, @LBL_RAIL@ & rail
    row = row + 1
    For k = 1 To ni
        w = items(1, k)
        If w < 1 Then w = 1
        If used + w > perRail And used > 0 Then
            rail = rail + 1: used = 0: lu = 0: col = 1
            SxLblCap ws, row + 1, @LBL_RAIL@ & rail
            row = row + 2
        ElseIf lu + w > perLine And lu > 0 Then
            lu = 0: col = 1
            SxLblCap ws, row + 1, @LBL_RAIL@ & rail & @LBL_CONT@
            row = row + 2
        End If
        With ws.Range(ws.Cells(row, col), ws.Cells(row, col + CLng(w) - 1))
            If w > 1 Then .Merge
            .Value = items(2, k)
            .WrapText = True
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4108
            .Font.Size = SxLblFont(CStr(items(2, k)), w)
            .Borders.LineStyle = 1
        End With
        ws.Rows(row).RowHeight = 48
        col = col + CLng(w): used = used + w: lu = lu + w
    Next
    On Error Resume Next
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(row, nCols)).Address
    On Error GoTo 0
    Application.ScreenUpdating = True
    ws.Activate
    Application.StatusBar = @LBL_DONE@ & ni
End Sub
'''

VBA_SHEET = r'''Option Explicit
' table "Lines": gray "auto" cells come back when emptied; "gr" typed into "Breaker" moves to "Group";
' whole rows inserted / deleted -> formulas are repaired
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lo As ListObject, rng As Range, c As Range, h As String, v As String, i As Long
    On Error GoTo done
    Set lo = SxTable()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set rng = Application.Intersect(Target, lo.DataBodyRange)
    If rng Is Nothing Then Exit Sub
    If Target.Columns.Count >= lo.ListColumns.Count Then
        SX_RepairFormulas
        Exit Sub
    End If
    If rng.Cells.Count > 3000 Then Exit Sub
    Application.EnableEvents = False
    For Each c In rng.Cells
        If Not c.HasFormula Then
            h = CStr(lo.HeaderRowRange.Cells(1, c.Column - lo.Range.Column + 1).Value)
            If h = @QF@ Then
                v = Trim$(CStr(c.Formula))
                If v = @GR@ Or v = "=" Or v = "^" Or v = ChrW(8593) Then
                    i = c.Row - lo.DataBodyRange.Row + 1
                    lo.ListColumns(@GRP@).DataBodyRange.Cells(i, 1).Value = @GR@
                    c.ClearContents
                End If
            End If
        End If
    Next
    SX_FillAuto rng
done:
    Application.EnableEvents = True
End Sub
'''

VBA_SHIELDS = r'''Option Explicit
' table "Shields": gray "auto" cells come back when emptied ("Fed from" = by the feeder line)
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lo As ListObject, rng As Range, c As Range, h As String
    On Error GoTo done
    Set lo = Me.ListObjects(@SHIELDS@)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set rng = Application.Intersect(Target, lo.DataBodyRange)
    If rng Is Nothing Then Exit Sub
    If rng.Cells.Count > 3000 Then Exit Sub
    Application.EnableEvents = False
    For Each c In rng.Cells
        If Not c.HasFormula Then
            If Len(Trim$(CStr(c.Formula))) = 0 Then
                h = CStr(lo.HeaderRowRange.Cells(1, c.Column - lo.Range.Column + 1).Value)
                If h = @FEED@ Then c.Formula2 = "=[@[" & @FEEDAUTO@ & "]]"
            End If
        End If
    Next
done:
    Application.EnableEvents = True
End Sub
'''

VBA_WB = r'''Option Explicit
Private Sub Workbook_Open()
    SX_AutoFillOff
End Sub

Private Sub Workbook_Activate()
    SX_AutoFillOff
End Sub

Private Sub Workbook_Deactivate()
    SX_AutoFillRestore
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    SX_AutoFillRestore
End Sub
'''

VBA_REPL = {"@C_BOARD@": uq("Щит"), "@C_LINE@": uq("Номер линии"), "@C_MODE@": uq("Режим"), "@C_START@": uq("_свой автомат"),
            "@C_UZ@": uq("_на шине УЗО"), "@C_MST@": uq("_блок модуля"), "@C_ZONE@": uq("_зона"), "@C_CTL@": uq("_управление (итог)"),
            "@C_CODE@": uq("Код"), "@C_ROOM@": uq("Помещение"), "@C_SORTKEY@": uq("_ключ упорядочивания"), "@C_PHASE@": uq("Фаза"),
            "@MANUAL@": uq("вручную"), "@NODEZONE@": uq("вводной узел"), "@L_SHTORY@": uq("шторы"), "@L_KLAPAN@": uq("клапан"),
            "@L_UZO@": uq("УЗО"), "@ASK_BOARD@": uq("Щит (имя, как в «Линиях»):"), "@TITLE@": uq("Однолинейка v5"),
            "@SORT_DONE@": uq("Щит упорядочен: "), "@PH_DONE@": uq("Фазы закреплены, строк: "), "@SNAP_SHEET@": uq("_выдача"),
            "@ASK_REV@": uq("Номер изменения (0 — первая выдача):"), "@SNAP_DONE@": uq("Выдача зафиксирована, линий: "),
            "@F_BOARD@": uq("Щит"), "@F_LINE@": uq("Номер линии"), "@F_NAME@": uq("Наименование"), "@F_P@": uq("Р, кВт"),
            "@F_QF@": uq("Автомат"), "@F_S@": uq("Сечение, мм²"), "@F_LEN@": uq("Длина, м"), "@F_PHASE@": uq("Фаза"), "@F_CABLE@": uq("Кабель"),
            "@F_MOD@": uq("_Модуль (схема)"), "@F_CH@": uq("Канал (итог)"), "@F_XT@": uq("Клемма"), "@F_PLACE@": uq("Лист / место"),
            "@F_SHEET@": uq("Лист схемы"), "@CAT_SHEET@": uq("Каталог"), "@C_POS@": uq("_№ места (итог)"), "@C_QF@": uq("_QF гр."),
            "@C_H@": uq("Автомат (схема)"), "@C_MODSTART@": uq("_модуль нач."), "@C_MODEL@": uq("_модель гр."), "@C_NAME@": uq("Наименование (итог)"),
            "@C_DEV1@": uq("Аппарат 1"), "@C_DEV2@": uq("Аппарат 2"), "@C_DEV3@": uq("Аппарат 3"), "@LBL_SHEET@": uq("Наклейки"),
            "@LBL_TITLE@": uq("НАКЛЕЙКИ щита "), "@LBL_RAIL@": uq("Рейка "), "@LBL_CONT@": uq(", продолжение"), "@ASK_FREEZE@": uq("Закрепить фазы всех щитов? Тогда правки после выдачи не переставят фазы."), "@LBL_DONE@": uq("Наклеек: "),
            "@SH_BOARD@": uq("Щит"), "@SH_JOURNAL@": uq("Кабельный_журнал"), "@SH_SPEC@": uq("Спецификация"), "@COL_BOARD@": uq("Щит"),
            "@EXP_SUFFIX@": uq("_AutoCAD.xlsx"), "@SFX_SCHEME@": uq(" схема"), "@SFX_TOTALS@": uq(" итоги"), "@SFX_JOURNAL@": uq(" журнал"),
            "@SFX_SPEC@": uq(" спец"), "@EXP_DONE@": uq("Выгружено щитов: "), "@EXP_FAIL@": uq("Выгрузка не удалась: "),
            "@SHIELDS@": uq("Щиты"), "@FEED@": uq("Питается от"), "@FEEDAUTO@": uq("_авто питается от"), "@LINES@": uq("Линии"), "@TPL@": uq("_формулы"), "@QF@": uq("Автомат"), "@GRP@": uq("Группа"), "@GR@": uq("гр"),
            "@REPAIR_BAD@": uq("Не удалось записать формулы столбцов: "),
            "@REPAIR_OK@": uq("Формулы таблицы «Линии» восстановлены")}


def vba_text(s):
    for k, v in VBA_REPL.items():
        s = s.replace(k, v)
    assert all(ord(ch) < 128 for ch in s), "в модуле VBA не-ASCII символы"
    return s


def add_vba(wb, wsl):
    vbp = wb.VBProject
    std = vbp.VBComponents.Add(1)
    std.Name = "SX_v5"
    std.CodeModule.AddFromString(vba_text(VBA_STD))
    vbp.VBComponents(wsl.CodeName).CodeModule.AddFromString(vba_text(VBA_SHEET))
    vbp.VBComponents(wb.CodeName).CodeModule.AddFromString(vba_text(VBA_WB))
    vbp.VBComponents(wb.Worksheets("Щиты").CodeName).CodeModule.AddFromString(vba_text(VBA_SHIELDS))
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "vba")
    os.makedirs(out_dir, exist_ok=True)
    for nm, txt in (("SX_v5.bas", 'Attribute VB_Name = "SX_v5"\r\n' + vba_text(VBA_STD)), ("Линии_лист.cls.txt", vba_text(VBA_SHEET)),
                    ("ЭтаКнига.cls.txt", vba_text(VBA_WB)), ("Щиты_лист.cls.txt", vba_text(VBA_SHIELDS))):
        with open(os.path.join(out_dir, nm), "w", encoding="ascii", newline="\r\n") as fh:
            fh.write(txt)


def add_button(ws, left, top, width, height, text, macro):
    try:
        shp = ws.Shapes.AddFormControl(0, left, top, width, height)      # 0 = кнопка (элемент управления формы)
        shp.OnAction = macro
        try:
            shp.OLEFormat.Object.Caption = text
        except Exception:
            shp.TextFrame.Characters().Text = text
        shp.Placement = 3                                                  # не двигать с ячейками
        return shp
    except Exception as e:
        log("кнопка не создана:", text, e)


CARD = ('=LET(v,FILTER(Линии[Щит],(Линии[_видна]=1)*(Линии[Щит]<>""),""),u,UNIQUE(v),'
        'IF(OR(ROWS(u)<>1,INDEX(u,1)=""),"Объект: щитов "&ROWS(Щиты[Щит])&", линий "&SUM(--(Линии[Номер линии]<>""))&'
        '", замечаний "&COUNTIF(Линии[Замечания],"?*")&CHAR(10)&"Кнопка щита справа — показать один щит и его итоги",'
        'LET(щ,INDEX(u,1),X,LAMBDA(h,XLOOKUP(щ,Щиты[Щит],h,"")),'
        'щ&IF(X(Щиты[Питается от (итог)])<>"","  ← от "&X(Щиты[Питается от (итог)]),"")&"   ·   листов схемы: "&MAXIFS(Линии[Лист схемы],Линии[Щит],щ)&'
        '"   ·   замечаний: "&COUNTIFS(Линии[Щит],щ,Линии[Замечания],"?*")&CHAR(10)&'
        '"Руст "&FIXED(X(Щиты[Руст, кВт]),1)&" · Рр "&FIXED(X(Щиты[Рр, кВт]),1)&" кВт · Iр "&FIXED(X(Щиты[Iр, А]),1)&" А · cos φ "&'
        'FIXED(X(Щиты[cos φ]),2)&" · ΔU до шин "&IFERROR(FIXED(X(Щиты[ΔU до шин, %]),2),"—")&" %")))')


def add_slicer(wb, wsl, lol, n_boards):
    import pythoncom
    try:
        sc = wb.SlicerCaches.Add2(lol, "Щит")
        ncol = max(1, min(n_boards, 12))
        nrow = -(-max(1, n_boards) // ncol)
        h = max(30, 4 + 21 * nrow)
        wsl.Rows(1).RowHeight = max(48, h + 4)
        left = wsl.Range("D1").Left + 3
        width = 8 + 74 * ncol
        sl = sc.Slicers.Add(wsl, pythoncom.Missing, "Кнопки_щитов", "Щит", 2, left, width, h)
        sl.DisplayHeader = False
        sl.NumberOfColumns = ncol
        sl.Top, sl.Left, sl.Width, sl.Height = 2, left, width, h          # при создании координаты не применяются
        try:
            sl.Style = "SlicerStyleLight6"
        except Exception:
            pass
        try:
            sl.ColumnWidth = 70
            sl.RowHeight = 18
        except Exception:
            pass
        try:
            sc.SortItems = 1
            sc.CrossFilterType = 1
        except Exception:
            pass
        return sl
    except Exception as e:
        log("срез «Щит» не создан:", e)


# ------------------------------------------------------------------------------------------------
# «Щит»: то, что уходит в AutoCAD по выбранному щиту — таблица под схемой (строка на место) и итоги по каждому вводу
BOARD_TABLE = ('=LET(щ,$C$3,m,(Линии[Щит]=щ)*ISNUMBER(Линии[_№ места (итог)]),IF(SUM(m)=0,"в щите нет строк на схеме",'
               'LET(pos,FILTER(Линии[_№ места (итог)],m),n,MAX(pos),k,SEQUENCE(n),'
               'f,LAMBDA(c_,LET(v,XLOOKUP(k,pos,FILTER(c_,m),""),IF(v="","",v))),'
               'HSTACK(k,INT((k-1)/Столбцов_на_лист)+1,f(Линии[Номер линии]),f(Линии[_фаза итог]),f(Линии[_Р схемы]),f(Линии[_Ток схемы]),'
               'f(Линии[Наименование (итог)]),f(Линии[Кабель (строкой)]),f(Линии[_автомат гр.])))))')
BOARD_INPUTS = ('=LET(РАСЧЁТ,{R},щ,$C$3,inp,(Линии[Щит]=щ)*(Линии[_ввод щита]=1),k_,SUM(inp),'
                'nm,IF(k_=0,"Щит "&щ,FILTER(Линии[Номер линии],inp,"")),srcs,IF(k_=0,"*",FILTER(Линии[Источник для режимов],inp,"")),'
                'cats,IF(k_=0,XLOOKUP(щ,Щиты[Щит],Щиты[Категория шины «Ввод»],""),FILTER(Линии[Категория ввода (итог)],inp,"")),'
                'own,(Линии[Щит]=щ)*(Линии[Группа Кс (итог)]<>"Нет")*(Линии[Номер линии]<>""),ph,Линии[_фаза итог],pw,IFERROR(--Линии[_Р линии, кВт],0),'
                'ибп,IF(SUM(inp*(Линии[Источник для режимов]="ИБП"))>0,"",LET(u_,XLOOKUP(щ,Щиты[Щит],Щиты[ИБП от],"")&"",IF(u_="","Ввод 1",u_))),'
                'col,LAMBDA(j,LET(s,INDEX(srcs,j),m,AA3#*IF(s="*",1,((AC3#=s)+(AC3#="ИБП")*(ибп=s))>0),x,РАСЧЁТ(m,0,0),'
                'o,own*IF(s="*",1,((Линии[Источник для режимов]=s)+(Линии[Источник для режимов]="ИБП")*(ибп=s))>0),ф1,SUM(o*pw*((ph="L1,L2,L3")/3+(ph="L1"))),'
                'ф2,SUM(o*pw*((ph="L1,L2,L3")/3+(ph="L2"))),ф3,SUM(o*pw*((ph="L1,L2,L3")/3+(ph="L3"))),фс,ф1+ф2+ф3,рр,INDEX(x,5),'
                'кт,INDEX(cats,j),VSTACK(INDEX(nm,j),IF(кт="I особая","I особой категории",IF(кт="","",кт&" категории")),'
                'INDEX(x,1),рр,INDEX(x,8),INDEX(x,3),INDEX(x,2),IF(фс=0,0,рр*ф1/фс),IF(фс=0,0,рр*ф2/фс),IF(фс=0,0,рр*ф3/фс)))),'
                'IF(ROWS(srcs)=1,col(1),REDUCE(col(1),SEQUENCE(ROWS(srcs)-1,1,2),LAMBDA(acc,j,HSTACK(acc,col(j))))))')


def build_board(wb, root, sep):
    ws = wb.Worksheets.Add(After=wb.Worksheets("Линии"))
    ws.Name = "Щит"
    ws.Range("B1").Value = "ЩИТ: данные под однолинейку — таблица под схемой и итоги по каждому вводу (то, что уходит в AutoCAD)"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    ws.Range("B3").Value = "Щит:"
    ws.Range("B3").Font.Bold = True
    ws.Range("C3").Value = root
    ws.Range("C3").Interior.Color = C_IN
    ws.Range("C3").Font.Bold = True
    validation(ws.Range("C3"), "=Список_щитов", sep)
    F2(ws.Range("E3"), '=LET(щ,$C$3,X,LAMBDA(h,XLOOKUP(щ,Щиты[Щит],h,"")),"Питается от: "&IF(X(Щиты[Питается от (итог)])="","—",X(Щиты[Питается от (итог)]))&'
                       '"   ·   вводов: "&X(Щиты[Вводов])&"   ·   АВР: "&X(Щиты[АВР])&"   ·   мест под вводной узел: "&'
                       'IF(X(Щиты[Мест под вводной узел])="",Столбцов_на_лист,X(Щиты[Мест под вводной узел]))&"   ·   замечаний: "&'
                       'COUNTIFS(Линии[Щит],щ,Линии[Замечания],"?*"))')
    ws.Range("E3").Font.Color = 8421504
    # служебные массивы (скрыты): линии щита и нижних щитов, источник (ввод) в выбранном щите
    ws.Range("AA2:AC2").Value = [["в составе", "нижний щит", "ввод в щите"]]
    F2(ws.Range("AA3"), '=--(ISNUMBER(SEARCH("|"&$C$3&"|",Линии[_путь щита]))*(Линии[Номер линии]<>""))')
    F2(ws.Range("AB3"), '=IFERROR(TEXTBEFORE(TEXTAFTER(Линии[_путь щита],"|"&$C$3&"|"),"|"),"")')
    F2(ws.Range("AC3"), '=IF(AB3#="",Линии[Источник для режимов],XLOOKUP(AB3#,Щиты[Щит],Щиты[Источник фидера],"Ввод 1"))')
    ws.Columns("AA:AC").Hidden = True
    # таблица под схемой
    ws.Range("B5").Value = "ТАБЛИЦА ПОД СХЕМОЙ (строка = место на схеме, шаг 20 мм; в AutoCAD таблица поворачивается)"
    ws.Range("B5").Font.Bold = True
    heads = ["Место", "Лист", "Обозначение на плане", "Фаза", "Рн, кВт", "Iн, А", "Наименование электроприёмника", "Кабель", "Автомат"]
    ws.Range("B6:J6").Value = [heads]
    ws.Range("B6:J6").Font.Bold = True
    ws.Range("B6:J6").Interior.Color = C_HEAD
    ws.Range("B6:J6").WrapText = True
    F2(ws.Range("B7"), BOARD_TABLE)
    ws.Range("F7:G400").NumberFormatLocal = "0,00"
    # итоги по вводам
    ws.Range("L5").Value = "ИТОГИ ПО ВВОДАМ (под схемой)"
    ws.Range("L5").Font.Bold = True
    labels = ["Ввод", "Категория", "Руст, кВт", "Рр, кВт", "Iр, А", "cos φ", "Кс", "L1р, кВт", "L2р, кВт", "L3р, кВт"]
    ws.Range("L6:L15").Value = [[x] for x in labels]
    ws.Range("L6:L15").Font.Bold = True
    ws.Range("L6:L15").Interior.Color = C_HEAD
    F2(ws.Range("M6"), BOARD_INPUTS.replace("{R}", RASCHET))
    ws.Range("M8:T15").NumberFormatLocal = "0,00"
    ws.Range("M6:T7").Font.Bold = True
    ws.Range("L17").Value = ("Ввод = строки с кодом «Ввод» этого щита. Линии ввода — по шине и секции (столбцы «Шина», «Секция» в «Линиях»); "
                             "нижние щиты — по своему фидеру; шина ИБП — к вводу из «ИБП от» на листе «Щиты» (если у щита нет отдельного ввода с шины ИБП). "
                             "Итоги — нормальный режим. При АВР у строки «Ввод» в таблице слева — вся нагрузка щита (один ввод тянет всё): по ней подобраны автомат и кабель ввода. "
                             "Категория — у строки «Ввод» (столбец «Категория ввода» в «Ещё») или по схеме щита.")
    ws.Range("L17:P22").Merge()                          # заметка — компактным блоком под итогами
    ws.Range("L17").WrapText = True
    ws.Range("L17").VerticalAlignment = -4160
    ws.Range("L17").Font.Size = 9
    ws.Range("L17").Font.Color = 8421504
    for c_, w_ in (("A", 2), ("B", 7), ("C", 5), ("D", 15), ("E", 10), ("F", 8), ("G", 8), ("H", 44), ("I", 34), ("J", 18), ("K", 3), ("L", 14)):
        ws.Columns(c_).ColumnWidth = w_
    ws.Range("M:T").ColumnWidth = 15
    return ws


# ------------------------------------------------------------------------------------------------
def build(data, base, out, mode="нормы 2026", defaults=False):
    rows, boards, notes = prepare_rows(data)
    for n_ in notes:
        log("  ", n_)
    tmp = os.path.splitext(out)[0] + "_tmp.xlsx"
    shutil.copyfile(base, tmp)
    with ExcelSession() as xl:
        t0 = time.time()
        wb = xl.Workbooks.Open(os.path.abspath(tmp), 0, False)
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
        for src in (wb.LinkSources(1) or []):
            try:
                wb.BreakLink(src, 1)
            except Exception:
                pass
        for i in range(wb.Names.Count, 0, -1):
            nmo = wb.Names(i)
            try:
                if "#REF" in nmo.RefersTo:
                    nmo.Delete()
            except Exception:
                pass
        log("основа подготовлена: %.0f с" % (time.time() - t0))
        sep = xl.International[4]

        # --- «Объект»
        wo = wb.Worksheets.Add(Before=wb.Worksheets(1))
        wo.Name = "Объект"
        obj = [("Объект", "", "название и адрес"),
               ("ρ меди для Iкз min, Ом·мм²/м", data["kz"].get("rho_cu") or 0.0263, "≈ 1,5·ρ20 (ГОСТ 28249 п. 2.4.2)"),
               ("ρ алюминия для Iкз min, Ом·мм²/м", data["kz"].get("rho_al") or 0.042, "≈ 1,5·ρ20"),
               ("Переходные сопротивления на щит, Ом", data["kz"].get("r_contact") or 0.015, "контакты и аппараты на каждый щит"),
               ("c_min для Iкз min", data["kz"].get("c_min") or 0.95, "МЭК 60909-0"),
               ("Методика расчёта", mode, "v3.12 — как старая книга (для сверки); нормы 2026 — СП 256 с Изм. 9 (см. «Методика»)"),
               ("kп.к — региональный коэффициент", 1, "СП 256 табл. 7.5а (ЦФО 0,81, СЗФО 0,91…); для частного дома по умолчанию 1"),
               ("Наименование линии", "по коду и № помещения", "«по коду и № помещения» — «Розеточная группа, пом. 105»; «по экспликации» — название из листа «Экспликация»"),
               ("Писать «пом.» в наименовании", "да", "да — «…, пом. 105»; нет — «… 105»"),
               ("Марка кабеля для пожарных линий", "мфр", "обозначение из «Коды» → «Кабели» (мфр = ВВГнг(А)-FRLS)"),
               ("Запас мощности БП", 1.3, "линии до 50 В: мощность БП ≥ запас × Σ (Uвых·I) — как в листе «Расчет_БП» v1.05"),
               ("Допустимо ΔU в линиях до 50 В, %", 5, "от выхода БП до конца линии (LED 24 В, 12 В)")]
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
        wo.Columns("C").ColumnWidth = 22
        wo.Columns("D").ColumnWidth = 60
        validation(wo.Range("C8"), "v3.12|нормы 2026", sep)
        validation(wo.Range("C10"), "по коду и № помещения|по экспликации", sep)
        validation(wo.Range("C11"), "да|нет", sep)
        set_names(wb, {"ρкз_Cu": "=Объект!$C$4", "ρкз_Al": "=Объект!$C$5", "R_пер": "=Объект!$C$6", "c_min": "=Объект!$C$7",
                       "Методика": "=Объект!$C$8", "k_пк": "=Объект!$C$9", "Имя_режим": "=Объект!$C$10", "Имя_пом": "=Объект!$C$11",
                       "Марка_пож": "=Объект!$C$12", "Запас_БП": "=Объект!$C$13", "dU_LED": "=Объект!$C$14"})
        # модули по умолчанию для управления строки (как в проекте «Жуковка»): шторы — 2 выхода на линию; DALI — шлюз, линии на одной шине
        wo.Range("I1").Value = "МОДУЛИ ПО УМОЛЧАНИЮ (для «Управления» линии; модель — из «Каталога»)"
        wo.Range("I1").Font.Bold = True
        wo.Range("I2:K2").Value = [["Управление", "Модель", "Каналов на линию"]]
        wo.Range("I3:K7").Value = [["Relay", "ZIOMB24V2", 1], ["DIM", "ZDINDX4", 1], ["DALI", "ZDID64X2", 0], ["шторы", "ZIOMB24V2", 2],
                                   ["клапан", "ZCL-8HT230", 1]]
        lmd = wo.ListObjects.Add(1, wo.Range("I2:K7"), None, 1)
        lmd.Name = "Модули_умолч"
        lmd.TableStyle = "TableStyleLight9"
        wo.Columns("I").ColumnWidth = 14
        wo.Columns("J").ColumnWidth = 16
        wo.Columns("K").ColumnWidth = 18
        validation(lmd.ListColumns("Модель").DataBodyRange, "=Каталог!$AB$4:$AB$100", sep)
        try:
            wo.Range("K2").AddComment("0 — линии не занимают каналы (DALI: группы на одной шине шлюза); 2 — шторы (вверх / вниз). Пусто в «Модели» — модуль не подставляется.")
        except Exception:
            pass
        wo.Range("F2:G2").Value = [["Число силовых ЭП", "Кс (СП 256 табл. 7.5)"]]
        k75 = [(1, 1), (2, 1), (3, 0.9), (5, 0.8), (8, 0.75), (10, 0.7), (15, 0.65), (20, 0.65), (30, 0.6), (50, 0.55), (100, 0.55), (200, 0.5)]
        wo.Range("F3:G%d" % (2 + len(k75))).Value = k75
        wo.Range("F3:G%d" % (2 + len(k75))).Interior.Color = C_IN
        wo.Range("F2:G2").Font.Bold = True
        wo.Columns("F").ColumnWidth = 16
        wo.Columns("G").ColumnWidth = 20
        set_names(wb, {"Кс75_n": "=Объект!$F$3:$F$%d" % (2 + len(k75)), "Кс75_k": "=Объект!$G$3:$G$%d" % (2 + len(k75))})

        # --- «Коды»: типовые потребители (код задаёт всё по умолчанию) + УГО, этажи, кабели
        wk = wb.Worksheets.Add(After=wo)
        wk.Name = "Коды"
        wk.Range("A1").Value = ("КОДЫ = ТИПОВЫЕ ПОТРЕБИТЕЛИ (одна строка на код: марка, жилы, мин. сечение, кВт на шт, характеристика, утечка, "
                                "управление, аппараты, УГО, мин. автомат). Справа — этажи, кабели, экспликация (по желанию). Новая строка — сразу под таблицей.")
        wk.Range("A1").Font.Bold = True
        ref = data.get("ref", {})
        codes = {}
        for k_, v_ in ref.get("codes", {}).items():
            v_ = list(v_) + [None] * (13 - len(v_))
            ex = [None] * 5
            if defaults:
                d_ = V5_CODE_DEFAULTS.get(k_, {})
                if d_.get("brand") and not v_[5]:
                    v_[5] = d_["brand"]
                ex = [d_.get("kw"), None, None, d_.get("ctl"), None]
            codes[k_] = v_[:13] + ex + [None] * 4
        if "УЗО" not in codes:                               # строка УЗО на группу автоматов: без кабеля и нагрузки, УЗО в «Аппаратах»
            codes["УЗО"] = ["УЗО", "Шина УЗО (линии со своими автоматами, «узо» в «Группе»)", 0, None, 1, None, None, 230, 1, None, None, "Нет", "нет",
                            None, None, None, None, "УЗО 2P 40А 30мА", None, None, None, None]
        for k_, u_ in ref.get("ugo", {}).items():           # УГО, ток утечки, мин. автомат, 3L клеммы — в ту же строку кода
            u_ = (list(u_) + [None] * 5)[1:5]
            if k_ not in codes:
                codes[k_] = [k_] + [None] * 17 + [None] * 4
            codes[k_][18:22] = u_
        expl = {i: list(e) for i, e in enumerate(data.get("explication", []))}
        blocks = [("Коды", "A", CODE_HEADS, codes),
                  ("Этажи", "Y", ["Этаж", "Название этажа", "Подъём, м", "По другим этажам, м"], ref.get("floors", {})),
                  ("Кабели", "AD", ["Обозначение", "Кабель", "Сечение в марке (1 да / 2 нет)"], ref.get("cables", {})),
                  ("Экспликация", "AH", ["Этаж", "Помещение", "Название"], expl)]
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
        wk.Columns.ColumnWidth = 11
        wk.Columns("B").ColumnWidth = 28
        lk = wk.ListObjects("Коды")
        validation(lk.ListColumns("Характеристика (B/C/D)").DataBodyRange, "B|C|D", sep)
        validation(lk.ListColumns("Вид аппарата утечки").DataBodyRange, "дифавтомат|УЗО на группу|нет", sep)
        validation(lk.ListColumns("Управление").DataBodyRange, "Relay|DIM|DALI|шторы|клапан", sep)
        set_names(wb, {"Потребители_обозначение": "=Коды[[Код]:[Сечение]]", "Этаж_обозначение": "=Этажи",
                       "Кабель_обозначение": "=Кабели", "Коды_код": "=Коды[Код]", "Коды_группа": "=Коды[Группа Кс]",
                       "Коды_пожарная": "=Коды[Пожарная]", "Коды_вид": "=Коды[Вид нагрузки]", "Коды_имя": "=Коды[Наименование]",
                       "Коды_режим": "=Коды[1 - этаж/2 - шлейф]", "Коды_квт": "=Коды[кВт на шт]", "Коды_хар": "=Коды[Характеристика (B/C/D)]",
                       "Коды_утечка": "=Коды[Вид аппарата утечки]", "Коды_упр": "=Коды[Управление]", "Коды_апп": "=Коды[Аппараты]",
                       "Коды_уго": "=Коды[УГО]", "Коды_мА": "=Коды[Дифзащита, мА (да = 30)]", "Коды_минавт": "=Коды[Мин. автомат, А]",
                       "Коды_3L": "=Коды[3L клемм на линию]"})
        try:
            wb.Names("УГО_спр").Delete()
        except Exception:
            pass
        log("коды: %d, этажей: %d, кабелей: %d, помещений: %d" % tuple(len(b[3]) for b in blocks))
        wk.ListObjects("Экспликация").HeaderRowRange.Cells(1, 1).AddComment("Экспликация (по желанию): этаж и помещение — как в номере линии (R.1.05.3 → этаж 1, помещение 05). Включается на «Объекте»: «Наименование линии» = «по экспликации».")

        # --- «Щиты», «Шины», «Линии»: шапки + данные
        wsb = wb.Worksheets.Add(After=wk)
        wsb.Name = "Щиты"
        wsl = wb.Worksheets.Add(After=wb.Worksheets("Щиты"))
        wsl.Name = "Линии"
        sheads = [c["head"] for c in S]
        bvals = [[b.get(h) if c["kind"] == "in" else None for c, h in zip(S, sheads)] for b in boards]
        lob = make_table(wsb, "Щиты", sheads, bvals)
        # таблица «Шины» — справа от таблицы щитов (нужна, только если тип шины не определяется по имени)
        c0 = len(S) + 2
        bus_rows = [[b.get("bus"), b.get("board"), b.get("type")] for b in data.get("buses", [])] or [[None, None, None]]
        wsb.Cells(1, c0).Value = "ШИНЫ — только если тип шины не определяется по имени (ИБП, ДГУ) или по секции"
        wsb.Cells(1, c0).Font.Bold = True
        wsb.Range(wsb.Cells(2, c0), wsb.Cells(2, c0 + 2)).Value = [["Шина", "Щит", "Тип"]]
        wsb.Range(wsb.Cells(3, c0), wsb.Cells(2 + len(bus_rows), c0 + 2)).Value = bus_rows
        lsh = wsb.ListObjects.Add(1, wsb.Range(wsb.Cells(2, c0), wsb.Cells(2 + len(bus_rows), c0 + 2)), None, 1)
        lsh.Name = "Шины"
        lsh.TableStyle = "TableStyleLight9"
        validation(lsh.ListColumns("Тип").DataBodyRange, "Ввод 1|Ввод 2|Ввод 3|ИБП|ДГУ|Одиночная", sep)
        IN_KINDS = ("in", "x")
        lheads = [c["head"] for c in L]
        lvals = [[as_text(r.get(c["key"])) if c["kind"] in IN_KINDS else None for c in L] for r in rows]
        lol = make_table(wsl, "Линии", lheads, lvals)
        log("таблицы: щитов %d, строк %d" % (len(boards), len(rows)))
        set_names(wb, {"Список_щитов": "=Щиты[Щит]", "Список_марок": "=Кабели[Обозначение]"})
        # формулы служебных и результатов
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
        # серые вводы: формула =Линии[@[_авто …]], поверх — ручные значения из данных
        n_man = 0
        for c in L:
            if c["kind"] in ("a", "xa"):
                lol.ListColumns(c["head"]).DataBodyRange.Formula2 = "=Линии[@[%s]]" % esc(HEAD[c["auto"]])
        # ручные значения поверх: автозаполнение формул в таблицах на время записи выключено,
        # иначе Excel по первой ячейке-формуле заполнит весь столбец и затрёт ручные значения
        af_prev = xl.AutoCorrect.AutoFillFormulasInLists
        xl.AutoCorrect.AutoFillFormulasInLists = False
        try:
            for c in L:
                if c["kind"] not in ("a", "xa"):
                    continue
                rng = lol.ListColumns(c["head"]).DataBodyRange
                for i_, r in enumerate(rows):
                    v = r.get(c["key"])
                    if v not in (None, ""):
                        rng.Cells(i_ + 1, 1).Value = as_text(v)
                        n_man += 1
        finally:
            xl.AutoCorrect.AutoFillFormulasInLists = af_prev
        log("серые вводы: ручных значений %d" % n_man)

        # оформление шапки: ввод — жёлтый, результат — серый
        C_RES = 15132390
        for i, c in enumerate(L):
            hc = lol.HeaderRowRange.Cells(1, i + 1)
            hc.WrapText = True
            hc.Interior.Color = C_IN if c["kind"] in ("in", "a", "x", "xa") else C_RES
            if c["kind"] in ("a", "xa"):
                hc.Font.Italic = True
            if c.get("note"):
                try:
                    hc.AddComment(c["note"])
                except Exception:
                    pass
        for i, c in enumerate(S):
            hc = lob.HeaderRowRange.Cells(1, i + 1)
            hc.WrapText = True
            hc.Interior.Color = C_IN if c["kind"] in ("in", "a") else C_CALC
            if c["kind"] == "a":
                hc.Font.Italic = True
            if c.get("note"):
                try:
                    hc.AddComment(c["note"])
                except Exception:
                    pass
        wsl.Rows(2).RowHeight = 48
        wsb.Rows(2).RowHeight = 60
        wsb.Columns.ColumnWidth = 11
        wsl.Columns.ColumnWidth = 10
        widths = {"board": 9, "line": 13, "name_in": 30, "n": 9, "p_in": 8, "grp_mark": 7, "qf_in": 16, "rcd_in": 8, "ctl_in": 12, "u_in": 12,
                  "dev_in": 14, "s_in": 8, "lplan_in": 9, "phase_in": 9, "xt_in": 8, "issues": 38, "i": 8, "l": 8, "du_vru": 8,
                  "cab_show": 22, "place_lbl": 10}
        for k_, w_ in widths.items():
            wsl.Columns(L.index(next(c for c in L if c["key"] == k_)) + 1).ColumnWidth = w_
        # группы столбцов: «Ещё» — свернуть кнопкой «+» (кнопка над последним видимым столбцом), служебные — скрыть
        try:
            wsl.Outline.SummaryColumn = -4131
        except Exception:
            pass
        for i, c in enumerate(L):
            if c["kind"] in ("x", "xa", "xc"):
                wsl.Columns(i + 1).OutlineLevel = 2
        for i, c in enumerate(S):
            if c["kind"] == "h":
                wsb.Columns(i + 1).OutlineLevel = 2
        # формат чисел у результатов
        for lo, spec in ((lob, S), (lol, L)):
            for c in spec:
                if c["kind"] not in ("in", "a", "x", "xa") and any(u in c["head"] for u in ("кВт", ", А", "%", "Ом", "кА", "квар", "кВА", "cos φ", "Кс", "Ко", "Iкз / In", "Длина")):
                    try:
                        lo.ListColumns(c["head"]).DataBodyRange.NumberFormatLocal = "0,00"
                    except Exception:
                        pass
        lol.ListColumns("Длина, м").DataBodyRange.NumberFormatLocal = "0"
        lol.ListColumns("Ток утечки, мА").DataBodyRange.NumberFormatLocal = "0,0"
        lol.ListColumns("Lмакс по КЗ, м").DataBodyRange.NumberFormatLocal = "0"
        lol.ListColumns("БП: нужно, Вт").DataBodyRange.NumberFormatLocal = "0"
        lol.ListColumns("Uвых БП, В").DataBodyRange.NumberFormatLocal = "0,0"
        # выпадающие списки (подсказки, не запрет)
        validation(lol.ListColumns("Щит").DataBodyRange, "=Список_щитов", sep)
        validation(lol.ListColumns("Группа").DataBodyRange, "гр|узо", sep)
        validation(lol.ListColumns("Утечка, мА").DataBodyRange, "10|30|100|300|нет", sep)
        validation(lol.ListColumns("Управление").DataBodyRange, "Relay|DIM|DALI|шторы|клапан", sep)
        validation(lol.ListColumns("Режим").DataBodyRange, "вручную|рисовать", sep)
        validation(lol.ListColumns("Марка кабеля").DataBodyRange, "=Список_марок", sep)
        # условное форматирование «Линий» (формулы — в записи русского Excel)
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
        for i, c in enumerate(L):
            col_ = letters(i + 1)
            rng = wsl.Range(f"{col_}{R0}:{col_}{lastr}")
            if c["kind"] in ("a", "xa"):
                g_ = rng.FormatConditions.Add(2, None, f'=ЕФОРМУЛА({col_}{R0})')
                g_.Font.Color = 8421504
                g_.Font.Italic = True
                b_ = rng.FormatConditions.Add(2, None, f'=И(НЕ(ЕФОРМУЛА({col_}{R0}));{col_}{R0}<>"")')
                b_.Font.Color = 12611584
                b_.Font.Bold = True
            if c["key"] == "issues":
                e_ = rng.FormatConditions.Add(2, None, f'={col_}{R0}<>""')
                e_.Interior.Color = 13421823
            if c["key"] == "grp_mark":
                m_ = rng.FormatConditions.Add(2, None, f'={col_}{R0}<>""')
                m_.Font.Color = 12611584
                m_.Font.Bold = True
        for i, c in enumerate(S):
            if c["kind"] == "a":
                col_ = letters(i + 1)
                rng = wsb.Range(f"{col_}3:{col_}{3 + len(boards) + 100}")
                g_ = rng.FormatConditions.Add(2, None, f'=ЕФОРМУЛА({col_}3)')
                g_.Font.Color = 8421504
                g_.Font.Italic = True
                b_ = rng.FormatConditions.Add(2, None, f'=И(НЕ(ЕФОРМУЛА({col_}3));{col_}3<>"")')
                b_.Font.Color = 12611584
                b_.Font.Bold = True
        # служебные столбцы — скрыть
        for i, c in enumerate(L):
            if c["kind"] == "h":
                wsl.Columns(i + 1).Hidden = True
        try:
            wsl.Outline.ShowLevels(0, 1)
            wsb.Outline.ShowLevels(0, 1)
        except Exception:
            pass

        # --- «Однолинейка»: интерфейс для SXDRAW (спилл столбцов «Линий»)
        w1 = wb.Worksheets("Однолинейка")
        w1.Range("A3:CA1000").ClearContents()
        for c_, k in ONE_MAP.items():
            w1.Range(f"{c_}3").Formula2 = '=LET(v,FILTER(Линии[%s],Линии[_№ места (итог)]<>""),IF(v="","",v))' % esc(HEAD[k])
        w1.Range("AE1").Formula2 = '="Ошибок: "&COUNTIF($BU$3:$BU$3000,2)&"   Предупреждений: "&COUNTIF($BU$3:$BU$3000,1)'

        root = next((b["Щит"] for b in boards if not b.get("Питается от")), boards[0]["Щит"])
        build_loads(wb, xl, root, sep)
        build_pz(wb, root, sep)
        extend_cable_diameters(wb)
        build_journal_v5(wb, sep)
        spec_by_board(wb, sep)
        build_spec_v5(wb, sep)
        build_board(wb, root, sep)
        build_tags(wb)
        single_line_view(wb)
        # --- «Проверка»
        wp = wb.Worksheets.Add(Before=wb.Worksheets(1))
        wp.Name = "Проверка"
        wp.Range("A1").Value = "ПРОВЕРКА ОБЪЕКТА"
        wp.Range("A1").Font.Bold = True
        wp.Range("A1").Font.Size = 14
        wp.Range("A2").Formula2 = '="Замечаний: щиты — "&COUNTIF(Щиты[Замечания щита],"?*")&", линии — "&COUNTIF(Линии[Замечания],"?*")&". Методика: "&Методика'
        wp.Range("A4:B4").Value = [["Щит", "Замечания щита"]]
        wp.Range("A5").Formula2 = '=FILTER(HSTACK(Щиты[Щит],Щиты[Замечания щита]),Щиты[Замечания щита]<>"","замечаний нет")'
        wp.Range("D4:H4").Value = [["Щит", "Номер линии", "Наименование", "Замечания", "Перейти"]]
        wp.Range("D5").Formula2 = ('=FILTER(HSTACK(Линии[Щит],Линии[Номер линии],Линии[Наименование (итог)],Линии[Замечания],'
                                   'HYPERLINK("#\'Линии\'!B"&(Линии[_n]+2),"→")),Линии[Замечания]<>"","замечаний нет")')
        for a in ("A4:B4", "D4:H4"):
            wp.Range(a).Font.Bold = True
            wp.Range(a).Interior.Color = C_HEAD
        for c_, w_ in (("A", 12), ("B", 70), ("D", 12), ("E", 16), ("F", 40), ("G", 70), ("H", 8)):
            wp.Columns(c_).ColumnWidth = w_
        wp.Columns("B").WrapText = True
        wp.Columns("G").WrapText = True

        build_changes_v5(wb, sep)
        build_labels_sheet(wb)

        # --- «Линии»: карточка щита, кнопки щитов, закрепление
        wsl.Range("A1:C1").Merge()
        wsl.Range("A1").Formula2 = CARD
        wsl.Range("A1").WrapText = True
        wsl.Range("A1").VerticalAlignment = -4108
        wsl.Range("A1").Font.Size = 10
        wsl.Range("A1:C1").Interior.Color = 15921906
        sl = add_slicer(wb, wsl, lol, len(boards))
        bl = (sl.Left + sl.Width + 8) if sl is not None else wsl.Range("D1").Left + 4
        add_button(wsl, bl, 4, 120, 22, "Починить формулы", "SX_RepairFormulas")
        add_button(wsl, bl + 125, 4, 120, 22, "Упорядочить щит", "SX_SortBoard")
        add_button(wsl, bl + 250, 4, 115, 22, "Закрепить фазы", "SX_FreezePhases")
        add_button(wsl, bl + 370, 4, 150, 22, "Зафиксировать выдачу", "SX_Snapshot")
        add_button(wp, wp.Range("D1").Left, 4, 130, 22, "Починить формулы", "SX_RepairFormulas")
        wsc = wb.Worksheets("Щит")
        add_button(wsc, wsc.Range("L1").Left, 4, 190, 24, "Выгрузить для AutoCAD (все щиты)", "SX_ExportAutoCAD")
        add_button(wsc, wsc.Range("L1").Left + 200, 4, 120, 24, "Наклейки (A4)", "SX_Labels")
        wch = wb.Worksheets("Изменения")
        add_button(wch, wch.Range("J1").Left, 4, 150, 22, "Зафиксировать выдачу", "SX_Snapshot")
        wsl.Activate()
        xl.ActiveWindow.SplitRow = 2
        xl.ActiveWindow.SplitColumn = 3
        xl.ActiveWindow.FreezePanes = True

        xl.Calculation = -4105
        t1 = time.time()
        xl.CalculateFull()
        log("пересчёт: %.1f с" % (time.time() - t1))
        for ws_ in (wsl, wsb, w1):
            try:
                e = ws_.UsedRange.SpecialCells(-4123, 16)
                log("ошибки формул на «%s»: %d (%s)" % (ws_.Name, e.Count, e.Address[:200]))
            except Exception:
                log("ошибки формул на «%s»: 0" % ws_.Name)
            try:
                cr = ws_.CircularReference
                if cr is not None:
                    log("ЦИКЛИЧЕСКАЯ ССЫЛКА на «%s»: %s" % (ws_.Name, cr.Address))
            except Exception:
                pass
        res = compare(wsl, lol, wsb, lob, rows, boards)

        # --- шаблоны формул для макроса «Починить формулы» (лист скрыт)
        wt = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        wt.Name = "_формулы"
        tpl = [["Столбец", "Формула (R1C1)", "Вид", "Серое из"]]
        for c in L:
            if c["kind"] in ("a", "xa"):
                tpl.append([c["head"], "'=Линии[@[%s]]" % esc(HEAD[c["auto"]]), c["kind"], HEAD[c["auto"]]])
            elif c["f"]:
                f1 = lol.ListColumns(c["head"]).DataBodyRange.Cells(1, 1).Formula2R1C1
                tpl.append([c["head"], "'" + f1, "c" if c["kind"] == "xc" else c["kind"], None])
        wt.Range("A1:D%d" % len(tpl)).Value = tpl
        wt.Visible = 2

        # гайд, методика, порядок листов
        import guide_v5
        import guide_v312
        import methodology_sheet
        guide_v312.G = guide_v5.G
        guide_v312.build(wb)
        methodology_sheet.render(wb)
        put_guides(wb)
        tidy_reference(wb)
        try:
            wb.Worksheets("Нагрузки").Name = "Режимы"
        except Exception:
            pass
        order = ["Гайд", "Проверка", "Объект", "Щиты", "Линии", "Коды", "Щит", "Кабельный_журнал", "Спецификация", "Бирки", "Изменения",
                 "Наклейки", "Каталог",
                 "Методика", "Таблицы_норм", "Режимы", "Расчёт_нагрузок", "Спецификация_авто", "Однолинейка", "Справочная_потребители"]
        for nm in reversed(order):
            try:
                wb.Worksheets(nm).Move(Before=wb.Worksheets(1))
            except Exception:
                pass
        for nm in ("Справочная_потребители", "Спецификация_авто", "Однолинейка", "_выдача"):     # служебные: читают программы и макросы
            try:
                wb.Worksheets(nm).Visible = 0
            except Exception:
                pass
        green, blue, grey, violet, yellow = 5287936, 12419407, 10921638, 10498160, 13431551
        for nm, color in (("Гайд", green), ("Проверка", 255), ("Объект", yellow), ("Щиты", yellow), ("Линии", yellow), ("Коды", yellow),
                          ("Щит", blue), ("Кабельный_журнал", blue), ("Спецификация", blue), ("Бирки", blue), ("Изменения", blue), ("Наклейки", blue), ("Каталог", grey),
                          ("Методика", green), ("Таблицы_норм", grey), ("Режимы", violet), ("Расчёт_нагрузок", violet)):
            try:
                wb.Worksheets(nm).Tab.Color = color
            except Exception:
                pass
        add_vba(wb, wsl)
        wb.Worksheets("Линии").Activate()
        wb.Worksheets("Линии").Range("D3").Select()
        if os.path.exists(out):
            os.remove(out)
        wb.SaveAs(os.path.abspath(out), 52)
        log("сохранено: %s (%.0f с)" % (out, time.time() - t0))
        wb.Close(False)
    try:
        os.remove(tmp)
    except Exception:
        pass
    return res


def compare(wsl, lol, wsb, lob, rows, boards):
    """сверка с результатами v3.12 (rows[i]['_ref'], boards[i]['_ref'])."""
    def colvals(lo, head):
        v = lo.ListColumns(head).DataBodyRange.Value
        return [x[0] for x in v] if isinstance(v, tuple) else [v]
    keys = (("p", "p"), ("i", "i"), ("qf", "in_grp"), ("s", "s"), ("l", "l"), ("du", "du"), ("du_vru", "du_vru"), ("grp", "grp"), ("cos", "cos"))
    got = {kg: colvals(lol, HEAD[kg]) for _, kg in keys}
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
                diffs.append(f"{r['board']} {r.get('line')}: {kr} v3.12={a} v5={b}")
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
    ap.add_argument("--out", required=True, help="файл .xlsm")
    ap.add_argument("--mode", default="нормы 2026", help="v3.12 или «нормы 2026»")
    ap.add_argument("--defaults", action="store_true", help="заполнить «Коды» значениями v5 по умолчанию (шаблон, пример)")
    a = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    with open(a.data, encoding="utf-8") as fh:
        data = json.load(fh)
    build(data, a.base, a.out, a.mode, a.defaults)


if __name__ == "__main__":
    main()
