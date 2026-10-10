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
