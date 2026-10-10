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
