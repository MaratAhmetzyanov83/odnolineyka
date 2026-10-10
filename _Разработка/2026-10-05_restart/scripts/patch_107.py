# -*- coding: utf-8 -*-
"""1.06 -> 1.07: автоподбор в «Исходных данных» (шаг 3).

- пустая марка (N) и жилы (O) — подставляются (итог в BB/BC); своё значение главнее;
- номинал автомата не меньше минимального по коду (розетки 16 А, свет 10 А);
- дифзащита по коду (BD), ток утечки линии по ПУЭ 7.1.83 (BE), Lмакс по КЗ (BF);
- коэффициент прокладки («Настройки») в подборе сечения;
- наименование: по экспликации (лист «Экспликация»), «Пом. №» — по настройке;
- на «Диапазонах» по каждому щиту: мест под вводной узел, Zпетли на шинах.
Запуск: py patch_107.py <1.06.xlsx> <1.07.xlsx>
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import patch_106 as P6  # noqa: E402

R0, R1 = 3, 1000

# Код | УГО | Дифзащита, мА | Мин. автомат, А | Каналов модуля на линию | Клемм 3L на линию
CODES = [
    ("R", "Розетка", 30, 16, 1, None, 2.5), ("RС", "Розетка", 30, 16, 1, None, 2.5),
    ("L", "Светильник", None, 10, 1, None, None), ("LD", "Светильник", None, 10, 1, None, None),
    ("LED", "Светильник", None, 10, 1, None, None), ("LDT", "Светильник", None, 10, 1, None, None),
    ("DALI", "Светильник", None, 10, 1, None, None), ("Н", "Прочее", None, 10, 1, None, None),
    ("SH", "Привод (двигатель)", None, None, 2, 2, None), ("М", "Привод (двигатель)", None, None, 2, 2, None),
    ("К", "Привод (двигатель)", None, 16, 1, None, 2.5), ("K", "Привод (двигатель)", None, 16, 1, None, 2.5),
    ("П", "Привод (двигатель)", None, None, 1, None, None), ("В", "Привод (двигатель)", None, None, 1, None, None),
    ("НО", "Привод (двигатель)", None, None, 1, None, None), ("УВ", "Привод (двигатель)", None, None, 1, None, None),
    ("ТП", "Нагреватель", 30, 16, 1, None, 1.5), ("ПН", "Нагреватель", 30, 16, 1, None, 2.5),
    ("КМН", "Нагреватель", None, None, 1, None, None),
    ("ЭК", "Клапан", None, None, 1, None, None), ("ЭK", "Клапан", None, None, 1, None, None),
    ("ТЕ-О", "Сервопривод", None, None, 1, None, None), ("ТЕ-П", "Сервопривод", None, None, 1, None, None),
    ("ТЕ-У", "Сервопривод", None, None, 1, None, None), ("TE", "Сервопривод", None, None, 1, None, None),
    ("ТЕ", "Сервопривод", None, None, 1, None, None),
    ("WD", "Прочее", None, None, 1, None, None), ("АПС", "Прочее", None, None, 1, None, None),
    ("PS", "Прочее", None, None, 1, None, None), ("LM", "Прочее", None, None, 1, None, None),
    ("АСУ", "Прочее", None, None, 1, None, None), ("ШТК", "Прочее", None, 16, 1, None, 2.5),
    ("ЩОПС", "Прочее", None, 16, 1, None, 2.5), ("KV", "Прочее", None, None, 1, None, None),
    ("RG", "Прочее", None, 16, 1, None, 2.5), ("КМ", "Прочее", None, None, 1, None, None),
    ("Ввод", "", None, None, 1, None, None), ("АВР", "", None, None, 1, None, None),
    ("КУП", "", None, None, 1, None, None), ("КуП", "", None, None, 1, None, None),
    ("РЕ", "", None, None, 1, None, None), ("PE", "", None, None, 1, None, None),
]
UGO = [("Розетка", "Розетка1", "SOCKET"), ("Светильник", "$RECOVER_230831092053-0", "LAMP"),
       ("Сервопривод", "Сервопривод-1", "SERVO"), ("Клапан", "EK.1", "VALVE"),
       ("Привод (двигатель)", "Привод", "MOTOR"), ("Нагреватель", "(значок программы)", "HEAT"),
       ("Прочее", "(прямоугольник с подписью)", "BOX")]

LOG_107 = ("1. Марка и жилы: пусто — по умолчанию (марка — «Настройки», жилы: 1ф — 3, 3ф — 5, до 50 В — 2); "
           "итог в BB/BC, своё значение главнее.\n"
           "2. Автомат не меньше минимального по коду («Настройки»: розетки 16 А, свет 10 А).\n"
           "3. Дифзащита по коду (BD), ток утечки линии по ПУЭ 7.1.83 (BE, красный — больше 1/3 уставки), "
           "Lмакс по КЗ (BF, если на «Диапазонах» задано Zпетли на шинах щита).\n"
           "4. Сечение = наибольшее из: по току (с коэффициентом прокладки, «Настройки»), минимального по коду "
           "(розетки 2,5) и по потерям (не больше 4,5 %, до 50 В — 5 %). «Сечение (ручное)» главнее.\n"
           "5. Наименование: вариант «по экспликации» (Настройки: «да») — по названию помещения с листа «Экспликация»; "
           "«Пом. №» можно отключить («Настройки»).\n"
           "6. «Диапазоны» H:I — по щиту: мест под вводной узел (пусто = 17) и Zпетли на шинах, Ом (из ТКЗ).")


def set_table(sh, top_left_row, col, header, rows):
    c0 = sh.Range(f"{col}{top_left_row}").Column
    n = len(header)
    sh.Range(sh.Cells(top_left_row, c0), sh.Cells(top_left_row, c0 + n - 1)).Value = (tuple(header),)
    sh.Range(sh.Cells(top_left_row, c0), sh.Cells(top_left_row, c0 + n - 1)).Font.Bold = True
    for i, r in enumerate(rows):
        sh.Range(sh.Cells(top_left_row + 1 + i, c0), sh.Cells(top_left_row + 1 + i, c0 + n - 1)).Value = \
            (tuple("" if x is None else x for x in r),)


def patch(wb, log=print):
    app = wb.Application
    sp = wb.Worksheets("Справочная")
    src = wb.Worksheets("Исходные данные")
    if str(sp.Range("AQ3").Value or "") == "Код":
        log("1.07 уже применён — пропуск")
        return

    # --- Справочная: коды, УГО, настройки ---
    sp.Range("AQ2").Value = "НАСТРОЙКИ КОДОВ (1.07): код = буквы номера линии до первой точки; пусто — нет ограничения"
    sp.Range("AQ2").Font.Bold = True
    set_table(sp, 3, "AQ", ["Код", "УГО", "Дифзащита, мА", "Мин. автомат, А", "Каналов модуля на линию",
                            "Клемм 3L на линию", "Мин. сечение, мм²"], CODES)
    P6.add_name(wb, "Коды_доп", "=Справочная!$AQ$4:$AW$150")
    P6.add_name(wb, "Коды_доп_код", "=Справочная!$AQ$4:$AQ$150")
    sp.Range("AX2").Value = "УГО → БЛОК AutoCAD (SXDRAW)"
    sp.Range("AX2").Font.Bold = True
    set_table(sp, 3, "AX", ["УГО", "Блок AutoCAD", "Код программы (не менять)"], UGO)
    P6.add_name(wb, "УГО_блоки", "=Справочная!$AX$4:$AZ$40")
    sp.Range("AL25").Value = "Коэффициент прокладки (гофра, пучок; 1 — как в ПУЭ табл. 1.3.6)"
    sp.Range("AM25").Value = 1
    sp.Range("AL26").Value = "Марка кабеля по умолчанию (код из L)"
    sp.Range("AM26").Value = "мс"
    sp.Range("AL27").Value = "Наименование: писать «Пом. №» (да / нет)"
    sp.Range("AM27").Value = "да"
    sp.Range("AL28").Value = "Мест на лист схемы (A3, шаг 20 мм)"
    sp.Range("AM28").Value = 17
    P6.add_name(wb, "К_прокл", "=Справочная!$AM$25")
    P6.add_name(wb, "Марка_умолч", "=Справочная!$AM$26")
    P6.add_name(wb, "Наим_пом", "=Справочная!$AM$27")
    P6.add_name(wb, "Мест_на_лист", "=Справочная!$AM$28")
    sp.Range("AL32").Value = "Наименование по экспликации (да / нет)"
    sp.Range("AM32").Value = "нет"
    P6.add_name(wb, "Наим_экспл", "=Справочная!$AM$32")
    log("Справочная: коды AQ:AV, УГО AX:AZ, настройки AM25:AM28")

    # --- Экспликация ---
    ex = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ex.Name = "Экспликация"
    ex.Range("A1").Value = "ЭКСПЛИКАЦИЯ ПОМЕЩЕНИЙ: название идёт в наименование линии вместо «Этаж Пом. №…» " \
                           "(номер линии Код.Этаж.Помещение.Группа)"
    ex.Range("A1").Font.Bold = True
    ex.Range("A2:D2").Value = (("Этаж (как в номере)", "Помещение (как в номере)", "Название", "_ключ", None),)
    ex.Range("A2:D2").Font.Bold = True
    ex.Range("D3:D500").Formula = '=IF(OR(A3="",B3=""),"",A3&"."&B3)'
    ex.Range("A3:B500").NumberFormat = "@"
    ex.Columns("A:B").ColumnWidth = 14
    ex.Columns("C").ColumnWidth = 40
    ex.Columns("D").Hidden = True
    P6.add_name(wb, "Экспл_ключ", "=Экспликация!$D$3:$D$500")
    P6.add_name(wb, "Экспл_назв", "=Экспликация!$C$3:$C$500")
    log("Лист «Экспликация»")

    # --- Диапазоны: по щиту ---
    dg = wb.Worksheets("Диапазоны")
    dg.Range("H3:I3").Value = (("Мест под вводной узел (пусто = 17)", "Zпетли на шинах, Ом (из ТКЗ)"),)
    dg.Range("H3:I3").Font.Bold = True
    dg.Range("H3:I3").WrapText = True
    dg.Columns("H:I").ColumnWidth = 18
    P6.add_name(wb, "Щиты_имя", "=Диапазоны!$B$4:$B$33")
    P6.add_name(wb, "Щиты_узел", "=Диапазоны!$H$4:$H$33")
    P6.add_name(wb, "Щиты_Zш", "=Диапазоны!$I$4:$I$33")
    log("Диапазоны: H:I по щитам")

    # --- Исходные данные: итоговые марка/жилы и новые столбцы ---
    elv = "AND(ISNUMBER(F{r}),F{r}>0,F{r}<=50)"
    E = lambda s: s.replace("{r}", str(R0))
    heads = {"BB": "Марка (итог)", "BC": "Жил (итог)", "BD": "Дифзащита по коду, мА",
             "BE": "Ток утечки линии, мА (ПУЭ 7.1.83)", "BF": "Lмакс по КЗ, м", "BG": "Мин. автомат по коду, А"}
    for c, h in heads.items():
        src.Range(f"{c}2").Value = h
    src.Range("BB2:BG2").Font.Bold = True
    src.Range("BB2:BG2").WrapText = True
    src.Range(f"BB{R0}:BB{R1}").Formula2 = E('=IF(D{r}="","",IF(TRIM(N{r}&"")="",Марка_умолч,N{r}))')
    src.Range(f"BC{R0}:BC{R1}").Formula2 = E('=IF(D{r}="","",IF(TRIM(O{r}&"")<>"",O{r},IF(' + elv + ',2,IF(H{r}=3,5,3))))')
    src.Range(f"BD{R0}:BD{R1}").Formula2 = E(
        '=IF(D{r}="","",IFERROR(LET(v,XLOOKUP(AH{r},Коды_доп_код,INDEX(Коды_доп,0,3),""),IF(N(v)>0,v,"")),""))')
    src.Range(f"BE{R0}:BE{R1}").Formula2 = E(
        '=IF(OR(D{r}="",' + elv + ',F{r}=16),"",IFERROR(ROUND(0.4*N(T{r})+0.01*N(M{r}),2),""))')
    src.Range(f"BG{R0}:BG{R1}").Formula2 = E(
        '=IF(D{r}="","",IFERROR(LET(v,XLOOKUP(AH{r},Коды_доп_код,INDEX(Коды_доп,0,4),""),IF(N(v)>0,v,"")),""))')
    src.Range(f"BF{R0}:BF{R1}").Formula2 = E(
        '=IF(OR(D{r}="",' + elv + ',F{r}=16),"",IFERROR(LET(sh,INDEX(Разбивка_по_щитам!$F$3:$F$1000,'
        'MATCH(D{r},Разбивка_по_щитам!$H$3:$H$1000,0)),zs,XLOOKUP(sh,Щиты_имя,Щиты_Zш,""),'
        'IF(OR(N(zs)=0,NOT(ISNUMBER(V{r})),NOT(ISNUMBER(U{r}))),"",'
        'MAX(0,ROUNDDOWN((0.95*Uн_В/SQRT(3)/(10*U{r})-zs)*V{r}/(2*W{r}),0)))),""))')
    src.Range(f"BB{R0}:BG{R1}").Interior.Color = 0xF2F2F2

    # U: не меньше минимального по коду
    old_u = "=IF(RC[-3]=0,VLOOKUP(RC[-1],Автомат_номиналы,2,TRUE),RC[-3])"
    new_u = ('=IF(RC[-3]=0,LET(a,VLOOKUP(RC[-1],Автомат_номиналы,2,TRUE),m,N(RC[38]),'
             'IF(AND(ISNUMBER(a),m>a),m,a)),RC[-3])')
    _swap(src, "U", old_u, new_u, log)
    # BH: минимальное сечение по коду
    src.Range("BH2").Value = "Мин. сечение по коду, мм²"
    src.Range("BH2").Font.Bold = True
    src.Range("BH2").WrapText = True
    src.Range(f"BH{R0}:BH{R1}").Formula2 = E(
        '=IF(D{r}="","",IFERROR(LET(v,XLOOKUP(AH{r},Коды_доп_код,INDEX(Коды_доп,0,7),""),IF(N(v)>0,v,"")),""))')
    src.Range(f"BH{R0}:BH{R1}").Interior.Color = 0xF2F2F2
    # V: сечение = max(по току с К прокладки, мин. по коду, по потерям ≤ 4,5 % / 5 % до 50 В);
    # жилы и марка — итоговые (BB/BC). Ссылки R1C1 от столбца V:
    # U=RC[-1] T=RC[-2] P=RC[-6] M=RC[-9] H=RC[-14] F=RC[-16] N=RC[-8] Y=RC[3] Z=RC[4] BB=RC[32] BC=RC[33] BH=RC[38]
    old_pick = ('IFERROR(XLOOKUP(RC[-1]/IF(AND(RC[-14]=3,RC[-7]>=4),0.92,1),'
                'IF(RC[-7]=1,Iдоп_1ж,IF(RC[-14]=3,Iдоп_3ж,Iдоп_2ж)),Iдоп_S,"Не исп",1),"Не исп")')
    dc = "AND(ISNUMBER(RC[-16]),RC[-16]>0,RC[-16]<=50,RC[-16]<>16)"
    new_pick = (
        'LET(sI,IFERROR(XLOOKUP(RC[-1]/(IF(AND(RC[-14]=3,RC[33]>=4),0.92,1)*К_прокл),'
        'IF(RC[33]=1,Iдоп_1ж,IF(RC[-14]=3,Iдоп_3ж,Iдоп_2ж)),Iдоп_S,"Не исп",1),"Не исп"),'
        'dc,' + dc + ',rho,INDEX(Кабель_сопротивление,2,3),'
        'b,IF(dc,2,IF(RC[-14]=3,1,2)),cs,IF(dc,1,IFERROR(RC[-6],0.95)),sn,IF(dc,0,IFERROR(SQRT(1-cs^2),0.31)),'
        'u0,IF(dc,RC[-16],Uн_В/SQRT(3)),lim,IF(dc,Потери_до50В,4.5),'
        'den,lim/100*u0-b*0.08*N(RC[-9])*sn*N(RC[-2])/1000,'
        'sreq,IF(OR(N(RC[-9])=0,N(RC[-2])=0),0,IF(den<=0,9999,b*rho*N(RC[-9])*cs*N(RC[-2])/den)),'
        'sD,IF(sreq=0,0,XLOOKUP(sreq,Iдоп_S,Iдоп_S,MAX(Iдоп_S),1)),'
        'IF(ISNUMBER(sI),MAX(sI,N(RC[38]),sD),sI))')
    rng = src.Range(f"V{R0}:V{R1}")
    fl = [x[0] for x in rng.Formula2R1C1]
    nfl, n = [], 0
    for f in fl:
        if isinstance(f, str):
            if old_pick in f:
                n += 1
            f = f.replace(old_pick, new_pick)
            f = f.replace("VLOOKUP(RC[-8],Кабель_обозначение", "VLOOKUP(RC[32],Кабель_обозначение")
        nfl.append((f,))
    rng.Formula2R1C1 = tuple(nfl)
    log(f"V: подбор сечения (ток, код, потери) — {n} строк")
    # AD, AE: итоговая марка и жилы
    src.Range(f"AD{R0}:AD{R1}").Formula = f"=VLOOKUP(BB{R0},Кабель_обозначение,2,FALSE)"
    src.Range(f"AE{R0}:AE{R1}").Formula = \
        f'=IF(VLOOKUP(BB{R0},Кабель_обозначение,3,FALSE)=1,AD{R0}&" ("&BC{R0}&"х"&V{R0}&")",AD{R0})'
    # P (cos φ): неизвестный код — 0,95 вместо ошибки (ручные значения не трогаем)
    rng = src.Range(f"P{R0}:P{R1}")
    fl = [x[0] for x in rng.FormulaR1C1]
    std = fl[0] if isinstance(fl[0], str) and fl[0].startswith("=IF(VLOOKUP(") else None
    if std:
        rng.FormulaR1C1 = tuple(((("=IFERROR(" + f[1:] + ",0.95)") if f == std else f),) for f in fl)
    # B: экспликация и «Пом. №»
    old_b = src.Range(f"B{R0}").Formula
    if old_b.startswith("=IFERROR("):
        inner = old_b[1:]
        new_b = ('=LET(nm,IFERROR(' + inner + ',""),rm,IFERROR(XLOOKUP(AI3&"."&AG3,Экспл_ключ,Экспл_назв,""),""),'
                 'bs,IFERROR(VLOOKUP(AH3,Потребители_обозначение,2,FALSE),""),'
                 'IF(AND(Наим_экспл="да",rm<>"",bs<>""),bs&" "&rm,IF(Наим_пом="нет",SUBSTITUTE(nm," Пом. №"," "),nm)))')
        src.Range(f"B{R0}:B{R1}").Formula2 = new_b
    # условное форматирование: ток утечки > 1/3 уставки
    fc = None
    for f1 in (f"=AND(N($BD{R0})>0,N($BE{R0})>$BD{R0}/3)", f"=И(Ч($BD{R0})>0;Ч($BE{R0})>$BD{R0}/3)"):
        try:
            fc = src.Range(f"BE{R0}:BE{R1}").FormatConditions.Add(2, 0, f1)
            break
        except Exception:
            fc = None
    if fc is not None:
        fc.Interior.Color = 0xCEC7FF
    else:
        log("ВНИМАНИЕ: подсветка тока утечки не добавлена")
    for addr, text in {
        "N2": "1.07: пусто — марка по умолчанию («Настройки»). Итог — столбец BB.",
        "O2": "1.07: пусто — 1ф: 3 жилы, 3ф: 5, до 50 В: 2. Итог — столбец BC.",
        "BE2": "ПУЭ 7.1.83: 0,4 мА на 1 А тока нагрузки + 0,01 мА на 1 м длины. Сумма токов утечки линий "
               "за одним УЗО/дифавтоматом — не больше 1/3 уставки (красный — уже по одной линии).",
        "BF2": "Наибольшая длина при этом сечении и номинале, когда Iкз в конце ≥ 10·In (характеристика C, "
               "ПУЭ 1.7.79). Нужен Zпетли на шинах щита (лист «Диапазоны», столбец I, — из ТКЗ).",
    }.items():
        c = src.Range(addr)
        try:
            if c.Comment is not None:
                c.Comment.Delete()
        except Exception:
            pass
        c.AddComment(text)
    log("Исходные данные: BB:BG, U/V/AD/AE/B переписаны")

    # --- Краткое руководство ---
    kr = wb.Worksheets("Краткое руководство")
    kr.Range("C1").Value = "Версия таблицы 1.07"
    r = 2
    while kr.Range(f"E{r}").Value not in (None, ""):
        r += 1
    kr.Range(f"E{r}").Value = "1.07"
    kr.Range(f"F{r}").Value = LOG_107
    kr.Range(f"F{r}").WrapText = True
    log("Краткое руководство: 1.07")


def _swap(sh, col, old, new, log):
    rng = sh.Range(f"{col}{R0}:{col}{R1}")
    fl = [x[0] for x in rng.FormulaR1C1]
    n = sum(1 for f in fl if f == old)
    rng.Formula2R1C1 = tuple(((new if f == old else f),) for f in fl)
    log(f"{col}: заменено {n} из {len(fl)}")


def main(src, dst):
    import shutil
    src, dst = os.path.abspath(src), os.path.abspath(dst)
    shutil.copyfile(src, dst)
    app = P6.excel()
    try:
        wb = app.Workbooks.Open(dst, UpdateLinks=0)
        app.Calculation = -4135
        patch(wb)
        app.Calculation = -4105
        app.CalculateFull()
        wb.Save()
        wb.Close(False)
    finally:
        P6.quit_excel(app)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
