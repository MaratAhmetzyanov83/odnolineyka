# -*- coding: utf-8 -*-
"""1.07 -> 1.08: данные схемы в «Разбивке по щитам» (шаг 4) и лист «Однолинейка» для SXDRAW (шаг 5).

«Разбивка_по_щитам» A:U — как в 1.05. Справа:
  ввод   V Группа (пусто / гр / узо) · W Автомат (пусто = авто, «-» = нет) · X Дифзащита, мА (пусто = по коду, «нет»)
         Y Аппараты (через «;») · Z Модуль (пусто = по пометке выноски, «-» = без модуля) · AA Канал (пусто / № / «=» / «р»)
         AB Клемма (пусто = авто, «нет») · AC Шина (пусто = основная) · AD Рисовать (пусто = да; вводной узел — нет)
  итог   AE Место · AF Лист · AG Зона · AH QF · AI Автомат · AJ Дифзащита, мА · AK Модуль · AL Модель · AM Канал
         AN Клемма · AO Статус
  служ.  AP:BS (свёрнуты)
Фаза (T) — автобаланс по нагрузке: каждый новый автомат — на наименее загруженную фазу щита,
линии «гр»/«узо» — на фазе строки выше; ручная фаза («Нагрузка_щитов» E) главнее.
Лист «Однолинейка» (скрыт, BC1 = SX_DATA) — те же строки в раскладке, которую читает SXDRAW.
Запуск: py patch_108.py <1.07.xlsx> <1.08.xlsx>
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import patch_106 as P6  # noqa: E402

R0, R1 = 3, 1000
RZ = "Разбивка_по_щитам"
SRC = "'Исходные данные'"

# Модель | Префикс | Каналов | Тип | Ширина, мод. DIN
MODULES = [
    ("ZIOMB24V2", "R", 24, "реле", None), ("ZIO-MB24", "R", 24, "реле", None),
    ("ZIOMB16V3", "R", 16, "реле", None), ("ZIOMB8V3", "R", 8, "реле", None),
    ("ZCL8H230V2", "K", 8, "клапаны 230 В", None), ("ZCL-8HT230", "K", 8, "клапаны 230 В", None),
    ("ZDINDX4", "Di", 4, "диммер", None), ("ZDID64X2", "DA", 2, "DALI", None),
    ("ZIO-RQUAD8", "IN", 8, "входы", None), ("HDL M/S48.1", "IN", None, "входы", None),
    ("SMART-DALI-104", "Di", 4, "диммер DALI", None),
]
CTRL = [("Relay", "ZIOMB24V2"), ("Шторы", "ZIOMB24V2"), ("DIM", "ZDINDX4"), ("DALI", "ZDID64X2"),
        ("Клапан", "ZCL8H230V2")]

INPUT = {  # столбец: (заголовок, примечание)
    "V": ("Группа", "Пусто — свой автомат. «гр» — на автомате строки выше (одна фаза). "
                    "«узо» — свой автомат на шине УЗО строки выше (УЗО впишите в «Аппараты» строки выше)."),
    "W": ("Автомат", "Пусто — подбор (полюса, C, номинал, мА). Своё — как писать на схеме: «1P C16A». «-» — без автомата."),
    "X": ("Дифзащита, мА", "Пусто — по коду («Настройки»). 10 / 30 / 100 / 300 или «нет». В группе — самая строгая."),
    "Y": ("Аппараты", "После автомата, до трёх через «;»: «КМ; БП 24V 480W», «МК-5-1», «УЗО 2P 40A 30мА»."),
    "Z": ("Модуль", "Пусто — по пометке выноски (Relay/DIM/DALI) и коду (шторы, клапаны). Модель — из списка. «-» — без модуля."),
    "AA": ("Канал", "Пусто — следующий свободный. Число — свой канал. «=» — тот же канал, что строка выше. «р» — резерв канала (строка без линии)."),
    "AB": ("Клемма", "Пусто — XT по порядку. «нет» — без клеммы. Своё — например XT5."),
    "AC": ("Шина", "Пусто — основная. Имя второй шины (ИБП, ДГУ) — линия рисуется от неё."),
    "AD": ("Рисовать", "Пусто — да (кроме вводного узла). «нет» — не рисовать. «да» — нарисовать и в вводном узле."),
}
RESULT = {
    "AE": "Место на схеме", "AF": "Лист схемы", "AG": "Зона", "AH": "QF", "AI": "Автомат (итог)",
    "AJ": "Дифзащита гр., мА", "AK": "Модуль (итог)", "AL": "Модель", "AM": "Канал (итог)",
    "AN": "Клемма (итог)", "AO": "Статус",
}

# служебные (A1 для строки 3; «{p}» — строка 2, «{r}» — строка 3)
HELP = [
    ("AP", "_есть линия", '=IF(TRIM(H{r}&"")<>"",1,0)'),
    ("AT", "_строка Исходных", '=IF(AP{r}=0,"",IFERROR(MATCH(H{r},' + SRC + '!$D$3:$D$1000,0),""))'),
    ("AQ", "_начало группы", '=IF(AND(AP{r}=1,TRIM(V{r}&"")<>"гр",TRIM(W{r}&"")<>"-"),1,0)'),
    ("AR", "_QF №", '=IF(AQ{r}=1,COUNTIFS(F$3:F{r},F{r},AQ$3:AQ{r},1),IF(AND(AP{r}=1,TRIM(V{r}&"")="гр",F{p}=F{r},AR{p}<>""),AR{p},""))'),
    ("AU", "_ключ группы", '=IF(AR{r}="","",F{r}&"|"&AR{r})'),
    ("AS", "_дифзащита по коду", '=IF(AT{r}="",0,N(INDEX(' + SRC + '!$BD$3:$BD$1000,AT{r})))'),
    ("BO", "_ток утечки", '=IF(AT{r}="",0,N(INDEX(' + SRC + '!$BE$3:$BE$1000,AT{r})))'),
    ("AV", "_ток группы", '=IF(AQ{r}=1,SUMIFS(K$3:K$1000,AU$3:AU$1000,AU{r}),"")'),
    ("AW", "_In группы", '=IF(AQ{r}=1,LET(u,IFERROR(INDEX(' + SRC + '!$U$3:$U$1000,AT{r}),""),g,IFERROR(VLOOKUP(AV{r},Автомат_номиналы,2,TRUE),""),'
                         'IF(AND(ISNUMBER(u),ISNUMBER(g)),MAX(u,g),u)),"")'),
    ("AX", "_дифзащита группы", '=IF(AQ{r}=1,IF(TRIM(X{r}&"")="нет",0,IF(N(X{r})>0,N(X{r}),IF(TRIM(V{r}&"")="узо",0,'
                                'MAXIFS(AS$3:AS$1000,AU$3:AU$1000,AU{r})))),"")'),
    ("AY", "_утечка группы", '=IF(AQ{r}=1,SUMIFS(BO$3:BO$1000,AU$3:AU$1000,AU{r}),"")'),
    ("BB", "_фаза вручную", '=IF(N(E{r})=0,"",LET(m,INDEX(Нагрузка_щитов!$E$3:$E$1000,E{r})&"",IF(OR(m="",m="0",m="упр"),"",'
                            'IF(m="1,2,3","L1,L2,L3","L"&m))))'),
    ("BC", "_фаза итог", '=IF(BB{r}<>"",BB{r},T{r})'),
    ("BD", "_управление", '=IF(AT{r}="","",LET(mk,UPPER(INDEX(' + SRC + '!$BA$3:$BA$1000,AT{r})&""),cd,INDEX(' + SRC + '!$AH$3:$AH$1000,AT{r})&"",'
                          'IF(ISNUMBER(SEARCH("DALI",mk)),"DALI",IF(ISNUMBER(SEARCH("DIM",mk)),"DIM",'
                          'IF(OR(ISNUMBER(SEARCH("RELAY",mk)),ISNUMBER(SEARCH("РЕЛЕ",mk))),"Relay",'
                          'IF(OR(cd="SH",cd="М"),"Шторы",IF(OR(cd="ЭК",cd="ЭK",LEFT(cd,2)="ТЕ",LEFT(cd,2)="TE"),"Клапан",'
                          'IF(cd="DALI","DALI",""))))))))'),
    ("BE", "_модель", '=IF(TRIM(Z{r}&"")="-","",IF(TRIM(Z{r}&"")<>"",TRIM(Z{r}&""),IF(AND(AP{r}=0,OR(AA{r}&""="р",AA{r}&""="="),F{p}=F{r}),BF{p},'
                      'IF(BD{r}="","",IFERROR(XLOOKUP(BD{r},Модули_умолч_упр,Модули_умолч_мод,""),"")))))'),
    ("BF", "_посл. модель", '=IF(BE{r}<>"",BE{r},IF(F{p}=F{r},BF{p},""))'),
    ("BG", "_каналов на линию", '=IF(BE{r}="",0,IF(AA{r}&""="=",0,IFERROR(MAX(1,N(XLOOKUP(INDEX(' + SRC + '!$AH$3:$AH$1000,AT{r})&"",'
                                'Коды_доп_код,INDEX(Коды_доп,0,5),1))),1)))'),
    ("BH", "_каналов модели", '=IF(BE{r}="",0,IFERROR(N(XLOOKUP(BE{r},Модули_спр_мод,Модули_спр_кан,0)),0))'),
    ("BJ", "_каналов модели нарастающим", '=IF(BE{r}="","",SUMIFS(BG$3:BG{r},F$3:F{r},F{r},BE$3:BE{r},BE{r}))'),
    ("BI", "_новый модуль", '=IF(OR(BE{r}="",BG{r}=0),0,IF(MOD(BJ{r}-BG{r},IF(BH{r}=0,999,BH{r}))=0,1,0))'),
    ("BK", "_префикс", '=IF(BE{r}="","",IFERROR(XLOOKUP(BE{r},Модули_спр_мод,Модули_спр_преф,"R")&"","R"))'),
    ("BM", "_модуль (имя)", '=IF(BE{r}="","",BK{r}&"."&(INT((MAX(1,BJ{r})-1)/IF(BH{r}=0,999,BH{r}))+1))'),
    ("BN", "_авто клемма", '=IF(AND(AP{r}=1,TRIM(AB{r}&"")=""),1,0)'),
    ("BL", "_узел щита, мест", '=IF(F{r}="","",IFERROR(LET(v,XLOOKUP(F{r},Щиты_имя,Щиты_узел,""),IF(N(v)>0,N(v),Мест_на_лист)),Мест_на_лист))'),
    ("BP", "_рисовать", '=IF(F{r}="",0,IF(TRIM(AD{r}&"")="нет",0,IF(AND(AG{r}="вводной узел",TRIM(AD{r}&"")<>"да"),0,1)))'),
    ("BQ", "_код", '=IF(AT{r}="","",INDEX(' + SRC + '!$AH$3:$AH$1000,AT{r})&"")'),
    ("BR", "_УГО", '=IF(BQ{r}="","",IFERROR(XLOOKUP(BQ{r},Коды_доп_код,INDEX(Коды_доп,0,2),"")&"",""))'),
    ("BT", "_Р для баланса фаз", '=IF(OR(IFERROR(N(Q{r}),1)=3,LEFT(BQ{r},4)="Ввод",'
                                 'ISNUMBER(SEARCH("!",IFERROR(FORMULATEXT(INDEX(' + SRC + '!$E$3:$E$1000,AT{r})),"")))),0,IFERROR(N(L{r}),0))'),
    ("BS", "_цепь", '=IF(AP{r}=0,"",TEXTJOIN("|",TRUE,AH{r},SUBSTITUTE(SUBSTITUTE(TRIM(Y{r}&""),"; ","|"),";","|"),'
                    'IF(AK{r}<>"",AK{r}&" к."&AM{r},""),AN{r}))'),
]
AO_PARTS = [
    'IF(AND(AP{r}=1,AT{r}=""),"номера нет в «Исходных данных»","")',
    'IF(AND(AP{r}=1,TRIM(V{r}&"")="гр",AR{r}=""),"«гр»: нет автомата выше","")',
    'IF(AND(AP{r}=1,BC{r}=""),"нет фазы","")',
    'IF(AND(AQ{r}=1,AG{r}<>"вводной узел",ISNUMBER(AV{r}),ISNUMBER(AW{r}),AV{r}>AW{r}),"ток группы "&FIXED(AV{r},1)&" А > "&AW{r}&" А","")',
    'IF(AND(AQ{r}=1,BQ{r}="R",N(AX{r})=0),"розетка без дифзащиты","")',
    'IF(AND(AQ{r}=1,N(AX{r})>0,N(AY{r})>N(AX{r})/3),"ток утечки группы "&FIXED(AY{r},1)&" мА > 1/3 уставки","")',
    'IF(AND(ISNUMBER(AA{r}),BH{r}>0,N(AA{r})>BH{r}),"канала "&AA{r}&" нет у "&BE{r}&" ("&BH{r}&")","")',
    'IF(AND(BE{r}<>"",BH{r}=0),"нет числа каналов у модели "&BE{r}&" («Настройки»)","")',
    'IF(AND(AN{r}<>"",COUNTIFS(F$3:F$1000,F{r},AN$3:AN$1000,AN{r})>1),"дубль клеммы "&AN{r},"")',
    'IF(AND(AP{r}=1,ISERROR(P{r})),"кабель не считается: проверьте код и данные в «Исходных данных»","")',
    'IF(TRIM(V{r}&"")="узо","УЗО на группу: на схеме — вручную","")',
]
RESULT_F = {
    "AE": '=IF(F{r}="","",C{r})',
    "AF": '=IF(F{r}="","",INT((C{r}-1)/Мест_на_лист)+1)',
    "AG": '=IF(F{r}="","",IF(C{r}<=BL{r},"вводной узел",""))',
    "AH": '=IF(AQ{r}=1,IF(N(AX{r})>0,"QFD","QF")&AR{r},IF(AND(AR{r}<>"",F{p}=F{r}),AH{p},""))',
    "AI": '=IF(AQ{r}<>1,"",IF(TRIM(W{r}&"")<>"",W{r},IF(AW{r}="","",IF(Q{r}=3,IF(N(AX{r})>0,"4P","3P"),IF(N(AX{r})>0,"2P","1P"))'
          '&" C"&AW{r}&"A"&IF(N(AX{r})>0," "&AX{r}&"мА",""))))',
    "AJ": '=IF(AQ{r}<>1,"",IF(N(AX{r})>0,AX{r},""))',
    "AK": '=IF(BE{r}="","",BM{r})',
    "AL": '=BE{r}',
    "AM": '=IF(BE{r}="","",IF(ISNUMBER(AA{r}),AA{r},IF(AA{r}&""="=",MOD(MAX(1,BJ{r})-1,IF(BH{r}=0,999,BH{r}))+1,MOD(BJ{r}-BG{r},IF(BH{r}=0,999,BH{r}))+1)))',
    "AN": '=IF(AP{r}=0,"",IF(TRIM(AB{r}&"")="нет","",IF(TRIM(AB{r}&"")<>"",AB{r},"XT"&COUNTIFS(F$3:F{r},F{r},BN$3:BN{r},1))))',
    "AO": '=IF(AND(AP{r}=0,AA{r}&""=""),"",TEXTJOIN("; ",TRUE,' + ",".join('IFERROR(' + c + ',"")' for c in AO_PARTS) + '))',
}
# фаза: автобаланс по нагрузке (T), заменяет круговую 1.05
T_PHASE = ('=IF(TRIM(H{r}&"")="","",IF(IFERROR(N(Q{r}),1)=3,"L1,L2,L3",IF(AND(OR(TRIM(V{r}&"")="гр",TRIM(V{r}&"")="узо"),F{p}=F{r},BC{p}<>""),BC{p},'
           'LET(ka,SUMIFS(BT$2:BT{p},F$2:F{p},F{r},BC$2:BC{p},"L1"),kb,SUMIFS(BT$2:BT{p},F$2:F{p},F{r},BC$2:BC{p},"L2"),'
           'kc,SUMIFS(BT$2:BT{p},F$2:F{p},F{r},BC$2:BC{p},"L3"),IF(AND(ka<=kb,ka<=kc),"L1",IF(kb<=kc,"L2","L3"))))))')

# лист «Однолинейка» (SXDRAW): столбец -> формула (ссылка на «Разбивку» той же строки)
R_ = RZ + "!"
ONE = {
    "A": '=IF({R}F{r}="","",{R}F{r})',
    "B": '=IF({R}F{r}="","",{R}C{r})',
    "C": '={R}AF{r}',
    "D": '=TRIM({R}AC{r}&"")',
    "E": '=IF({R}AQ{r}=1,"авто","")',
    "F": '={R}AJ{r}',
    "G": '=IF(OR({R}AP{r}=1,{R}AM{r}<>"",{R}AN{r}<>""),{R}AH{r},"")',
    "H": '={R}AI{r}',
    "I": '=IF({R}AQ{r}<>1,"",IFERROR(TRIM(INDEX(TEXTSPLIT({R}Y{r}&"",";"),1)),""))',
    "J": '=IF({R}AQ{r}<>1,"",IFERROR(TRIM(INDEX(TEXTSPLIT({R}Y{r}&"",";"),2)),""))',
    "K": '=IF({R}AQ{r}<>1,"",IFERROR(TRIM(INDEX(TEXTSPLIT({R}Y{r}&"",";"),3)),""))',
    "L": '=TRIM({R}Z{r}&"")',
    "M": '={R}AK{r}',
    "N": '=TRIM({R}AA{r}&"")',
    "O": '={R}AM{r}',
    "P": '={R}AN{r}',
    "Q": '=IF({R}AP{r}=0,"",IFERROR(SUBSTITUTE({R}P{r}&"",CHAR(10)," "),""))',
    "R": '=TRIM({R}H{r}&"")',
    "S": '=IF({R}AP{r}=0,"",IFERROR({R}J{r}&"",""))',
    "T": '=IF({R}AP{r}=0,"",IFERROR(--{R}L{r},""))',
    "U": '=IF({R}AP{r}=0,"",IFERROR(--{R}K{r},""))',
    "V": '=IF({R}AP{r}=0,"",IFERROR({R}M{r}&"",""))',
    "Y": '=IF({R}BP{r}=0,"вручную","")',
    "AE": '={R}AO{r}',
    "AG": '=IF({R}AP{r}=1,{R}AH{r},IF(AND({R}AM{r}<>"",{R}AH{r}<>""),{R}AH{r},""))',
    "AH": '={R}AK{r}',
    "AI": '={R}AL{r}',
    "AJ": '=IF({R}AQ{r}=1,{R}AH{r},"")',
    "AL": '={R}BQ{r}',
    "AN": '=IF(AND({R}AP{r}=1,{R}F{p}={R}F{r},TRIM({R}H{p}&"")=TRIM({R}H{r}&"")),1,0)',
    "AR": '=IF(IFERROR(N({R}Q{r}),1)=3,3,1)',
    "AY": '=IF({R}AN{r}<>"",1,0)',
    "BD": '={R}BP{r}',
    "BE": '=IF(D{r}<>"",1,0)',
    "BF": '=IF(LEFT(AJ{r},3)="QFD",1,0)',
    "BG": '=IF(N{r}="=","=",IF(N{r}="р","R",""))',
    "BH": None, "BI": None, "BJ": None,   # тип аппарата — ниже
    "BK": '=IF(R{r}="","",IF({R}BR{r}="","BOX",IFERROR(XLOOKUP({R}BR{r},INDEX(УГО_блоки,0,1),INDEX(УГО_блоки,0,3),"BOX")&"","BOX")))',
    "BL": '=IF(R{r}="",0,1)',
    "BY": '={R}BR{r}',
}
DEVT = ('=IF({d}{r}="","",IF(OR(LEFT({d}{r},1)="T",LEFT({d}{r},1)="Т"),"T",IF(LEFT(UPPER({d}{r}),2)="БП","PS",'
        'IF(OR(LEFT(UPPER({d}{r}),2)="МК",LEFT(UPPER({d}{r}),2)="MK"),"MK",IF(OR(LEFT(UPPER({d}{r}),2)="КМ",'
        'LEFT(UPPER({d}{r}),2)="KM"),"KM","X")))))')

LOG_108 = ("Схема (шаги 4–5).\n"
           "1. «Разбивка_по_щитам» V:AD — ввод для схемы: группа (гр/узо), автомат, дифзащита, аппараты, модуль, канал, "
           "клемма, шина, рисовать. Пусто — подбор таблицы. AE:AO — итог: место, лист, зона, QF, автомат, модуль, канал, клемма, статус.\n"
           "2. Фаза — автобаланс по нагрузке: новый автомат — на наименее загруженную фазу щита, «гр»/«узо» — на фазе строки выше; "
           "ручная фаза («Нагрузка_щитов» E) главнее.\n"
           "3. Модули — по пометке выноски (Relay, DIM, DALI) и коду (шторы, клапаны): модели по умолчанию — «Настройки», "
           "каталог модулей — BB:BF. Новый модуль — когда кончились каналы.\n"
           "4. Вводной узел: первые места щита («Диапазоны» H, пусто = 17) не рисуются, SXDRAW начинает со следующего листа.\n"
           "5. Скрытый лист «Однолинейка» — данные для SXDRAW (не править).")


def fmt(s, row=R0):
    return s.replace("{R}", R_).replace("{r}", str(row)).replace("{p}", str(row - 1))


def note(cell, text):
    try:
        if cell.Comment is not None:
            cell.Comment.Delete()
    except Exception:
        pass
    cell.AddComment(text)


def validation(rng, items):
    rng.Validation.Delete()
    rng.Validation.Add(3, 1, 1, ",".join(items))   # xlValidateList
    rng.Validation.IgnoreBlank = True
    rng.Validation.ShowError = False                # своё значение можно вписать


def patch(wb, log=print):
    sp = wb.Worksheets("Справочная")
    rz = wb.Worksheets(RZ)
    if str(rz.Range("V2").Value or "") == "Группа":
        log("1.08 уже применён — пропуск")
        return

    # --- Справочная: модули ---
    sp.Range("BB2").Value = "МОДУЛИ (каталог для схемы и спецификации)"
    sp.Range("BB2").Font.Bold = True
    hdr = ("Модель", "Префикс на схеме", "Каналов", "Тип", "Ширина, мод. DIN")
    sp.Range("BB3:BF3").Value = (hdr,)
    sp.Range("BB3:BF3").Font.Bold = True
    for i, m in enumerate(MODULES):
        sp.Range(f"BB{4+i}:BF{4+i}").Value = (tuple("" if x is None else x for x in m),)
    P6.add_name(wb, "Модули_спр", "=Справочная!$BB$4:$BF$60")
    P6.add_name(wb, "Модули_спр_мод", "=Справочная!$BB$4:$BB$60")
    P6.add_name(wb, "Модули_спр_преф", "=Справочная!$BC$4:$BC$60")
    P6.add_name(wb, "Модули_спр_кан", "=Справочная!$BD$4:$BD$60")
    sp.Range("BH2").Value = "МОДЕЛЬ ПО УМОЛЧАНИЮ (управление → модель)"
    sp.Range("BH2").Font.Bold = True
    sp.Range("BH3:BI3").Value = (("Управление", "Модель"),)
    sp.Range("BH3:BI3").Font.Bold = True
    for i, c in enumerate(CTRL):
        sp.Range(f"BH{4+i}:BI{4+i}").Value = (c,)
    P6.add_name(wb, "Модули_умолч_упр", "=Справочная!$BH$4:$BH$20")
    P6.add_name(wb, "Модули_умолч_мод", "=Справочная!$BI$4:$BI$20")
    log("Справочная: модули BB:BF, модели по умолчанию BH:BI")

    # --- Разбивка: заголовки ---
    for col, (h, n) in INPUT.items():
        c = rz.Range(f"{col}2")
        c.Value = h
        note(c, n)
    for col, h in RESULT.items():
        rz.Range(f"{col}2").Value = h
    for col, h, _ in HELP:
        rz.Range(f"{col}2").Value = h
    rz.Range("V2:BT2").Font.Bold = True
    rz.Range("V2:BT2").WrapText = True
    for c in ("V", "W", "Y", "Z", "AB", "AC", "AD"):
        rz.Range(f"{c}{R0}:{c}{R1}").NumberFormat = "@"
    # служебные формулы (порядок важен только для чтения; Excel считает сам)
    for col, f in [(c, f) for c, _, f in HELP] + list(RESULT_F.items()):
        try:
            rz.Range(f"{col}{R0}:{col}{R1}").Formula2 = fmt(f)
        except Exception as e:
            raise RuntimeError(f"формула {col}: {fmt(f)}") from e
    # фаза: T — автобаланс
    rz.Range(f"T{R0}:T{R1}").Formula2 = fmt(T_PHASE)
    note(rz.Range("T2"), "1.08: автобаланс по нагрузке — новый автомат на наименее загруженную фазу щита, "
                         "«гр»/«узо» — на фазе строки выше. Ручная фаза — «Нагрузка_щитов» E.")
    # оформление
    rz.Range(f"AE{R0}:AO{R1}").Interior.Color = 0xF2F2F2
    rz.Range("V2:AD2").Interior.Color = 0xDDEBF7
    rz.Range("AE2:AO2").Interior.Color = 0xE7E6E6
    for col, w in (("V", 6), ("W", 12), ("X", 9), ("Y", 20), ("Z", 12), ("AA", 7), ("AB", 7), ("AC", 8), ("AD", 8),
                   ("AE", 7), ("AF", 6), ("AG", 12), ("AH", 7), ("AI", 16), ("AJ", 9), ("AK", 8), ("AL", 12),
                   ("AM", 7), ("AN", 7), ("AO", 40)):
        rz.Columns(col).ColumnWidth = w
    validation(rz.Range(f"V{R0}:V{R1}"), ["гр", "узо"])
    validation(rz.Range(f"X{R0}:X{R1}"), ["10", "30", "100", "300", "нет"])
    validation(rz.Range(f"AB{R0}:AB{R1}"), ["нет"])
    validation(rz.Range(f"AD{R0}:AD{R1}"), ["да", "нет"])
    try:
        rz.Range(f"Z{R0}:Z{R1}").Validation.Delete()
        rz.Range(f"Z{R0}:Z{R1}").Validation.Add(3, 1, 1, "=Модули_спр_мод")
        rz.Range(f"Z{R0}:Z{R1}").Validation.ShowError = False
    except Exception:
        pass
    # границы листов схемы и щитов, вводной узел
    rng = rz.Range(f"A{R0}:AO{R1}")
    for f1, kind in ((f'=AND($F{R0}<>"",MOD($C{R0},17)=0)', "sheet"), (f'=AND($F{R0}<>"",$F{R0}<>$F{R0+1})', "board"),
                     (f'=$AG{R0}="вводной узел"', "node")):
        try:
            fc = rng.FormatConditions.Add(2, 0, f1)
        except Exception:
            continue
        if kind == "node":
            fc.Interior.Color = 0xEFEFEF
        else:
            b = fc.Borders(4)   # xlEdgeBottom
            b.LineStyle = 1
            b.Color = 0x000000 if kind == "board" else 0x7F7F7F
    try:
        fc = rz.Range(f"AO{R0}:AO{R1}").FormatConditions.Add(2, 0, f'=LEN($AO{R0})>0')
        fc.Font.Color = 0x0000C0
    except Exception:
        pass
    # удобство: номер линии на экране при прокрутке вправо, замечание видно у номера, хвосты текста U не лезут в V
    try:
        rz.Activate()
        win = wb.Application.ActiveWindow
        win.FreezePanes = False
        win.ScrollRow = 1
        win.ScrollColumn = 1
        rz.Range("I3").Select()
        win.FreezePanes = True
    except Exception:
        pass
    rz.Range(f"U{R0}:U{R1}").ShrinkToFit = True
    try:
        fc = rz.Range(f"H{R0}:H{R1}").FormatConditions.Add(2, 0, f'=LEN($AO{R0})>0')
        fc.Interior.Color = 0x9CC7FF   # оранжевый: в «Статусе» (AO) есть замечание
        fc.StopIfTrue = False
    except Exception:
        pass
    note(rz.Range("H2"), "Оранжевый номер — в «Статусе» (столбец AO) есть замечание.")
    # служебные столбцы свернуть (плюсик над столбцами разворачивает):
    # 1.05 — A:E нумерация, I промежуточный расчёт, N повтор номера, O УГО, Q:U фазность и промежуточные;
    # 1.08 — AP:BT. Ввод для схемы (V) оказывается сразу за кабелем (P).
    for rg in ("A:E", "I:I", "N:O", "Q:U", "AP:BT"):
        try:
            rz.Range(rg).Columns.Group()
        except Exception:
            pass
    try:
        rz.Outline.SummaryColumn = 0      # плюсик слева от группы
        rz.Outline.ShowLevels(0, 1)
    except Exception:
        for rg in ("A:E", "I:I", "N:O", "Q:U", "AP:BT"):
            rz.Range(rg).EntireColumn.Hidden = True
    log("Разбивка: V:AO, служебные AP:BS, фаза T")

    # --- лист «Однолинейка» для SXDRAW ---
    one = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    one.Name = "Однолинейка"
    one.Range("A1").Value = "Данные для SXDRAW (формулы из «Разбивки по щитам») — не править"
    one.Range("BC1").Value = "SX_DATA"
    heads = {"A": "Щит", "B": "№ места", "C": "Лист", "D": "Шина", "H": "Автомат", "I": "Аппарат 1", "J": "Аппарат 2",
             "K": "Аппарат 3", "O": "Канал", "P": "Клемма", "Q": "Кабель", "R": "Линия", "S": "Наименование",
             "T": "Р, кВт", "U": "Ток", "V": "Фаза", "AE": "Статус", "AG": "QF гр.", "AH": "Модуль гр.",
             "AI": "Модель гр.", "AJ": "QF нач.", "AN": "в кабель выше", "AR": "фаз", "AY": "клемма",
             "BD": "draw", "BE": "bus2", "BF": "rcd", "BG": "chtype", "BH": "dev1", "BI": "dev2", "BJ": "dev3",
             "BK": "sym", "BL": "hasline", "BY": "УГО (рус)"}
    for c, h in heads.items():
        one.Range(f"{c}2").Value = h
    for col, f in ONE.items():
        if f is None:
            continue
        one.Range(f"{col}{R0}:{col}{R1}").Formula2 = fmt(f)
    for col, d in (("BH", "I"), ("BI", "J"), ("BJ", "K")):
        one.Range(f"{col}{R0}:{col}{R1}").Formula2 = fmt(DEVT.replace("{d}", d))
    one.Visible = 0   # xlSheetHidden
    log("Лист «Однолинейка» (скрыт, BC1 = SX_DATA)")

    # спецификация по щиту: выбор щита; нормы расхода для надстройки
    spec = wb.Worksheets("Спецификация")
    spec.Range("G1").Value = "Щит (пусто — все):"
    spec.Range("G1").HorizontalAlignment = -4152
    spec.Range("H1").Interior.Color = 0xDDEBF7
    try:
        spec.Range("H1").Validation.Delete()
        spec.Range("H1").Validation.Add(3, 1, 1, "=Щиты_имя")
        spec.Range("H1").Validation.ShowError = False
    except Exception:
        pass
    note(spec.Range("H1"), "Надстройка «Однолинейка» → «Обновить спецификацию» заполняет лист по этому щиту "
                           "(пусто — по всем). Строки с «вручную» в «Примечании» сохраняются.")
    sp.Range("AL29").Value = "Клипсы на 1 м гофры, шт"
    sp.Range("AM29").Value = 2
    sp.Range("AL30").Value = "Запас кабеля в спецификации (доля)"
    sp.Range("AM30").Value = 0.2
    sp.Range("AL31").Value = "Кабель и гофра в спецификации — округление вверх, м"
    sp.Range("AM31").Value = 10
    P6.add_name(wb, "Клипсы_на_м", "=Справочная!$AM$29")
    P6.add_name(wb, "Запас_кабеля", "=Справочная!$AM$30")
    P6.add_name(wb, "Округл_м", "=Справочная!$AM$31")

    kr = wb.Worksheets("Краткое руководство")
    kr.Range("C1").Value = "Версия таблицы 1.08"
    r = 2
    while kr.Range(f"E{r}").Value not in (None, ""):
        r += 1
    kr.Range(f"E{r}").Value = "1.08"
    kr.Range(f"F{r}").Value = LOG_108
    kr.Range(f"F{r}").WrapText = True
    log("Краткое руководство: 1.08")


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
