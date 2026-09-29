"""v3.11: связи однолинеек (вышестоящий ↔ нижестоящий щит), питание щита, ТКЗ по линиям, К прокладки на линию, данные для ПЗ.

Новые листы:
  «Питание_щита»      — импорт от вышестоящего щита (Вставить связь) или питающая линия вручную; ΔU и Zпетли до шин,
                         Iкз min/max на шинах, проверки питающей линии; ЭКСПОРТ этого щита для вышестоящего.
  «Нижестоящие_щиты»  — подключение других однолинеек: строка на линию-фидер, импорт Руст/Руст быт/сил/Рр сил/пост/cos/Iр
                         (Вставить связь из «Питание_щита» нижестоящего файла), экспорт ΔU/Zпетли/Iкз/In для него.
Линия, у которой есть связь с нижестоящим файлом, получает группу «Щит»: её Руст быт идёт в общий Руст быт (Кс по таблице
пересчитывается по сумме), Рр сил — в Рр сил (Ко применяется один раз наверху), постоянная — в постоянную.
По линиям («Исходные данные» BH:BM): К прокладки на линию, Zпетли и Iкз min/max в конце, Iкз/In, отключение ≤ 0,4 с.
«Расчёт_нагрузок»: блок «Основные показатели для ПЗ».

Запуск: python3 patch9.py  (out10/ → out11/)
"""
import os, re, sys, html
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import patch7 as P7
import patch8 as P8
from patch7 import X, rd, wr, ct, cn, cf, ce, bulk, set_cols, add_dxf, add_cf, add_dv, add_name, unshare, esc

OUT = os.path.join(HERE, "out11"); os.makedirs(OUT, exist_ok=True)
ISX = "'Исходные данные'!"
NG = "Нагрузка_щитов!"
OL = "Однолинейка!"
RZ = "Разбивка_по_щитам!"
PS = "Питание_щита!"
NS = "Нижестоящие_щиты!"
BLK = [4, 13, 22]
UF = "(Uн_лин*1000/SQRT(3))"      # фазное, В
NCH = 30                          # строк нижестоящих щитов


def new_sheet(WD, order, name, after, cells, merges, widths, dv_xml, tab, freeze_row, print_area=None, heights=None):
    wbp = WD + "/xl/workbook.xml"; wb = rd(wbp)
    rp = WD + "/xl/_rels/workbook.xml.rels"; rels = rd(rp)
    ctp = WD + "/[Content_Types].xml"
    rid = "rId%d" % (max(int(x) for x in re.findall(r'Id="rId(\d+)"', rels)) + 1)
    sn = max(int(x) for x in re.findall(r"worksheets/sheet(\d+)\.xml", rels)) + 1
    part = f"xl/worksheets/sheet{sn}.xml"
    sid = max(int(x) for x in re.findall(r'sheetId="(\d+)"', wb)) + 1
    names = re.findall(r'<sheet name="([^"]+)"', wb)
    pos = names.index(after) + 1
    tags = re.findall(r"<sheet [^>]*/>", wb)
    tags.insert(pos, f'<sheet name="{name}" sheetId="{sid}" r:id="{rid}"/>')
    wb = re.sub(r"<sheets>.*?</sheets>", "<sheets>" + "".join(tags) + "</sheets>", wb, flags=re.S)
    wb = re.sub(r'localSheetId="(\d+)"', lambda m: f'localSheetId="{int(m.group(1)) + (1 if int(m.group(1)) >= pos else 0)}"', wb)
    wb = re.sub(r'activeTab="(\d+)"', lambda m: f'activeTab="{int(m.group(1)) + (1 if int(m.group(1)) >= pos else 0)}"', wb)
    if print_area:
        k = wb.find("</definedNames>")
        wb = wb[:k] + f'<definedName name="_xlnm.Print_Area" localSheetId="{pos}">{print_area}</definedName>' + wb[k:]
    wr(wbp, wb)
    wr(rp, rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{sn}.xml"/></Relationships>'))
    wr(ctp, rd(ctp).replace("</Types>", f'<Override PartName="/{part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'))
    order.append(part)
    byrow = {}
    for ref_, xml in cells.items():
        r = int(re.search(r"\d+", ref_).group(0)); byrow.setdefault(r, {})[re.match(r"[A-Z]+", ref_).group(0)] = xml
    heights = heights or {}
    rows = []
    for r in sorted(byrow):
        ht = f' ht="{heights[r]}" customHeight="1"' if r in heights else ""
        rows.append(f'<row r="{r}"{ht}>' + "".join(byrow[r][c_] for c_ in sorted(byrow[r], key=X.col2n)) + "</row>")
    cols = "".join(f'<col min="{X.col2n(a)}" max="{X.col2n(a)}" width="{w}" customWidth="1"/>' for a, w in widths.items())
    mc = f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in merges) + "</mergeCells>" if merges else ""
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           f'<sheetPr><tabColor rgb="{tab}"/><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1"/>'
           f'<sheetViews><sheetView workbookViewId="0"><pane ySplit="{freeze_row}" topLeftCell="A{freeze_row + 1}" activePane="bottomLeft" state="frozen"/>'
           '<selection pane="bottomLeft" activeCell="C6" sqref="C6"/></sheetView></sheetViews>'
           '<sheetFormatPr defaultRowHeight="14.4"/>'
           f"<cols>{cols}</cols><sheetData>{''.join(rows)}</sheetData>{mc}{dv_xml}"
           '<pageMargins left="0.5" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
           '<pageSetup paperSize="9" orientation="landscape" fitToHeight="0"/></worksheet>')
    wr(os.path.join(WD, part), xml)


def dvx(items):
    body = "".join(f'<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="{sq}"><formula1>{esc(f1)}</formula1></dataValidation>'
                   for sq, f1 in items)
    return f'<dataValidations count="{len(items)}">{body}</dataValidations>'


# =====================================================================================  «Питание_щита»
def sheet_supply(WD, order, S):
    cells = {}; merges = []
    def T(ref, v, s): cells[ref] = ct(ref, v, S[s])
    def N(ref, v, s): cells[ref] = cn(ref, v, S[s])
    def F(ref, f, s): cells[ref] = cf(ref, f, S[s])
    def E(ref, s): cells[ref] = ce(ref, S[s])
    T("B1", "ПИТАНИЕ ЩИТА: потери напряжения и ток КЗ до шин, связь с однолинейкой вышестоящего щита", "T")
    T("B2", "Жёлтое — ввод. Если питающий щит тоже в такой однолинейке: откройте его файл → лист «Нижестоящие_щиты» → выделите K:N строки этого щита → Копировать → "
            "здесь в разделе 1 выделите C:F нужной строки → Специальная вставка → «Вставить связь». Иначе заполните раздел 2. "
            "Для своего вышестоящего щита скопируйте строку раздела 4 и вставьте её связью в его «Нижестоящие_щиты» (столбцы D:J).", "NOTE")
    merges.append("B2:I2")
    # 1. импорт
    T("B4", "1. ОТ ВЫШЕСТОЯЩЕГО ЩИТА (Вставить связь)", "BT"); merges.append("B4:G4")
    for col in "CDEFG": E(f"{col}4", "BT")
    for col, t in zip("BCDEFG", ["Щит", "ΔU от ВРУ до шин, %", "Zпетли ф–0 на шинах, Ом", "Iкз max (3ф) на шинах, кА",
                                  "In аппарата питающей линии, А", "Файл вышестоящего (для себя)"]):
        T(f"{col}5", t, "H")
    for k in range(3):
        r = 6 + k
        F(f"B{r}", f'IFERROR(Диапазоны!B{4 + k}&"","")', "TXT")
        for col in "CDEFG": E(f"{col}{r}", "IN")
    # 2. вручную
    T("B10", "2. ПИТАЮЩАЯ ЛИНИЯ ВРУЧНУЮ (если раздел 1 пуст)", "BT"); merges.append("B10:G10")
    for col in "CDEFG": E(f"{col}10", "BT")
    T("B11", "Параметр", "H"); T("F11", "ед.", "H"); T("G11", "Пояснение", "H")
    for k, col in enumerate("CDE"): F(f"{col}11", f"$B${6 + k}", "H")
    INP = [(12, "Питается от (ВРУ, ГРЩ, щит … или «ТП»)", "", "«ТП» — этот щит и есть ВРУ: потери по СП 256 п. 8.23 считаются от его шин (0 %)"),
           (13, "ΔU от ВРУ до шин питающего щита", "%", "из его файла («Питание_щита», стр. 33); для ВРУ — пусто"),
           (14, "Iкз min (1ф) на шинах питающего щита", "А", "из его файла (стр. 35) или от сетевой организации"),
           (15, "Iкз max (3ф) на шинах питающего щита", "кА", "из его файла (стр. 36) или от сетевой организации"),
           (16, "Если от ТП: мощность трансформатора", "кВА", "если Iкз неизвестны; uк = 5,5 %"),
           (17, "Если от ТП: схема трансформатора", "", "Д/Yн или Y/Yн (для Y/Yн сопротивление петли ≈ ×3)"),
           (18, "Питающий кабель: марка", "", "для ПЗ"),
           (19, "Сечение фазной жилы", "мм²", ""),
           (20, "Жил в кабеле (5 / 4 / 1 — одножильные)", "", "пусто = 5"),
           (21, "Кабелей параллельно", "шт.", "пусто = 1"),
           (22, "Длина", "м", ""),
           (23, "Материал жил", "", "медь / алюминий; пусто = медь"),
           (24, "Аппарат питающей линии: In", "А", ""),
           (25, "Характеристика аппарата", "", "B / C / D; пусто = C")]
    for r, lab, u, hint in INP:
        T(f"B{r}", lab, "TXT"); T(f"F{r}", u, "INT"); T(f"G{r}", hint, "TXT")
        for col in "CDE": E(f"{col}{r}", "IN")
    # 3. результат
    T("B27", "3. РЕЗУЛЬТАТ ДЛЯ ЭТОГО ЩИТА", "BT"); merges.append("B27:G27")
    for col in "CDEFG": E(f"{col}27", "BT")
    T("B28", "Показатель", "H"); T("F28", "ед.", "H"); T("G28", "Норма / пояснение", "H")
    for k, col in enumerate("CDE"): F(f"{col}28", f"$B${6 + k}", "H")
    RES = [(29, "Рр щита", "кВт", "«Нагрузка_щитов», весь щит"), (30, "Iр щита", "А", "по наиболее нагруженной фазе"),
           (31, "cos φ", "", ""), (32, "ΔU питающей линии", "%", "ГОСТ Р 50571.5.52 прил. G"),
           (33, "ΔU от ВРУ до шин этого щита → в линии", "%", "прибавляется к ΔU каждой линии; норма СП 256 п. 8.23: свет ≤ 3 %, прочие ≤ 4 % от ВРУ"),
           (34, "Zпетли фаза–ноль на шинах → в линии", "Ом", "по нему считается Iкз в конце каждой линии"),
           (35, "Iкз min (1ф) на шинах", "А", "c_min = 0,95, проводники нагретые, + переходные сопротивления"),
           (36, "Iкз max (3ф) на шинах", "кА", "сравнить с отключающей способностью аппаратов щита"),
           (37, "Питающий кабель: Iдоп", "А", "ПУЭ табл. 1.3.6 × 0,92 (4–5 жил) × К прокл. × число кабелей"),
           (38, "Проверка аппарата и кабеля", "", "Iр ≤ In ≤ Iдоп (ГОСТ Р 50571.4.43)"),
           (39, "Отключение КЗ на шинах (питающая линия)", "", "ПУЭ 1.7.79: для распределительных цепей ≤ 5 с"),
           (40, "Термическая стойкость питающего кабеля", "", "t доп = (k·S / Iкз max)², k = 115 (Cu, ПВХ) / 76 (Al)"),
           (41, "Селективность с отходящими", "", "ориентир In ≥ 1,6·In отходящего; подтвердить картами селективности производителя")]
    for r, lab, u, hint in RES:
        T(f"B{r}", lab, "BT" if r in (33, 34) else "TXT"); T(f"F{r}", u, "INT"); T(f"G{r}", hint, "TXT")
    # служебные строки 57–60
    T("B56", "служебные", "G")
    for k, X_ in enumerate("CDE"):
        ir = 6 + k; rb = BLK[k]
        c_ = lambda r: f"{X_}{r}"
        tp = f'LEFT(UPPER(TRIM({X_}12)),2)="ТП"'
        n = f"MAX(1,N({X_}21))"; S_ = f"N({X_}19)"; L_ = f"N({X_}22)"; al = f'{X_}23="алюминий"'
        cores = f"IF(N({X_}20)=0,5,{X_}20)"
        # служебные: Z на шинах питающего (1ф), Z3 на шинах питающего, In действующий
        F(f"{X_}57", f'IF(N({X_}14)>0,c_min*{UF}/{X_}14,IF(AND({tp},N({X_}16)>0),0.055*Uн_лин^2/({X_}16/1000)*IF({X_}17="Y/Yн",3,1),""))', "GN")
        F(f"{X_}58", f'IF(N({X_}15)>0,1.05*{UF}/({X_}15*1000),IF(AND({tp},N({X_}16)>0),0.055*Uн_лин^2/({X_}16/1000),""))', "GN")
        F(f"{X_}59", f'IF(ISNUMBER($F${ir}),$F${ir},N({X_}24))', "GN")
        F(f"{X_}29", f'IFERROR({NG}$V${rb + 2}*1,"")', "N2")
        F(f"{X_}30", f'IFERROR({NG}$V${rb + 3}*1,"")', "N2")
        F(f"{X_}31", f'IFERROR({NG}$V${rb + 4}*1,0.95)', "N2")
        F(f"{X_}32", f'IF(OR({S_}=0,{L_}=0,NOT(ISNUMBER({X_}30))),"",(IF({al},0.036,0.0225)*{L_}/({S_}*{n})*{X_}31+0.00008*{L_}/{n}*SIN(ACOS({X_}31)))*{X_}30/{UF}*100)', "N2")
        F(f"{X_}33", f'IF(ISNUMBER($C${ir}),$C${ir},IF({tp},0,N({X_}13)+N({X_}32)))', "BN2")
        F(f"{X_}34", f'IF(ISNUMBER($D${ir}),$D${ir}+R_пер,IF({X_}57="","",{X_}57+IF({S_}>0,2*IF({al},ρкз_Al,ρкз_Cu)*{L_}/({S_}*{n}),0)+R_пер))', "BN2")
        F(f"{X_}35", f'IF(ISNUMBER({X_}34),c_min*{UF}/{X_}34,"")', "N2")
        F(f"{X_}36", f'IF(ISNUMBER($E${ir}),$E${ir},IF({X_}58="","",1.05*{UF}/({X_}58+IF({S_}>0,IF({al},0.028,0.0175)*{L_}/({S_}*{n}),0))/1000))', "N2")
        F(f"{X_}37", f'IF({S_}=0,"",IF({al},"алюм.: ПУЭ 1.3.7",IFERROR(INDEX(CHOOSE(IF({cores}=1,1,2),Iдоп_1ж,Iдоп_3ж),MATCH({S_},Iдоп_S,0))*IF({cores}>=4,0.92,1)*{n}*К_прокл,"нет в табл.")))', "N2")
        F(f"{X_}38", f'IF(N({X_}24)=0,"",IF(AND(ISNUMBER({X_}30),{X_}30>{X_}24),"Iр > In — увеличить аппарат",IF(AND(ISNUMBER({X_}37),{X_}24>{X_}37),"In > Iдоп — кабель не защищён","ok")))', "TXT")
        kk = f'IF({X_}25="B",5,IF({X_}25="D",20,10))'
        F(f"{X_}39", f'IF(OR({X_}59=0,NOT(ISNUMBER({X_}35))),"",IF({X_}35>={kk}*{X_}59,"ok","Iкз < "&{kk}&"·In — увеличить сечение или уменьшить In"))', "TXT")
        F(f"{X_}40", (f'IF(OR({S_}=0,{X_}58=""),"",IF((IF({al},76,115)*{S_}*{n}/(1.05*{UF}/{X_}58))^2<0.1,'
                      f'"t доп = "&TEXT((IF({al},76,115)*{S_}*{n}/(1.05*{UF}/{X_}58))^2,"0.000")&" с < 0,1 с — проверить по I²t аппарата","ok"))'), "TXT")
        mx = f'_xlfn.MAXIFS({OL}$AU$3:$AU$1000,{OL}$A$3:$A$1000,{X_}$28)'
        F(f"{X_}41", f'IF(OR({X_}59=0,{mx}=0),"",IF({X_}59>=1.6*{mx},"ok по номиналам","In "&{X_}59&" < 1,6·"&{mx}&" А — проверить картами селективности"))', "TXT")
    # пустой щит (нет во «Диапазонах») → пусто, а не 0 / 0,95
    for X_ in "CDE":
        for r in range(29, 42):
            ref = f"{X_}{r}"
            cells[ref] = re.sub(r"<f>(.*?)</f>", lambda m: f'<f>IF({X_}$28="","",{m.group(1)})</f>', cells[ref], count=1, flags=re.S)
    # 4. экспорт
    T("B43", "4. ДЛЯ ВЫШЕСТОЯЩЕГО ЩИТА (Копировать C:I → в его «Нижестоящие_щиты» столбцы D:J → Вставить связь)", "BT"); merges.append("B43:I43")
    for col in "CDEFGHI": E(f"{col}43", "BT")
    for col, t in zip("BCDEFGHI", ["Щит", "Руст, кВт", "Руст быт, кВт", "Руст сил, кВт", "Рр сил (Σ P·Кс), кВт", "Руст пост, кВт", "cos φ", "Iр, А"]):
        T(f"{col}44", t, "H")
    for k in range(3):
        r = 45 + k; rb = BLK[k]
        F(f"B{r}", f"$B${6 + k}", "TXT")
        F(f"C{r}", f'IFERROR({NG}$AF${rb + 1}*1,"")', "BN2")
        F(f"D{r}", f'IFERROR({NG}$AF${rb + 2}*1,"")', "BN2")
        F(f"E{r}", f'IFERROR(C{r}-D{r}-G{r},"")', "BN2")
        F(f"F{r}", f'IFERROR({NG}$AF${rb + 4}*1,"")', "BN2")
        F(f"G{r}", f'IFERROR({NG}$AF${rb + 5}*1,"")', "BN2")
        F(f"H{r}", f'IFERROR({NG}$V${rb + 4}*1,"")', "BN2")
        F(f"I{r}", f'IFERROR({NG}$V${rb + 3}*1,"")', "BN2")
    # 5. настройки
    T("B50", "5. НАСТРОЙКИ РАСЧЁТА КЗ", "BT"); merges.append("B50:G50")
    for col in "CDEFG": E(f"{col}50", "BT")
    SET = [(51, "ρ меди для Iкз min (нагретые жилы)", 0.0263, "Ом·мм²/м", "≈ 1,5·ρ20 (нагрев при КЗ, ГОСТ 28249)"),
           (52, "ρ алюминия для Iкз min", 0.042, "Ом·мм²/м", "≈ 1,5·ρ20"),
           (53, "Переходные сопротивления (контакты) на каждый щит", 0.015, "Ом", "ГОСТ 28249: учёт переходных сопротивлений"),
           (54, "Коэффициент напряжения для Iкз min", 0.95, "", "c_min (ГОСТ 28249 / МЭК 60909)")]
    for r, lab, v, u, hint in SET:
        T(f"B{r}", lab, "TXT"); N(f"C{r}", v, "IN"); T(f"F{r}", u, "INT"); T(f"G{r}", hint, "TXT")
    widths = {"A": 2, "B": 44, "C": 14, "D": 14, "E": 14, "F": 14, "G": 58, "H": 12, "I": 12}
    dv = dvx([("C17:E17", '"Д/Yн,Y/Yн"'), ("C23:E23", '"медь,алюминий"'), ("C25:E25", '"B,C,D"')])
    new_sheet(WD, order, "Питание_щита", "Однолинейка", cells, merges, widths, dv, "FF70AD47", 3,
              heights={2: 45, 33: 30, 34: 30, 43: 30})
    add_name(WD, "ρкз_Cu", "Питание_щита!$C$51"); add_name(WD, "ρкз_Al", "Питание_щита!$C$52")
    add_name(WD, "R_пер", "Питание_щита!$C$53"); add_name(WD, "c_min", "Питание_щита!$C$54")
    add_name(WD, "ПЩ_щиты", "Питание_щита!$C$28:$E$28")
    add_name(WD, "ПЩ_dU", "Питание_щита!$C$33:$E$33")
    add_name(WD, "ПЩ_Z", "Питание_щита!$C$34:$E$34")
    add_name(WD, "ПЩ_Iкз3", "Питание_щита!$C$36:$E$36")


# =====================================================================================  «Нижестоящие_щиты»
def sheet_children(WD, order, S, prefill=()):
    cells = {}; merges = []
    def T(ref, v, s): cells[ref] = ct(ref, v, S[s])
    def F(ref, f, s): cells[ref] = cf(ref, f, S[s])
    def E(ref, s): cells[ref] = ce(ref, S[s])
    T("A1", "НИЖЕСТОЯЩИЕ ЩИТЫ — подключение других однолинеек", "T")
    T("A2", "1) B — линия этого щита, которая питает нижестоящий щит (номер как в «Исходных данных»). 2) Откройте файл нижестоящего щита → «Питание_щита», раздел 4 → "
            "выделите C:I строки его щита → Копировать → здесь выделите D:J этой строки → Специальная вставка → «Вставить связь». "
            "3) В файле нижестоящего вставьте связью K:N этой строки в его «Питание_щита», раздел 1. Пока связи нет — линия считается как обычная.", "NOTE")
    merges.append("A2:O2")
    T("D3", "ИЗ ФАЙЛА НИЖЕСТОЯЩЕГО (Вставить связь)", "H"); merges.append("D3:J3")
    for col in "EFGHIJ": E(f"{col}3", "H")
    T("K3", "ДЛЯ НИЖЕСТОЯЩЕГО (копировать)", "H"); merges.append("K3:N3")
    for col in "LMN": E(f"{col}3", "H")
    for col in "ABCO": E(f"{col}3", "H")
    HEAD = ["№", "Линия (фидер)", "Файл нижестоящего (для себя)", "Руст, кВт", "Руст быт, кВт", "Руст сил, кВт", "Рр сил, кВт",
            "Руст пост, кВт", "cos φ", "Iр, А", "ΔU от ВРУ до его шин, %", "Zпетли на его шинах, Ом", "Iкз max на его шинах, кА",
            "In аппарата фидера, А", "Проверка фидера"]
    for i, t in enumerate(HEAD):
        T(f"{chr(65 + i)}4", t, "H")
    for i in range(NCH):
        r = 5 + i
        mt = f"MATCH($B{r},{ISX}$D$1:$D$1000,0)"
        F(f"A{r}", f'IF($B{r}="","",{i + 1})', "INT")
        if i < len(prefill): T(f"B{r}", prefill[i], "IN")
        else: E(f"B{r}", "IN")
        E(f"C{r}", "IN")
        for col in "DEFGHIJ": E(f"{col}{r}", "IN")
        F(f"K{r}", f'IF($B{r}="","",IFERROR(INDEX({ISX}$BG$1:$BG$1000,{mt})*1,""))', "N2")
        F(f"L{r}", f'IF($B{r}="","",IFERROR(INDEX({ISX}$BL$1:$BL$1000,{mt})*1,""))', "N2")
        F(f"M{r}", f'IF($B{r}="","",IFERROR(INDEX({ISX}$BM$1:$BM$1000,{mt})*1,""))', "N2")
        F(f"N{r}", f'IF($B{r}="","",IFERROR(--INDEX({OL}$AU$1:$AU$1000,MATCH($B{r},{OL}$R$1:$R$1000,0)),IFERROR(INDEX({ISX}$U$1:$U$1000,{mt})*1,"")))', "N2")
        pn = f"INDEX({RZ}$F$1:$F$1000,MATCH($B{r},{RZ}$H$1:$H$1000,0))"
        i3 = f"INDEX(ПЩ_Iкз3,MATCH({pn},ПЩ_щиты,0))"
        sv = f"INDEX({ISX}$V$1:$V$1000,{mt})"
        th = f'IFERROR(IF((115*{sv}/({i3}*1000))^2<0.1,"термостойкость кабеля: проверить I²t аппарата",""),"")'
        F(f"O{r}", (f'IF($B{r}="","",IF(ISERROR({mt}),"нет такой линии в «Исходных данных»",IF(NOT(ISNUMBER(D{r})),"нет связи с файлом",'
                    f'IF(_xlfn.TEXTJOIN("; ",TRUE,IF(AND(ISNUMBER(J{r}),ISNUMBER(N{r}),J{r}>N{r}),"Iр "&TEXT(J{r},"0")&" > In "&N{r},""),{th})="","ok",'
                    f'_xlfn.TEXTJOIN("; ",TRUE,IF(AND(ISNUMBER(J{r}),ISNUMBER(N{r}),J{r}>N{r}),"Iр "&TEXT(J{r},"0")&" > In "&N{r},""),{th})))))'), "TXT")
    widths = {"A": 4, "B": 14, "C": 26, "D": 9, "E": 9, "F": 9, "G": 9, "H": 9, "I": 7, "J": 9, "K": 11, "L": 11, "M": 11, "N": 10, "O": 40}
    new_sheet(WD, order, "Нижестоящие_щиты", "Питание_щита", cells, merges, widths, "", "FF70AD47", 4,
              heights={2: 45, 4: 45})
    for nm, col in (("НЩ_линии", "B"), ("НЩ_Руст", "D"), ("НЩ_быт", "E"), ("НЩ_сил", "F"), ("НЩ_Ррсил", "G"),
                    ("НЩ_пост", "H"), ("НЩ_cos", "I"), ("НЩ_Iр", "J")):
        add_name(WD, nm, f"Нижестоящие_щиты!${col}$5:${col}${4 + NCH}")


def LK(nm, d):
    """значение из таблицы нижестоящих по номеру линии (пусто, если связи нет)"""
    m = f"MATCH({d},НЩ_линии,0)"
    return f'IFERROR(IF(INDEX({nm},{m})="","",INDEX({nm},{m})*1),"")'


# =====================================================================================  правки существующих листов
def fix_isx(WD):
    X.load_sst(WD)
    isx = X.Sheet(X.sheet_path(WD, "Исходные данные"))
    S_H = isx.style("R2"); S_IN = isx.style("R3"); S_F = isx.style("AT3"); S_1 = isx.style("AT1")
    print("  общих формул развёрнуто (Исходные AL):", unshare(isx, {"AL"}))
    c = {}
    for i, (col, t) in enumerate([("BH", "К прокл. линии (пусто = общий)"), ("BI", "Iкз min в конце линии, А"), ("BJ", "Iкз / In"),
                                  ("BK", "Отключение ≤ 0,4 с (ПУЭ 1.7.79)"), ("BL", "Zпетли в конце, Ом"), ("BM", "Iкз max в конце, кА")]):
        c[f"{col}1"] = cn(f"{col}1", 60 + i, S_1); c[f"{col}2"] = ct(f"{col}2", t, S_H)
    def ftext(ref):
        m = isx.get(ref)
        if not m: return None
        mm = re.search(r"<f[^>]*>([^<]*)</f>", m.group(0))
        return html.unescape(mm.group(1)) if mm else None
    n = [0] * 6
    for r in range(3, 1001):
        d = f"D{r}"
        lk = lambda nm: LK(nm, d)
        # группа «Щит», если есть связь
        f = ftext(f"BC{r}")
        if f and "НЩ_Руст" not in f:
            c[f"BC{r}"] = cf(f"BC{r}", f'IF(ISNUMBER({lk("НЩ_Руст")}),"Щит",{f})', isx.style(f"BC{r}")); n[0] += 1
        for col, nm in (("T", "НЩ_Iр"), ("P", "НЩ_cos")):
            f = ftext(f"{col}{r}")
            if f:
                c[f"{col}{r}"] = cf(f"{col}{r}", f'IF(ISNUMBER({lk(nm)}),{lk(nm)},{f})', isx.style(f"{col}{r}")); n[1] += 1
        f = ftext(f"AL{r}")
        if f == f"E{r}":
            c[f"AL{r}"] = cf(f"AL{r}", f'IF(ISNUMBER({lk("НЩ_Руст")}),{lk("НЩ_Руст")},E{r})', isx.style(f"AL{r}")); n[2] += 1
        # К прокладки на линию
        f = ftext(f"V{r}")
        if f and "*К_прокл" in f:
            c[f"V{r}"] = cf(f"V{r}", f.replace("*К_прокл", f"*IF(ISNUMBER(BH{r}),BH{r},К_прокл)"), isx.style(f"V{r}")); n[3] += 1
        # ΔU от ВРУ: + ΔU до шин щита линии
        pnl = f"INDEX({RZ}$F$1:$F$1000,MATCH(D{r},{RZ}$H$1:$H$1000,0))"
        pc = f"MATCH({pnl},ПЩ_щиты,0)"
        c[f"BG{r}"] = cf(f"BG{r}", f'IF(ISNUMBER(AB{r}),AB{r}+IFERROR(INDEX(ПЩ_dU,{pc})*1,0),"")', isx.style(f"BG{r}")); n[4] += 1
        if isx.is_empty(f"BH{r}"): c[f"BH{r}"] = ce(f"BH{r}", S_IN)
        c[f"BL{r}"] = cf(f"BL{r}", f'IFERROR(IF(OR(D{r}="",NOT(ISNUMBER(V{r})),N(M{r})=0),"",INDEX(ПЩ_Z,{pc})+2*ρкз_Cu*M{r}/V{r}),"")', S_F)
        c[f"BI{r}"] = cf(f"BI{r}", f'IF(ISNUMBER(BL{r}),c_min*{UF}/BL{r},"")', S_F)
        c[f"BM{r}"] = cf(f"BM{r}", f'IFERROR(IF(OR(D{r}="",NOT(ISNUMBER(V{r}))),"",1.05*{UF}/(1.05*{UF}/(INDEX(ПЩ_Iкз3,{pc})*1000)+0.0175*N(M{r})/V{r})/1000),"")', S_F)
        rr = f"MATCH(D{r},{OL}$R$1:$R$1000,0)"
        g1 = f"MATCH(INDEX({OL}$AF$1:$AF$1000,{rr}),{OL}$AF$1:$AF$1000,0)"   # первая строка группы автомата
        inl = f"IFERROR(--INDEX({OL}$AU$1:$AU$1000,{g1}),N(U{r}))"
        c[f"BJ{r}"] = cf(f"BJ{r}", f'IFERROR(IF(ISNUMBER(BI{r}),BI{r}/{inl},""),"")', S_F)
        h = f'" "&IFERROR(INDEX({OL}$H$1:$H$1000,{g1}),"")'
        kk = f'IF(OR(ISNUMBER(SEARCH(" B",{h})),ISNUMBER(SEARCH(" В",{h}))),5,IF(ISNUMBER(SEARCH(" D",{h})),20,10))'
        rcd = f"IFERROR(INDEX({OL}$AS$1:$AS$1000,{g1}),0)"
        c[f"BK{r}"] = cf(f"BK{r}", f'IF(NOT(ISNUMBER(BJ{r})),"",IF(BJ{r}>={kk},"ok",IF({rcd}=1,"ok (УЗО 30 мА)","мало: Iкз < "&{kk}&"·In")))', S_F); n[5] += 1
    bulk(isx, c)
    D_RED = add_dxf(WD, '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>')
    add_cf(isx, "BK3:BK1000", 'LEFT(BK3,4)="мало"', D_RED)
    set_cols(isx, {"BH": 12, "BI": 11, "BJ": 8, "BK": 16, "BL": 10, "BM": 10})
    isx.save()
    print("  Исходные: BC/T,P/AL/V/BG/BK:", n)


def block_formulas(c, r, Bv):
    """блок AE:AH «РАСЧЁТ Рр» щита в строке r — с учётом нижестоящих щитов (группа «Щит»)"""
    Cc, A, H, N, AJ, AK = ("$C$3:$C$1000", "$A$3:$A$1000", "$H$3:$H$1000", "$N$3:$N$1000", "$AJ$3:$AJ$1000", "$AK$3:$AK$1000")
    AN, AO, AP, AQ = "$AN$3:$AN$1000", "$AO$3:$AO$1000", "$AP$3:$AP$1000", "$AQ$3:$AQ$1000"
    nm = f"$R${r}"; cos = f"$V${r + 4}"
    for col, sec in (("AF", ""), ("AG", f",{AJ},1"), ("AH", f",{AJ},2")):
        sif = lambda rng, g: f'SUMIFS({rng},{Cc},{nm},{A},"{g}"{sec})'
        R = lambda k: f"{col}{r + k}"
        g_ = lambda f: f'IF(ISERROR({nm}),"",{f})'
        c[R(1)] = cf(R(1), g_(f'{sif(H, "Быт")}+{sif(H, "Сил")}+{sif(H, "Пост")}+{sif(H, "Щит")}'), Bv)
        c[R(2)] = cf(R(2), g_(f'{sif(H, "Быт")}+{sif(AN, "Щит")}'), Bv)
        c[R(3)] = cf(R(3), g_(P7.ks_interp(R(2))), Bv)
        c[R(4)] = cf(R(4), g_(f'{sif(AK, "Сил")}+{sif(AP, "Щит")}'), Bv)
        c[R(5)] = cf(R(5), g_(f'{sif(H, "Пост")}+{sif(AQ, "Щит")}'), Bv)
        c[R(6)] = cf(R(6), g_(f"{R(2)}*{R(3)}+Ко_сил*{R(4)}+{R(5)}"), Bv)
        c[R(7)] = cf(R(7), g_(f"IF({R(1)}=0,0,{R(6)}/{R(1)})"), Bv)
        ph = lambda k: (f'(SUMIFS({H},{N},{nm}&"L1,L2,L3",{A},"<>Нет"{sec})/3+SUMIFS({H},{N},{nm}&"L{k}",{A},"<>Нет"{sec}))')
        c[R(8)] = cf(R(8), g_(f"IFERROR(MAX(MAX({ph(1)},{ph(2)},{ph(3)})*{R(7)}/({UF}/1000*{cos}),{R(6)}/(SQRT(3)*Uн_лин*{cos})),0)"), Bv)


def fix_load(WD):
    ld = X.Sheet(X.sheet_path(WD, "Нагрузка_щитов"))
    print("  общих формул развёрнуто (Нагрузка H):", unshare(ld, {"H"}))
    Bv = ld.style("AF5"); L_H = ld.style("AJ2"); L_V = ld.style("AK3")
    c = {}
    for r in range(3, 1001):
        lk = lambda nm: LK(nm, f"D{r}")
        c[f"H{r}"] = cf(f"H{r}", f'IF(ISNUMBER({lk("НЩ_Руст")}),{lk("НЩ_Руст")},VLOOKUP(D{r},Исходные_данные,5,FALSE))', ld.style(f"H{r}"))
        m = ld.get(f"B{r}")
        f = html.unescape(re.search(r"<f[^>]*>([^<]*)</f>", m.group(0)).group(1))
        if not f.startswith('IF(A'):
            raise RuntimeError(("B", r, f[:40]))
        if 'A{0}="Щит"'.format(r) not in f:
            c[f"B{r}"] = cf(f"B{r}", f'IF(A{r}="Щит","щит",{f})', ld.style(f"B{r}"))
        for col, nm in (("AN", "НЩ_быт"), ("AO", "НЩ_сил"), ("AP", "НЩ_Ррсил"), ("AQ", "НЩ_пост")):
            c[f"{col}{r}"] = cf(f"{col}{r}", f'IF(A{r}="Щит",N({lk(nm)}),0)', L_V)
        mt = f"MATCH(D{r},{ISX}$D$1:$D$1000,0)"
        c[f"AM{r}"] = cf(f"AM{r}", f'IFERROR(INDEX({ISX}$BG$1:$BG$1000,{mt})*1,"")', L_V)
    for col, t in (("AM", "ΔU от ВРУ, %"), ("AN", "Щит: Руст быт"), ("AO", "Щит: Руст сил"), ("AP", "Щит: Рр сил"), ("AQ", "Щит: Руст пост")):
        c[f"{col}2"] = ct(f"{col}2", t, L_H)
    for r in BLK: block_formulas(c, r, Bv)
    bulk(ld, c)
    set_cols(ld, {"AM": 10, "AN": 10, "AO": 10, "AP": 10, "AQ": 10})
    ld.save()


def fix_sl(WD):
    sh = X.Sheet(X.sheet_path(WD, "Однолинейка"))
    c = {}; n = [0, 0]
    for r in range(3, 1001):
        m = sh.get(f"AB{r}")
        if m and "*К_прокл" in html.unescape(m.group(0)):
            f = html.unescape(re.search(r"<f[^>]*>([^<]*)</f>", m.group(0)).group(1))
            bh = f"INDEX({ISX}$BH$1:$BH$1000,MATCH(R{r},{ISX}$D$1:$D$1000,0))"
            c[f"AB{r}"] = cf(f"AB{r}", f.replace("*К_прокл", f"*IFERROR(IF(ISNUMBER({bh}),{bh},К_прокл),К_прокл)"), sh.style(f"AB{r}")); n[0] += 1
        m = sh.get(f"AE{r}")
        if m and "<f" in m.group(0):
            f = html.unescape(re.search(r"<f[^>]*>([^<]*)</f>", m.group(0)).group(1))
            if "Iкз мало" not in f and f.endswith(")))"):
                term = f'IFERROR(IF(LEFT(INDEX({ISX}$BK$1:$BK$1000,MATCH(R{r},{ISX}$D$1:$D$1000,0)),4)="мало","Iкз мало — отключение > 0,4 с",""),"")'
                # добавить аргумент в TEXTJOIN (последние две скобки — TEXTJOIN и внешний IF)
                c[f"AE{r}"] = cf(f"AE{r}", f[:-2] + "," + term + "))", sh.style(f"AE{r}")); n[1] += 1
    bulk(sh, c); sh.save()
    print("  Однолинейка AB/AE:", n)


def fix_pz(WD, S):
    """«Расчёт_нагрузок»: нижестоящие щиты в категориях + блок основных показателей"""
    pz = X.Sheet(X.sheet_path(WD, "Расчёт_нагрузок"))
    c = {}
    crit = f',{NG}$AJ$3:$AJ$1000,">="&$M$4,{NG}$AJ$3:$AJ$1000,"<="&$M$5'
    def sif(rng, grp, r):
        return f'SUMIFS({NG}${rng}$3:${rng}$1000,{NG}$C$3:$C$1000,$C$4,{NG}$AL$3:$AL$1000,$L{r},{NG}$A$3:$A$1000,"{grp}"{crit})'
    for r in range(11, 41):
        st = pz.style(f"M{r}")
        c[f"M{r}"] = cf(f"M{r}", f'IF($L{r}="",0,{sif("H", "Быт", r)}+{sif("AN", "Щит", r)})', st)
        c[f"N{r}"] = cf(f"N{r}", f'IF($L{r}="",0,{sif("H", "Сил", r)}+{sif("AO", "Щит", r)})', st)
        c[f"O{r}"] = cf(f"O{r}", f'IF($L{r}="",0,{sif("AK", "Сил", r)}+{sif("AP", "Щит", r)})', st)
        c[f"P{r}"] = cf(f"P{r}", f'IF($L{r}="",0,{sif("H", "Пост", r)}+{sif("AQ", "Щит", r)})', st)
        c[f"R{r}"] = cf(f"R{r}", f'IF($L{r}="",0,{sif("L", "Быт", r)}+{sif("L", "Сил", r)}+{sif("L", "Пост", r)}+{sif("L", "Щит", r)})', st)
    # основные показатели
    def T(ref, v, s): c[ref] = ct(ref, v, S[s])
    def F(ref, f, s): c[ref] = cf(ref, f, S[s])
    def E(ref, s): c[ref] = ce(ref, S[s])
    merges = []
    T("A52", "ОСНОВНЫЕ ПОКАЗАТЕЛИ (для ПЗ)", "BT"); merges.append("A52:J52")
    for col in "BCDEFGHIJ": E(f"{col}52", "BT")
    pc = "MATCH($C$4,ПЩ_щиты,0)"
    col_ps = lambda row: f'INDEX({PS}$C${row}:$E${row},{pc})'
    ROWS = [
        ("Напряжение сети", f'"~"&TEXT(Uн_лин*1000,"0")&"/"&TEXT({UF},"0")&" В, 50 Гц"', None),
        ("Система заземления", None, "TN-C-S"),
        ("Категория надёжности электроснабжения", "$J$4", None),
        ("Установленная мощность Руст, кВт", 'TEXT(C42,"0.00")', None),
        ("Расчётная мощность Рр, кВт / Sр, кВА / cos φ", 'TEXT(G42,"0.00")&" / "&TEXT(I42,"0.00")&" / "&TEXT(E42,"0.00")', None),
        ("Расчётный ток Iр, А", 'TEXT(J42,"0.0")', None),
        ("Питание щита", f'IFERROR(IF({col_ps(12)}="","(заполните лист «Питание_щита»)","от "&{col_ps(12)}&IF({col_ps(18)}="","",", кабель "&{col_ps(18)})&IF(N({col_ps(19)})=0,""," "&IF(N({col_ps(20)})=0,5,{col_ps(20)})&"×"&{col_ps(19)})&IF(N({col_ps(22)})=0,"",", L = "&{col_ps(22)}&" м")),"")', None),
        ("Аппарат защиты питающей линии, А", f'IFERROR(IF({col_ps(59)}=0,"",{col_ps(59)}&" А"),"")', None),
        ("ΔU от ВРУ до шин щита / max в линиях, % (СП 256 п. 8.23: свет ≤ 3, прочие ≤ 4)", f'IFERROR(TEXT({col_ps(33)},"0.00"),"—")&" / "&IFERROR(TEXT(_xlfn.MAXIFS({NG}$AM$3:$AM$1000,{NG}$C$3:$C$1000,$C$4),"0.00"),"—")', None),
        ("Iкз min (1ф) / Iкз max (3ф) на шинах щита", f'IF(ISNUMBER({col_ps(35)}),TEXT({col_ps(35)},"0")&" А","—")&" / "&IF(ISNUMBER({col_ps(36)}),TEXT({col_ps(36)},"0.00")&" кА","—")', None),
        ("Компенсация реактивной мощности", "C43", None),
        ("Учёт электроэнергии", None, "(тип счётчика, класс точности, прямого/трансформаторного включения)"),
    ]
    for i, (lab, f, inp) in enumerate(ROWS):
        r = 53 + i
        T(f"A{r}", "", "TXT"); T(f"B{r}", lab, "TXT")
        if f: F(f"C{r}", f, "TXT")
        else: T(f"C{r}", inp, "IN")
        merges.append(f"C{r}:J{r}")
        for col in "DEFGHIJ": E(f"{col}{r}", "IN" if inp else "TXT")
    bulk(pz, c)
    # сдвинуть примечания ниже? они на 46–50 — оставляем; добавить слияния и область печати
    mm = re.search(r'<mergeCells count="(\d+)">', pz.s)
    n0 = int(mm.group(1))
    k = pz.s.find("</mergeCells>")
    pz.s = pz.s[:k] + "".join(f'<mergeCell ref="{m}"/>' for m in merges) + pz.s[k:]
    pz.s = pz.s.replace(mm.group(0), f'<mergeCells count="{n0 + len(merges)}">', 1)
    pz.save()
    pzs = X.Sheet(X.sheet_path(WD, "Расчёт_нагрузок")); add_dv(pzs, [("C54", '"TN-C-S,TN-S,TT,IT"')]); pzs.save()
    wbp = WD + "/xl/workbook.xml"; wb = rd(wbp)
    wb = wb.replace("Расчёт_нагрузок!$A$1:$J$50", "Расчёт_нагрузок!$A$1:$J$64")
    wr(wbp, wb)


def fix_ref(WD):
    X.load_sst(WD)
    spr = X.Sheet(X.sheet_path(WD, "Справочная"))
    c = {"BP31": ce("BP31", spr.style("BP31")), "BQ31": ce("BQ31", spr.style("BQ31")),
         "BR31": ct("BR31", "ΔU до шин щита теперь считается на листе «Питание_щита».", spr.style("BR31"))}
    bulk(spr, c); spr.save()


def guide(WD):
    gd = X.Sheet(X.sheet_path(WD, "Краткое руководство"))
    sE = gd.style("E13"); sF = gd.style("F13"); sB = gd.style("B18"); sC = gd.style("C18")
    bulk(gd, {
        "E16": ct("E16", "v3.11 (связи щитов)", sE),
        "F16": ct("F16", "Листы «Питание_щита» (ΔU и Iкз до шин, проверка питающей линии) и «Нижестоящие_щиты» (подключение других однолинеек через «Вставить связь»); "
                         "Iкз в конце каждой линии и отключение ≤ 0,4 с; К прокладки на линию («Исходные данные» BH); основные показатели для ПЗ.", sF),
        "B21": ct("B21", "20", sB),
        "C21": ct("C21", "СВЯЗИ ЩИТОВ: вышестоящий файл, «Нижестоящие_щиты» — строка на линию-фидер; D:J — Вставить связь из «Питание_щита» (раздел 4) нижестоящего файла; "
                         "K:N — Вставить связью в «Питание_щита» (раздел 1) нижестоящего. Файлы не переименовывать (иначе Данные → Изменить связи).", sC)})
    gd.save()


PREFILL_GRSH = ["ЩР-001", "ЩР-01", "ЩР-1", "ЩР-2", "ЩО-2"]


def build(src, out, grsh=False):
    WD = os.path.join(HERE, "w11_" + os.path.basename(out)[:6].replace(" ", "_"))
    order = P7.unpack(src, WD)
    print("==", os.path.basename(out))
    S = P8.styles_add(WD)
    sheet_supply(WD, order, S)
    sheet_children(WD, order, S, PREFILL_GRSH if grsh else ())
    fix_isx(WD); fix_load(WD); fix_sl(WD); fix_pz(WD, S); fix_ref(WD); guide(WD)
    P7.pack(WD, order, out)


if __name__ == "__main__":
    SRC = os.path.join(HERE, "out10")
    build(os.path.join(SRC, "Однолинейка_пример_ЖК-Остров.xlsx"), os.path.join(OUT, "Однолинейка_пример_ЖК-Остров.xlsx"))
    build(os.path.join(SRC, "Однолинейка_шаблон.xlsx"), os.path.join(OUT, "Однолинейка_шаблон.xlsx"))
    build(os.path.join(SRC, "Однолинейка ГРЩ-ЖД.xlsx"), os.path.join(OUT, "Однолинейка ГРЩ-ЖД.xlsx"), grsh=True)
    print("готово")
