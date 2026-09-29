"""v3.10: проверка по нормам и лист «Расчёт_нагрузок» (таблица для ПЗ стадии П).

Исправления (по результатам проверки):
  1. Ток линии: I = P/(U·cos φ), Uн = 230/400 В (ГОСТ 29322-2014) — раньше без cos φ и при 220/380.
  2. Сечение: допустимый ток по ПУЭ табл. 1.3.6 (одножильные / двухжильные для 1ф / трёхжильные для 3ф,
     4–5-жильные ×0,92 по ПУЭ 1.3.10) и общий коэффициент прокладки. Раньше — одна таблица без учёта фаз и жил,
     для ≥10 мм² на 5–10 % выше ПУЭ. «Макс. автомат по кабелю» в «Однолинейке» — по той же таблице.
  3. Потери напряжения: ΔU = b·(ρ1·L/S·cos φ + λ·L·sin φ)·I (ГОСТ Р 50571.5.52 прил. G): длина — ИТОГОВАЯ (M, было I по плану),
     ρ1 = 0,0225 Ом·мм²/м (Cu при рабочей t°), % от 230 В. Норма СП 256 п. 8.23: от ВРУ свет ≤ 3 %, прочие ≤ 4 %
     (с учётом потерь до шин щита — настройка).
  4. Автоматы: таблица номиналов продлена до 630 А (было до 63 А, дальше «Не исп»).
  5. ТКЗ: O «Общее» = участок + предыдущий (было ρ + участок), Q = Iкз/In (было Z/In),
     сопротивление трансформатора Zт = uк·U²/S (было в 10 раз меньше).
Новое: лист «Расчёт_нагрузок» — категории (как в ПЗ), Руст, Кс, cos φ, tg φ, Рр, Qр, Sр, Iр; итог с Ко; КРМ; сверка.

Запуск: python3 patch8.py   (берёт out/*.xlsx v3.9 из patch7, пишет out10/)
"""
import os, re, sys, shutil, html
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import patch7 as P7
from patch7 import X, rd, wr, ct, cn, cf, ce, bulk, set_cols, add_dxf, add_cf, add_dv, add_name, unshare, esc

OUT = os.path.join(HERE, "out10"); os.makedirs(OUT, exist_ok=True)
ISX = "'Исходные данные'!"
NG = "Нагрузка_щитов!"

# ---------------------------------------------------------------- справочные данные
# ПУЭ табл. 1.3.6, медь, в воздухе: одножильные / двухжильные / трёхжильные
PUE = [(1.5, 23, 19, 19), (2.5, 30, 27, 25), (4, 41, 38, 35), (6, 50, 50, 42), (10, 80, 70, 55), (16, 100, 90, 75),
       (25, 140, 115, 95), (35, 170, 140, 120), (50, 215, 175, 145), (70, 270, 215, 180), (95, 325, 260, 220),
       (120, 385, 300, 260), (150, 440, 350, 305), (185, 510, 405, 350), (240, 605, None, None)]
BRK = [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630]
CATS = ["Рабочее освещение", "Аварийное освещение", "Наружное и фасадное освещение", "Розеточная сеть и бытовые приборы",
        "Электрический тёплый пол", "Электроотопление (конвекторы, камины)", "Кондиционирование и увлажнение", "Вентиляция",
        "Водонагреватели", "Приводы (шторы, ворота, краны)", "Слаботочные системы, автоматизация, KNX",
        "Системы безопасности и связи", "Системы противопожарной защиты", "Технология пищеприготовления",
        "Бассейн, SPA, хаммам", "Насосы (водоснабжение, дренаж)", "Обогрев водостоков и кровли", "Лифты",
        "Распределительные щиты жилой части", "Прочее"]
CAT_BY_CODE = {"L": 0, "LED": 0, "LD": 0, "DALI": 0, "R": 3, "ТП": 4, "КМН": 5, "K": 6, "УВ": 6, "В": 7, "ПВ": 7, "ПУ": 7,
               "ПН": 8, "SH": 9, "ЭK": 9, "АСУ": 10, "PS-KNX": 10, "TE-O": 10, "TE-ТП": 10, "WD": 10, "VK": 10, "R.ШТК": 10,
               "R.ЩОПС": 11, "R.GSM": 11, "АПС": 11}
CBC = {P7.norm(k): v for k, v in CAT_BY_CODE.items()}
NCAT = 30


def styles_add(WD):
    """добавить шрифты, заливки, границы и форматы ячеек для нового листа; вернуть словарь индексов xf"""
    p = WD + "/xl/styles.xml"; st = rd(p)
    def add(tag, xml):
        nonlocal st
        m = re.search(r'<%s count="(\d+)"' % tag, st); n = int(m.group(1))
        st = st.replace(m.group(0), f'<{tag} count="{n + 1}"', 1)
        k = st.find(f"</{tag}>"); st = st[:k] + xml + st[k:]
        return n
    F_N = add("fonts", '<font><sz val="10"/><name val="Arial"/><family val="2"/><charset val="204"/></font>')
    F_B = add("fonts", '<font><b/><sz val="10"/><name val="Arial"/><family val="2"/><charset val="204"/></font>')
    F_T = add("fonts", '<font><b/><sz val="12"/><name val="Arial"/><family val="2"/><charset val="204"/></font>')
    F_G = add("fonts", '<font><i/><sz val="9"/><color rgb="FF808080"/><name val="Arial"/><family val="2"/><charset val="204"/></font>')
    FL_H = add("fills", '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>')
    FL_IN = add("fills", '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>')
    B = add("borders", '<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right>'
                       '<top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>')
    def xf(font, fill=0, border=0, numfmt=0, h="left", wrap=True, v="center"):
        al = f'<alignment horizontal="{h}" vertical="{v}"' + (' wrapText="1"' if wrap else "") + "/>"
        return add("cellXfs", f'<xf numFmtId="{numfmt}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0" applyNumberFormat="1" '
                              f'applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">{al}</xf>')
    S = dict(
        T=xf(F_T, wrap=False), SUB=xf(F_N, wrap=False), LBL=xf(F_B, h="right", wrap=False),
        IN=xf(F_N, FL_IN, B), H=xf(F_B, FL_H, B, h="center"), HN=xf(F_N, FL_H, B, h="center"),
        TXT=xf(F_N, 0, B), INT=xf(F_N, 0, B, h="center"), N2=xf(F_N, 0, B, 2, h="center"),
        BT=xf(F_B, 0, B), BN2=xf(F_B, 0, B, 2, h="center"), NOTE=xf(F_N, v="top"), G=xf(F_G, wrap=False),
        GN=xf(F_G, 0, 0, 2, h="center", wrap=False), V2=xf(F_N, 0, 0, 2, h="left", wrap=False))
    wr(p, st)
    return S


def ks_interp(x):
    return P7.ks_interp(x)


def fix_reference(WD, grsh):
    X.load_sst(WD)
    spr = X.Sheet(X.sheet_path(WD, "Справочная"))
    R_T = spr.style("BE2"); R_H = spr.style("BE3"); R_L = spr.style("BE4"); R_V = spr.style("BF8"); R_IN = spr.style("BF4")
    c = {}
    # --- ρ для потерь напряжения
    c["D3"] = ct("D3", "ρ при рабочей t° (ГОСТ Р 50571.5.52 прил. G), Ом·мм²/м", spr.style("D3"))
    for r in range(4, 20):
        if X.value(spr, f"B{r}") not in (None, ""):
            c[f"D{r}"] = cn(f"D{r}", 0.0225, spr.style(f"D{r}"))
    # --- номиналы автоматов до 630 А (правило: I ≤ 0,9·In)
    for i, In in enumerate(BRK):
        r = 4 + i
        c[f"F{r}"] = cn(f"F{r}", 0 if i == 0 else round(0.9 * BRK[i - 1], 2), spr.style("F5"))
        c[f"G{r}"] = cn(f"G{r}", In, spr.style("G5"))
    r = 4 + len(BRK)
    c[f"F{r}"] = cn(f"F{r}", round(0.9 * BRK[-1], 2), spr.style("F5")); c[f"G{r}"] = ct(f"G{r}", "Не исп", spr.style("G5"))
    add_name(WD, "Автомат_номиналы", f"Справочная!$F$4:$G${r}")
    c["I2"] = ct("I2", "(устар., не используется с v3.10 — см. BX:CA)", R_L)
    # --- настройки сети
    c["BP28"] = ct("BP28", "НАСТРОЙКИ СЕТИ", R_H); c["BQ28"] = ce("BQ28", R_H)
    SET = [(29, "Uн, кВ", 0.4, "Номинальное линейное напряжение (ГОСТ 29322-2014: 230/400 В). Фазное = Uн/√3. Для 380 В впишите 0,38."),
           (30, "К прокл.", 1, "Общий коэффициент снижения допустимого тока (пучок, многослойно в лотке/коробе — ПУЭ табл. 1.3.12, температура). В воздухе/лотке однорядно — 1."),
           (31, "ΔU пит., %", 0, "Потери напряжения от ВРУ до шин этого щита, % (по питающей линии). Прибавляются к потерям каждой линии и сравниваются с нормой СП 256 п. 8.23.")]
    for r_, lab, v, t in SET:
        c[f"BP{r_}"] = ct(f"BP{r_}", lab, R_L); c[f"BQ{r_}"] = cn(f"BQ{r_}", v, R_IN); c[f"BR{r_}"] = ct(f"BR{r_}", t, R_L)
    add_name(WD, "Uн_лин", "Справочная!$BQ$29")
    add_name(WD, "К_прокл", "Справочная!$BQ$30")
    add_name(WD, "dU_пит", "Справочная!$BQ$31")
    # --- допустимые токи ПУЭ 1.3.6
    c["BX2"] = ct("BX2", "ДОПУСТИМЫЙ ДЛИТЕЛЬНЫЙ ТОК, А (ПУЭ табл. 1.3.6, медь, в воздухе)", R_T)
    for col, t in zip(["BX", "BY", "BZ", "CA"], ["Сечение, мм²", "Одножильные", "Двухжильные (1ф: L+N, PE не нагружен)", "Трёхжильные (3ф); 4–5 жил ×0,92"]):
        c[f"{col}3"] = ct(f"{col}3", t, R_H)
    for i, (s, a1, a2, a3) in enumerate(PUE):
        r_ = 4 + i
        c[f"BX{r_}"] = cn(f"BX{r_}", s, R_IN); c[f"BY{r_}"] = cn(f"BY{r_}", a1, R_IN)
        c[f"BZ{r_}"] = cn(f"BZ{r_}", a2, R_IN) if a2 else ce(f"BZ{r_}", R_IN)
        c[f"CA{r_}"] = cn(f"CA{r_}", a3, R_IN) if a3 else ce(f"CA{r_}", R_IN)
    c["BX20"] = ct("BX20", "Сечение подбирается как наименьшее, у которого Iдоп·k ≥ In автомата (ГОСТ Р 50571.4.43: Ib ≤ In ≤ Iz); k = 0,92 для 4–5-жильных в 3ф (ПУЭ 1.3.10) × «К прокл.». 240 мм² — только одножильные (в табл. 1.3.6 для многожильных нет).", R_L)
    add_name(WD, "Iдоп_S", "Справочная!$BX$4:$BX$18")
    add_name(WD, "Iдоп_1ж", "Справочная!$BY$4:$BY$18")
    add_name(WD, "Iдоп_2ж", "Справочная!$BZ$4:$BZ$17")
    add_name(WD, "Iдоп_3ж", "Справочная!$CA$4:$CA$17")
    # --- категории для ПЗ
    c["BT2"] = ct("BT2", "КАТЕГОРИИ НАГРУЗКИ ДЛЯ ПЗ (лист «Расчёт_нагрузок»)", R_T)
    c["BT3"] = ct("BT3", "Категория (как в ПЗ, можно переименовать/добавить)", R_H)
    c["BU3"] = ct("BU3", "cos φ для ПЗ (пусто = по линиям)", R_H)
    c["BV3"] = ct("BV3", "Примечание", R_H)
    for i in range(NCAT):
        r_ = 4 + i
        c[f"BT{r_}"] = ct(f"BT{r_}", CATS[i], R_IN) if i < len(CATS) else ce(f"BT{r_}", R_IN)
        c[f"BU{r_}"] = ce(f"BU{r_}", R_IN)
    c["BV4"] = ct("BV4", "Категория линии: «Исходные данные» BD (ручная) или по коду — столбец Z этой страницы.", R_L)
    c["BV5"] = ct("BV5", "Кс в ПЗ берётся из расчёта Кс (группы Быт/Сил/Пост), а не задаётся на категорию.", R_L)
    add_name(WD, "Категории_ПЗ", f"Справочная!$BT$4:$BT${3 + NCAT}")
    # --- код → категория
    c["Z3"] = ct("Z3", "Категория (ПЗ)", R_H)
    for r_ in range(4, 52):
        code = X.value(spr, f"P{r_}")
        k = CBC.get(P7.norm(code)) if code else None
        c[f"Z{r_}"] = ct(f"Z{r_}", CATS[k], R_IN) if k is not None else ce(f"Z{r_}", R_IN)
    bulk(spr, c)
    set_cols(spr, {"Z": 22, "BS": 3, "BT": 40, "BU": 12, "BV": 60, "BW": 3, "BX": 10, "BY": 12, "BZ": 16, "CA": 16})
    spr.save()


def fix_isx(WD):
    X.load_sst(WD)
    isx = X.Sheet(X.sheet_path(WD, "Исходные данные"))
    print("  общих формул развёрнуто (Исходные):", unshare(isx, {"T", "V", "AA", "AB", "AC", "AF"}))
    S_H = isx.style("R2"); S_IN = isx.style("R3"); S_F = isx.style("AT3"); S_1 = isx.style("AT1")
    c = {}
    for i, (col, t) in enumerate([("BD", "Категория ПЗ (ввод; пусто = по коду)"), ("BE", "Категория ПЗ (итог)"),
                                  ("BF", "Допустимо ΔU от ВРУ, % (СП 256 п. 8.23)"), ("BG", "ΔU от ВРУ, %")]):
        c[f"{col}1"] = cn(f"{col}1", 56 + i, S_1); c[f"{col}2"] = ct(f"{col}2", t, S_H)
    UF = "(Uн_лин/SQRT(3))"          # фазное, кВ
    n = {"T": 0, "V": 0, "AA": 0, "AB": 0, "AC": 0, "AF": 0}
    k_ = "IF(O{r}=1,1,IF(H{r}=1,1,IF(O{r}>=4,0.92,1)))*К_прокл"
    caps = "CHOOSE(IF(O{r}=1,1,IF(H{r}=1,2,3)),Iдоп_1ж,Iдоп_2ж,Iдоп_3ж)"
    for r in range(3, 1001):
        def has(col):
            m = isx.get(f"{col}{r}"); return m and "<f" in m.group(0)
        def ftext(col):
            return html.unescape(re.search(r"<f[^>]*>([^<]*)</f>", isx.get(f"{col}{r}").group(0)).group(1))
        if has("T"):
            c[f"T{r}"] = cf(f"T{r}", f"_xlfn.IFS(H{r}=1,E{r}/({UF}*P{r}),H{r}=3,E{r}/(SQRT(3)*Uн_лин*P{r}))", isx.style(f"T{r}")); n["T"] += 1
        if has("V"):
            f = ftext("V"); old = f"VLOOKUP(U{r},Кабель_номиналы,2,TRUE)"
            assert old in f, ("V", r, f)
            x = f"(U{r}/({k_.format(r=r)}))"
            cp = caps.format(r=r)
            pos = f"(IFERROR(MATCH({x},{cp},1),0)+IF(ISNUMBER(MATCH({x},{cp},0)),0,1))"
            new = f'IFERROR(IF({pos}>ROWS({cp}),"Не исп",INDEX(Iдоп_S,{pos})),"Не исп")'
            c[f"V{r}"] = cf(f"V{r}", f.replace(old, new), isx.style(f"V{r}")); n["V"] += 1
        if has("AA"):
            c[f"AA{r}"] = cf(f"AA{r}", f"Y{r}*(W{r}*M{r}/V{r}*P{r}+X{r}*M{r}*Z{r}/1000)*T{r}", isx.style(f"AA{r}")); n["AA"] += 1
        if has("AB"):
            c[f"AB{r}"] = cf(f"AB{r}", f"ROUNDUP(AA{r}/({UF}*1000)*100,2)", isx.style(f"AB{r}")); n["AB"] += 1
        if has("AC"):
            c[f"AC{r}"] = cf(f"AC{r}", f"_xlfn.IFS(H{r}=3,Uн_лин*1000,H{r}=1,ROUND({UF}*1000,0))", isx.style(f"AC{r}")); n["AC"] += 1
        if has("AF"):
            c[f"AF{r}"] = cf(f"AF{r}", f"IFERROR(IF(BG{r}>BF{r},1,2),0)", isx.style(f"AF{r}")); n["AF"] += 1
        if isx.is_empty(f"BD{r}"): c[f"BD{r}"] = ce(f"BD{r}", S_IN)
        lk = f"INDEX(Справочная!$Z$4:$Z$51,MATCH(AH{r},Справочная!$P$4:$P$51,0))&\"\""
        c[f"BE{r}"] = cf(f"BE{r}", f'IF(D{r}="","",IF(BD{r}<>"",BD{r},IF(IFERROR({lk},"")="","Прочее",{lk})))', S_F)
        c[f"BF{r}"] = cf(f"BF{r}", f'IF(D{r}="","",IF(IFERROR(VLOOKUP(AH{r},УГО_спр,2,FALSE),"")="Светильник",3,4))', S_F)
        c[f"BG{r}"] = cf(f"BG{r}", f'IF(ISNUMBER(AB{r}),AB{r}+dU_пит,"")', S_F)
    bulk(isx, c)
    add_dv(isx, [("BD3:BD1000", "Категории_ПЗ")])
    D_RED = add_dxf(WD, '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>')
    add_cf(isx, "BG3:BG1000", 'AND(ISNUMBER(BG3),BG3>BF3)', D_RED)
    set_cols(isx, {"BD": 30, "BE": 30, "BF": 12, "BG": 10})
    isx.save()
    print("  переписано:", n)


def fix_sl(WD):
    """«Однолинейка»: макс. автомат по кабелю — по ПУЭ 1.3.6; статус — норма ΔU от ВРУ"""
    sh = X.Sheet(X.sheet_path(WD, "Однолинейка"))
    print("  общих формул развёрнуто (Однолинейка):", unshare(sh, {"AB", "AE"}))
    ser = "{" + ",".join(map(str, BRK)) + "}"
    c = {}; n = [0, 0]
    old_tail = ('IF(IFERROR(_xlfn.NUMBERVALUE(MID(Q{r},FIND("-",Q{r},FIND("L=",Q{r}))+1,FIND("%",Q{r})-FIND("-",Q{r},FIND("L=",Q{r}))-1),","),0)>4,"потери > 4 %","")')
    for r in range(3, 1001):
        m = sh.get(f"AB{r}")
        if m and "<f" in m.group(0) and "LOOKUP(INDEX(Справочная!$I$" in html.unescape(m.group(0)):
            S_ = f"VLOOKUP(R{r},Исходные_данные,22,FALSE)"; ph = f"VLOOKUP(R{r},Исходные_данные,8,FALSE)"; co = f"VLOOKUP(R{r},Исходные_данные,15,FALSE)"
            cap = (f"INDEX(CHOOSE(IF({co}=1,1,IF({ph}=1,2,3)),Iдоп_1ж,Iдоп_2ж,Iдоп_3ж),MATCH({S_},Iдоп_S,0))"
                   f"*IF({co}=1,1,IF({ph}=1,1,IF({co}>=4,0.92,1)))*К_прокл")
            c[f"AB{r}"] = cf(f"AB{r}", f'IF(R{r}="","",IFERROR(LOOKUP({cap},{ser}),""))', sh.style(f"AB{r}")); n[0] += 1
        m = sh.get(f"AE{r}")
        if m and "<f" in m.group(0):
            f = html.unescape(re.search(r"<f[^>]*>([^<]*)</f>", m.group(0)).group(1))
            o = old_tail.format(r=r)
            if o in f:
                mt = f"MATCH(R{r},{ISX}$D$1:$D$1000,0)"
                new = (f'IFERROR(IF(AND(ISNUMBER(INDEX({ISX}$BG$1:$BG$1000,{mt})),INDEX({ISX}$BG$1:$BG$1000,{mt})>INDEX({ISX}$BF$1:$BF$1000,{mt})),'
                       f'"потери "&TEXT(INDEX({ISX}$BG$1:$BG$1000,{mt}),"0.0")&" % > "&INDEX({ISX}$BF$1:$BF$1000,{mt})&" %",""),"")')
                c[f"AE{r}"] = cf(f"AE{r}", f.replace(o, new), sh.style(f"AE{r}")); n[1] += 1
    bulk(sh, c); sh.save()
    print("  Однолинейка AB/AE:", n)


def _ftext(m):
    mm = re.search(r"<f[^>]*>([^<]*)</f>", m.group(0))
    return html.unescape(mm.group(1)) if mm else None


def fix_tkz(WD):
    sh = X.Sheet(X.sheet_path(WD, "ТКЗ"))
    allc = set(re.findall(r'<c r="([A-Z]+)\d+"', sh.s))
    print("  общих формул развёрнуто (ТКЗ, все столбцы):", unshare(sh, allc))
    c = {}; n = [0, 0, 0]
    for r in range(7, 2001):
        m = sh.get(f"O{r}")
        if m and "<f" in m.group(0) and _ftext(m) == f"L{r}+M{r}":
            c[f"O{r}"] = cf(f"O{r}", f"M{r}+IFERROR(N{r},0)", sh.style(f"O{r}")); n[0] += 1
        m = sh.get(f"Q{r}")
        if m and "<f" in m.group(0) and _ftext(m) == f"O{r}/G{r}":
            c[f"Q{r}"] = cf(f"Q{r}", f"P{r}/G{r}", sh.style(f"Q{r}")); n[1] += 1
    m = sh.get("N7")
    if m and "5.5*0.4^2/D$3" in html.unescape(m.group(0)):
        c["N7"] = cf("N7", "5.5/100*0.4^2/(D$3/1000)", sh.style("N7")); n[2] += 1
    c["N5"] = ct("N5", "Zт = uк·U²/S (uк = 5,5 %). Для Y/Yн — см. ГОСТ 28249 (Zт(1)/3).", sh.style("B5"))
    bulk(sh, c)
    D_RED = add_dxf(WD, '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>')
    add_cf(sh, "N8:N2000", 'AND($C8<>"",ISERROR(N8))', D_RED)
    sh.save()
    print("  ТКЗ O/Q/N7:", n)


def fix_load(WD):
    ld = X.Sheet(X.sheet_path(WD, "Нагрузка_щитов"))
    c = {}
    for r in range(3, 1001):
        mt = f"MATCH(D{r},{ISX}$D$1:$D$1000,0)"
        c[f"AL{r}"] = cf(f"AL{r}", f'IF(A{r}="","",IFERROR(INDEX({ISX}$BE$1:$BE$1000,{mt})&"","Прочее"))', ld.style("AJ3"))
    c["AL2"] = ct("AL2", "Категория (ПЗ)", ld.style("AJ2"))
    bulk(ld, c)
    # Iр в блоках AE:AH — через Uн
    def fx(mm):
        cell = mm.group(0)
        return cell.replace("(0.22*", "(Uн_лин/SQRT(3)*").replace("(1.732*0.38*", "(SQRT(3)*Uн_лин*")
    ld.s = re.sub(r'<c r="A[FGH]\d+"[^>]*>.*?</c>', fx, ld.s, flags=re.S)
    set_cols(ld, {"AL": 30})
    ld.save()


def add_pz_sheet(WD, order, S):
    """новый лист «Расчёт_нагрузок» после «Нагрузка_щитов»"""
    wbp = WD + "/xl/workbook.xml"; wb = rd(wbp)
    rp = WD + "/xl/_rels/workbook.xml.rels"; rels = rd(rp)
    ctp = WD + "/[Content_Types].xml"
    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]; rid = f"rId{max(ids) + 1}"
    nums = [int(x) for x in re.findall(r"worksheets/sheet(\d+)\.xml", rels)]; sn = max(nums) + 1
    part = f"xl/worksheets/sheet{sn}.xml"
    sid = max(int(x) for x in re.findall(r'sheetId="(\d+)"', wb)) + 1
    names = re.findall(r'<sheet name="([^"]+)"', wb)
    pos = names.index("Нагрузка_щитов") + 1
    tags = re.findall(r"<sheet [^>]*/>", wb)
    tags.insert(pos, f'<sheet name="Расчёт_нагрузок" sheetId="{sid}" r:id="{rid}"/>')
    wb = re.sub(r"<sheets>.*?</sheets>", "<sheets>" + "".join(tags) + "</sheets>", wb, flags=re.S)
    wb = re.sub(r'localSheetId="(\d+)"', lambda m: f'localSheetId="{int(m.group(1)) + (1 if int(m.group(1)) >= pos else 0)}"', wb)
    wb = re.sub(r'activeTab="(\d+)"', lambda m: f'activeTab="{int(m.group(1)) + (1 if int(m.group(1)) >= pos else 0)}"', wb)
    wr(wbp, wb)
    wr(rp, rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{sn}.xml"/></Relationships>'))
    wr(ctp, rd(ctp).replace("</Types>", f'<Override PartName="/{part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'))
    order.append(part)

    cells = {}
    def T(ref, v, s): cells[ref] = ct(ref, v, S[s])
    def N(ref, v, s): cells[ref] = cn(ref, v, S[s])
    def F(ref, f, s): cells[ref] = cf(ref, f, S[s])
    def E(ref, s): cells[ref] = ce(ref, S[s])
    merges = []
    F("A1", '"Расчёт электрических нагрузок объекта «"&$C$3&"»"', "T"); merges.append("A1:J1")
    F("A2", '"Щит (ВУ): "&$C$4&"   ·   режим: "&$C$5&"   ·   Uн = "&TEXT(Uн_лин*1000,"0")&" В"', "SUB"); merges.append("A2:J2")
    T("B3", "Объект:", "LBL"); T("C3", "(название и адрес объекта)", "IN"); merges.append("C3:J3")
    for col in "DEFGHIJ": E(f"{col}3", "IN")
    T("B4", "Щит (ВУ):", "LBL"); F("C4", "IFERROR(Диапазоны!B4&\"\",\"\")", "IN"); merges.append("C4:E4"); E("D4", "IN"); E("E4", "IN")
    T("B5", "Режим:", "LBL"); T("C5", "Весь щит (аварийный)", "IN"); merges.append("C5:E5"); E("D5", "IN"); E("E5", "IN")
    T("G4", "Ко сил. =", "LBL"); F("H4", "Ко_сил", "V2")
    T("G5", "Кс быт (табл.) =", "LBL"); F("H5", ks_interp("SUM($M$11:$M$40)"), "V2")
    T("I4", "Кат. надёжн.:", "LBL"); T("J4", "II", "IN")
    # служебные: границы секций
    T("L4", "секц. от", "G"); F("M4", 'IF($C$5="Секция 2",2,1)', "G")
    T("L5", "секц. до", "G"); F("M5", 'IF($C$5="Секция 1",1,2)', "G")
    # шапка
    heads = [("A7", "№", "A7:A8"), ("B7", "Наименование эл. приёмника", "B7:B8"),
             ("C7", "Установленная активная мощность, кВт", "C7:C8"), ("D7", "Коэффициенты", "D7:F7"),
             ("G7", "Расчётная мощность", "G7:I7"), ("J7", "Расчётный ток I, А", "J7:J8")]
    for ref, t, mg in heads:
        T(ref, t, "H"); merges.append(mg)
    for ref in ("A8", "B8", "C8", "E7", "F7", "H7", "I7", "J8"): E(ref, "H")
    for ref, t in (("D8", "Кс"), ("E8", "cos φ"), ("F8", "tg φ"), ("G8", "активная, кВт"), ("H8", "реактивная, квар"), ("I8", "полная, кВА")):
        T(ref, t, "H")
    for i, col in enumerate("ABCDEFGHIJ"): N(f"{col}9", i + 1, "HN")
    F("A10", "$C$4", "BT"); merges.append("A10:J10")
    for col in "BCDEFGHIJ": E(f"{col}10", "BT")
    # служебная таблица L:W — все категории
    for col, t in zip("LMNOPQRSTUVW", ["Категория", "Руст быт", "Руст сил", "Рр сил (Σ P·Кс)", "Руст пост", "Руст",
                                          "Σ P·cos", "cos φ", "Рр", "Рр с Ко", "Qр с Ко", "№ п/п"]):
        T(f"{col}9", t, "G")
    T("L8", "служебные расчёты (не печатаются)", "G")
    crit = f',{NG}$AJ$3:$AJ$1000,">="&$M$4,{NG}$AJ$3:$AJ$1000,"<="&$M$5'
    def sif(rng, grp, r):
        return f'SUMIFS({NG}${rng}$3:${rng}$1000,{NG}$C$3:$C$1000,$C$4,{NG}$AL$3:$AL$1000,$L{r},{NG}$A$3:$A$1000,"{grp}"{crit})'
    for i in range(NCAT):
        r = 11 + i; k = i + 1
        F(f"L{r}", f'IF(Справочная!$BT${4 + i}="","",Справочная!$BT${4 + i})', "G")
        F(f"M{r}", f'IF($L{r}="",0,{sif("H", "Быт", r)})', "GN")
        F(f"N{r}", f'IF($L{r}="",0,{sif("H", "Сил", r)})', "GN")
        F(f"O{r}", f'IF($L{r}="",0,{sif("AK", "Сил", r)})', "GN")
        F(f"P{r}", f'IF($L{r}="",0,{sif("H", "Пост", r)})', "GN")
        F(f"Q{r}", f"M{r}+N{r}+P{r}", "GN")
        F(f"R{r}", f'IF($L{r}="",0,{sif("L", "Быт", r)}+{sif("L", "Сил", r)}+{sif("L", "Пост", r)})', "GN")
        F(f"S{r}", f'IF(Q{r}=0,"",IF(ISNUMBER(Справочная!$BU${4 + i}),Справочная!$BU${4 + i},IF(R{r}=0,0.95,R{r}/Q{r})))', "GN")
        F(f"T{r}", f"M{r}*$H$5+O{r}+P{r}", "GN")
        F(f"U{r}", f"M{r}*$H$5+Ко_сил*O{r}+P{r}", "GN")
        F(f"V{r}", f'IF(Q{r}=0,0,U{r}*TAN(ACOS(S{r})))', "GN")
        F(f"W{r}", f'IF(Q{r}>0,MAX(W$10:W{r - 1})+1,"")', "GN")
        idx = f"MATCH($A{r},$W$11:$W$40,0)"
        F(f"A{r}", f'IF({k}>MAX($W$11:$W$40),"",{k})', "INT")
        F(f"B{r}", f'IF($A{r}="","",INDEX($L$11:$L$40,{idx}))', "TXT")
        F(f"C{r}", f'IF($A{r}="","",INDEX($Q$11:$Q$40,{idx}))', "N2")
        F(f"D{r}", f'IF($A{r}="","",IF(C{r}=0,0,G{r}/C{r}))', "N2")
        F(f"E{r}", f'IF($A{r}="","",INDEX($S$11:$S$40,{idx}))', "N2")
        F(f"F{r}", f'IF($A{r}="","",TAN(ACOS(E{r})))', "N2")
        F(f"G{r}", f'IF($A{r}="","",INDEX($T$11:$T$40,{idx}))', "N2")
        F(f"H{r}", f'IF($A{r}="","",G{r}*F{r})', "N2")
        F(f"I{r}", f'IF($A{r}="","",SQRT(G{r}^2+H{r}^2))', "N2")
        F(f"J{r}", f'IF($A{r}="","",I{r}/(SQRT(3)*Uн_лин))', "N2")
    T("W10", "№", "G")
    # итоги
    E("A41", "BT"); T("B41", "Итого (сумма по строкам)", "BT")
    F("C41", "SUM(C11:C40)", "BN2"); F("G41", "SUM(G11:G40)", "BN2"); F("H41", "SUM(H11:H40)", "BN2")
    E("A42", "BT"); T("B42", "Расчётная нагрузка щита: Рр = Руст.быт·Кс + Ко·Рр.сил + Рр.пост (СП 256)", "BT")
    F("C42", "C41", "BN2"); F("G42", "SUM(U11:U40)", "BN2"); F("H42", "SUM(V11:V40)", "BN2")
    for r in (41, 42):
        F(f"D{r}", f'IF(C{r}=0,"",G{r}/C{r})', "BN2")
        F(f"I{r}", f"SQRT(G{r}^2+H{r}^2)", "BN2")
        F(f"E{r}", f'IF(I{r}=0,"",G{r}/I{r})', "BN2")
        F(f"F{r}", f'IF(G{r}=0,"",H{r}/G{r})', "BN2")
        F(f"J{r}", f"I{r}/(SQRT(3)*Uн_лин)", "BN2")
    E("A43", "BT"); T("B43", "Компенсация реактивной мощности до tg φ = 0,35, квар", "BT")
    F("C43", 'IF(N(G42)=0,"",IF(G42*(F42-0.35)<50,"не требуется (< 50 квар на ввод, СП 256 разд. 7.3)",ROUND(G42*(F42-0.35),1)))', "BT")
    merges.append("C43:J43")
    for col in "DEFGHIJ": E(f"{col}43", "BT")
    E("A44", "TXT"); T("B44", "Сверка с «Нагрузка_щитов» (Рр того же щита и режима)", "TXT")
    blk = (f'IF($C$4={NG}$R$4,{{c}}10,IF($C$4={NG}$R$13,{{c}}19,IF($C$4={NG}$R$22,{{c}}28,"")))')
    col_by = 'IF($C$5="Секция 1",{s1},IF($C$5="Секция 2",{s2},{a}))'
    ref = col_by.format(a=blk.format(c=NG + "$AF$").replace("{c}", ""), s1=blk.format(c=NG + "$AG$"), s2=blk.format(c=NG + "$AH$"))
    F("C44", f'IFERROR(IF(ABS({ref}-G42)<0.01,"совпадает: "&TEXT(G42,"0.00")&" кВт","расхождение: "&TEXT({ref},"0.00")&" кВт"),"щит не в блоках 1–3 листа «Нагрузка_щитов»")', "TXT")
    merges.append("C44:J44")
    for col in "DEFGHIJ": E(f"{col}44", "TXT")
    NOTES = [
        "Как считается: Руст — сумма мощностей линий категории (без группы «Нет»). Кс = Рр/Руст. Рр строки = Руст.быт·Кс(табл. СП 256 по сумме Руст быт щита) + Σ(P·Кс) силовых + Руст постоянной нагрузки.",
        "Итоговая строка «Расчётная нагрузка щита» учитывает Ко = 0,9 для силовой нагрузки (Рр.ж.д = Ркв + 0,9·Рс). Qр = Рр·tg φ, Sр = √(Рр² + Qр²), Iр = Sр/(√3·Uн).",
        "cos φ строки — средневзвешенный по линиям («Исходные данные» P) или из «Справочной» BU (если задан для категории).",
        "Категории — «Справочная» BT:BV (переименуйте под ПЗ). Категория линии — «Исходные данные» BD (вручную) или по коду потребителя.",
        "Для ПЗ: выберите щит и режим, скопируйте A7:J44 → Вставить как значения в Word/Excel ПЗ. Категория надёжности — ПУЭ 1.2.17–1.2.21, СП 256 разд. 5.",
    ]
    for i, t in enumerate(NOTES):
        r = 46 + i; T(f"A{r}", t, "NOTE"); merges.append(f"A{r}:J{r}")
    # собрать XML
    byrow = {}
    for ref_, xml in cells.items():
        r = int(re.search(r"\d+", ref_).group(0)); byrow.setdefault(r, {})[re.match(r"[A-Z]+", ref_).group(0)] = xml
    rows = []
    for r in sorted(byrow):
        ht = ' ht="30" customHeight="1"' if r in (7, 8) else (' ht="27" customHeight="1"' if 46 <= r <= 50 or r in (42, 43) else "")
        rows.append(f'<row r="{r}"{ht}>' + "".join(byrow[r][c_] for c_ in sorted(byrow[r], key=X.col2n)) + "</row>")
    widths = {"A": 5, "B": 44, "C": 13, "D": 8, "E": 8, "F": 8, "G": 11, "H": 11, "I": 11, "J": 12, "K": 3,
              "L": 30, "M": 9, "N": 9, "O": 9, "P": 9, "Q": 9, "R": 9, "S": 7, "T": 9, "U": 9, "V": 9, "W": 5}
    cols = "".join(f'<col min="{X.col2n(a)}" max="{X.col2n(a)}" width="{w}" customWidth="1"/>' for a, w in widths.items())
    mc = "".join(f'<mergeCell ref="{m}"/>' for m in merges)
    dv = ('<dataValidations count="2">'
          '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="C4"><formula1>Диапазоны!$B$4:$B$33</formula1></dataValidation>'
          '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="C5"><formula1>"Весь щит (аварийный),Секция 1,Секция 2"</formula1></dataValidation>'
          '</dataValidations>')
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           '<sheetPr><tabColor rgb="FFED7D31"/><pageSetUpPr fitToPage="1"/></sheetPr>'
           '<dimension ref="A1:W50"/>'
           '<sheetViews><sheetView workbookViewId="0"><pane ySplit="10" topLeftCell="A11" activePane="bottomLeft" state="frozen"/>'
           '<selection pane="bottomLeft" activeCell="C3" sqref="C3"/></sheetView></sheetViews>'
           '<sheetFormatPr defaultRowHeight="14.4"/>'
           f"<cols>{cols}</cols><sheetData>{''.join(rows)}</sheetData>"
           f'<mergeCells count="{len(merges)}">{mc}</mergeCells>{dv}'
           '<pageMargins left="0.6" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
           '<pageSetup paperSize="9" orientation="portrait" fitToHeight="0"/></worksheet>')
    wr(os.path.join(WD, part), xml)
    wb = rd(wbp)
    k = wb.find("</definedNames>")
    wb = wb[:k] + f'<definedName name="_xlnm.Print_Area" localSheetId="{pos}">Расчёт_нагрузок!$A$1:$J$50</definedName>' + wb[k:]
    wr(wbp, wb)


def guide(WD):
    gd = X.Sheet(X.sheet_path(WD, "Краткое руководство"))
    sE = gd.style("E13"); sF = gd.style("F13"); sB = gd.style("B18"); sC = gd.style("C18")
    bulk(gd, {
        "E15": ct("E15", "v3.10 (нормы, ПЗ)", sE),
        "F15": ct("F15", "Проверка по ПУЭ/СП 256/ГОСТ Р 50571: ток с cos φ и 230/400 В; сечение по ПУЭ 1.3.6 (1ф/3ф, ×0,92 для 4–5 жил, К прокл.); "
                         "ΔU по итоговой длине и ГОСТ Р 50571.5.52 прил. G, норма от ВРУ 3 %/4 %; автоматы до 630 А; ТКЗ — исправлены сумма сопротивлений, Iкз/In и Zт. "
                         "Новый лист «Расчёт_нагрузок» — таблица для ПЗ.", sF),
        "B20": ct("B20", "19", sB),
        "C20": ct("C20", "ПЗ (стадия П): лист «Расчёт_нагрузок» — выбрать щит и режим, категории линий — «Исходные данные» BD (или по коду). "
                         "Настройки сети (Uн, К прокл., ΔU до шин щита) — «Справочная» BQ29:BQ31.", sC)})
    gd.save()


GRSH_CAT = {"ЩОВ-01": 7, "ЩОВ-001": 7, "ЩОВ-002": 7, "ЩОВ-2": 7, "ЩР-001": 18, "ЩР-01": 18, "ЩР-1": 18, "ЩР-2": 18, "ЩО-2": 18,
            "ЩС-Б": 14, "РЩ-1": 14, "Э.4": 14, "ЩС-ВК.001": 15, "ЩС-ВК.002": 15, "ЩС-ВК.003": 15, "ЩУОВ": 16, "ЩР-ПК": 13,
            "SEC1-21": 11, "ЩСПЗ-ЖД.1": 12, "ЩСПЗ-ЖД.2": 12}


def grsh_cats(WD):
    X.load_sst(WD)
    isx = X.Sheet(X.sheet_path(WD, "Исходные данные")); c = {}
    for r in range(3, 60):
        ln = X.value(isx, f"D{r}")
        if ln in GRSH_CAT:
            c[f"BD{r}"] = ct(f"BD{r}", CATS[GRSH_CAT[ln]], isx.style(f"BD{r}"))
    bulk(isx, c); isx.save()
    pz = X.Sheet(X.sheet_path(WD, "Расчёт_нагрузок"))
    bulk(pz, {"C3": ct("C3", "Резиденция, здание жилого дома — ГРЩ-ЖД", pz.style("C3"))}); pz.save()


def build(src, out, grsh=False):
    WD = os.path.join(HERE, "w10_" + os.path.basename(out)[:6].replace(" ", "_"))
    order = P7.unpack(src, WD)
    print("==", os.path.basename(out))
    S = styles_add(WD)
    fix_reference(WD, grsh); fix_isx(WD); fix_sl(WD); fix_tkz(WD); fix_load(WD)
    add_pz_sheet(WD, order, S); guide(WD)
    if grsh: grsh_cats(WD)
    P7.pack(WD, order, out)


if __name__ == "__main__":
    SRC = os.path.join(HERE, "out")
    build(os.path.join(SRC, "Однолинейка_пример_ЖК-Остров.xlsx"), os.path.join(OUT, "Однолинейка_пример_ЖК-Остров.xlsx"))
    build(os.path.join(SRC, "Однолинейка_шаблон.xlsx"), os.path.join(OUT, "Однолинейка_шаблон.xlsx"))
    build(os.path.join(SRC, "Однолинейка ГРЩ-ЖД.xlsx"), os.path.join(OUT, "Однолинейка ГРЩ-ЖД.xlsx"), grsh=True)
    print("готово")
