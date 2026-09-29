"""v3.9: коэффициент спроса (Кс) по СП 256 в «Нагрузке_щитов».

Группы нагрузки: Быт (Кс по таблице квартир/домов повышенной комфортности по сумме Руст быт щита),
Сил (свой Кс линии, «Исходные данные» BA), Пост (Кс = 1), Нет (не учитывается: вводы, резерв, транзит).
Рр = Руст.быт·Кс(табл.) + Ко·Σ(Pсил·Кс) + ΣPпост,  Ко = 0,9 (Справочная BQ14).
Секции 1/2 («Исходные данные» BB) — Рр и Iр по секциям и для аварийного режима (весь щит).
Iр — по наиболее нагруженной фазе. Исправлено: двойной Кс в Iр блоков 2 и 3, cos φ = 0,95 константой в блоке 1,
#Н/Д в cos φ у линий с неизвестным кодом.

Запуск: python3 patch7.py  → out/Однолинейка_шаблон.xlsx, out/Однолинейка_пример_ЖК-Остров.xlsx, out/Однолинейка ГРЩ-ЖД.xlsx
"""
import sys, os, re, shutil, zipfile, html
HERE = os.path.dirname(os.path.abspath(__file__))
REPO_OUT = "/home/claude/maratahmetzyanov83/odnolineyka/chats/2026-09-26_single-line-excel-autocad/output/"
sys.path.insert(0, REPO_OUT + "scripts")
import xlsxlib as X
import openpyxl

SRC_TPL = REPO_OUT + "Однолинейка_шаблон.xlsx"
SRC_EX = REPO_OUT + "Однолинейка_пример_ЖК-Остров.xlsx"
SRC_GRSH = sys.argv[1] if len(sys.argv) > 1 else None
OUT = os.path.join(HERE, "out"); os.makedirs(OUT, exist_ok=True)
esc = lambda s: html.escape(str(s), quote=False)
ISX = "'Исходные данные'!"


def rd(p): return open(p, encoding="utf-8").read()
def wr(p, s): open(p, "w", encoding="utf-8").write(s)
def _s(s): return f' s="{s}"' if s not in (None, "None", "") else ""
def ct(ref, v, s): return f'<c r="{ref}"{_s(s)} t="inlineStr"><is><t xml:space="preserve">{esc(v)}</t></is></c>'
def cn(ref, v, s): return f'<c r="{ref}"{_s(s)}><v>{v}</v></c>'
def cf(ref, f, s): return f'<c r="{ref}"{_s(s)}><f>{esc(f)}</f></c>'
def ce(ref, s): return f'<c r="{ref}"{_s(s)}/>'


def bulk(sh_, cells):
    """вставить/заменить много ячеек за один проход (порядок ячеек в строке сохраняется)"""
    byrow = {}
    for ref, xml in cells.items():
        r = int(re.search(r"\d+", ref).group(0)); byrow.setdefault(r, {})[re.match(r"[A-Z]+", ref).group(0)] = xml
    def fix(mm):
        r = int(mm.group(1)); head = mm.group(0)
        if r not in byrow: return head
        if head.endswith("/>"):
            attrs = head[:-2]; body = ""
        else:
            attrs = head[:head.index(">")]; body = head[head.index(">") + 1:-len("</row>")]
        cur = {cm.group(1): cm.group(0) for cm in re.finditer(r'<c r="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)', body, flags=re.S)}
        cur.update(byrow.pop(r))
        attrs = re.sub(r'\sspans="[^"]*"', "", attrs)
        return attrs + ">" + "".join(cur[c] for c in sorted(cur, key=X.col2n)) + "</row>"
    sh_.s = re.sub(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', fix, sh_.s, flags=re.S)
    for r in sorted(byrow):
        cells_ = byrow[r]
        row = f'<row r="{r}">' + "".join(cells_[c] for c in sorted(cells_, key=X.col2n)) + "</row>"
        rows = [(int(m.group(1)), m.start()) for m in re.finditer(r'<row r="(\d+)"', sh_.s)]
        pos = next((p for n, p in rows if n > r), None)
        if pos is None: pos = sh_.s.find("</sheetData>")
        sh_.s = sh_.s[:pos] + row + sh_.s[pos:]


def set_cols(sh_, widths, unhide=()):
    m = re.search(r"<cols>(.*?)</cols>", sh_.s, flags=re.S)
    cols = re.findall(r"<col [^>]*/>", m.group(1)) if m else []
    tgt = {X.col2n(a) for a in widths}
    keep = []
    for x in cols:
        a = int(re.search(r'min="(\d+)"', x).group(1)); b = int(re.search(r'max="(\d+)"', x).group(1))
        inter = [n for n in range(a, b + 1) if n in tgt]
        if not inter: keep.append(x); continue
        # разрезать диапазон вокруг наших столбцов
        rest = [n for n in range(a, b + 1) if n not in tgt]
        runs = []
        for n in rest:
            if runs and runs[-1][1] == n - 1: runs[-1][1] = n
            else: runs.append([n, n])
        for lo, hi in runs:
            keep.append(re.sub(r'max="\d+"', f'max="{hi}"', re.sub(r'min="\d+"', f'min="{lo}"', x)))
    for a, w in widths.items():
        keep.append(f'<col min="{X.col2n(a)}" max="{X.col2n(a)}" width="{w}" customWidth="1"/>')
    keep.sort(key=lambda x: int(re.search(r'min="(\d+)"', x).group(1)))
    new = "<cols>" + "".join(keep) + "</cols>"
    if m: sh_.s = sh_.s[:m.start()] + new + sh_.s[m.end():]
    else: sh_.s = sh_.s.replace("<sheetData", new + "<sheetData", 1)


def add_dxf(WD, xml):
    stp = WD + "/xl/styles.xml"; st = rd(stp)
    m = re.search(r'<dxfs count="(\d+)"', st); n = int(m.group(1))
    st = st.replace(m.group(0), f'<dxfs count="{n + 1}"', 1)
    k = st.find("</dxfs>"); st = st[:k] + xml + st[k:]
    wr(stp, st); return n


def add_after_cf(sh_, xml):
    """вставить элемент (dataValidations / conditionalFormatting) после последнего conditionalFormatting
    (или после mergeCells/phoneticPr/autoFilter/sheetProtection) — порядок элементов по схеме OOXML"""
    k = sh_.s.rfind("</conditionalFormatting>")
    if k >= 0:
        k += len("</conditionalFormatting>")
    else:
        for tag in ("</mergeCells>", "<phoneticPr", "<autoFilter", "<sheetProtection", "</sheetData>"):
            j = sh_.s.rfind(tag)
            if j >= 0:
                k = sh_.s.find(">", j) + 1 if not tag.startswith("</") else j + len(tag); break
    sh_.s = sh_.s[:k] + xml + sh_.s[k:]


def add_cf(sh_, sqref, formula, dxf, prio=None):
    if prio is None:
        prio = max([int(x) for x in re.findall(r'priority="(\d+)"', sh_.s)] or [0]) + 1
    xml = (f'<conditionalFormatting sqref="{sqref}"><cfRule type="expression" dxfId="{dxf}" priority="{prio}">'
           f'<formula>{esc(formula)}</formula></cfRule></conditionalFormatting>')
    k = sh_.s.rfind("</conditionalFormatting>")
    if k >= 0:
        k += len("</conditionalFormatting>"); sh_.s = sh_.s[:k] + xml + sh_.s[k:]
    else:
        add_after_cf(sh_, xml)


def add_dv(sh_, items):
    """items: [(sqref, formula1)]"""
    body = "".join(f'<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="{sq}"><formula1>{esc(f1)}</formula1></dataValidation>'
                   for sq, f1 in items)
    m = re.search(r'<dataValidations count="(\d+)">', sh_.s)
    if m:
        n = int(m.group(1)) + len(items)
        k = sh_.s.find("</dataValidations>")
        sh_.s = sh_.s[:k] + body + sh_.s[k:]
        sh_.s = sh_.s.replace(m.group(0), f'<dataValidations count="{n}">', 1)
    else:
        add_after_cf(sh_, f'<dataValidations count="{len(items)}">{body}</dataValidations>')


def add_name(WD, name, ref):
    p = WD + "/xl/workbook.xml"; wb = rd(p)
    if f'name="{name}"' in wb:
        wb = re.sub(r'(<definedName name="%s"[^>]*>)[^<]*(</definedName>)' % re.escape(name), lambda mm: mm.group(1) + esc(ref) + mm.group(2), wb)
    else:
        k = wb.find("</definedNames>")
        wb = wb[:k] + f'<definedName name="{name}">{esc(ref)}</definedName>' + wb[k:]
    wr(p, wb)


# ------------------------------------------------------------------ коды → группа по умолчанию
GROUP_BY_CODE = {
    "Ввод": "Нет", "На ИБП": "Нет", "От ИБП": "Нет", "КУП": "Нет", "PE": "Нет",
    "R": "Быт", "L": "Быт", "LED": "Быт", "LD": "Быт", "DALI": "Быт", "SH": "Быт", "В": "Быт", "ТП": "Быт",
    "K": "Быт", "УВ": "Быт", "ПН": "Быт", "КМН": "Быт",
    "R.ШТК": "Пост", "R.ЩОПС": "Пост", "R.GSM": "Пост", "АСУ": "Пост", "PS-KNX": "Пост", "ЭK": "Пост",
    "TE-O": "Пост", "TE-ТП": "Пост", "АПС": "Пост", "WD": "Пост", "VK": "Пост",
    "ПВ": "Сил", "ПУ": "Сил",
}
LAT2CYR = str.maketrans("ABCEHKMOPTXaceopxy", "АВСЕНКМОРТХасеорху")
def norm(x): return str(x).strip().translate(LAT2CYR)
GBC = {norm(k): v for k, v in GROUP_BY_CODE.items()}
P_PAT = ("IF(VLOOKUP(AH{r},Потребители_обозначение,5,FALSE)=2,_xlfn.IFS(H{r}=3,VLOOKUP(E{r},Справочная_двигатели,8,FALSE),"
         "H{r}=1,VLOOKUP(E{r},Справочная_двигатели,10,TRUE)),0.95)")
KS_TABLE = [(0, 0.8), (14, 0.8), (20, 0.65), (30, 0.6), (40, 0.55), (50, 0.5), (60, 0.48), (70, 0.45)]


def ks_interp(x):
    m = f"MATCH({x},Кс_Р,1)"
    return (f"IF({x}>=INDEX(Кс_Р,ROWS(Кс_Р)),INDEX(Кс_К,ROWS(Кс_К)),"
            f"INDEX(Кс_К,{m})+({x}-INDEX(Кс_Р,{m}))*(INDEX(Кс_К,{m}+1)-INDEX(Кс_К,{m}))/(INDEX(Кс_Р,{m}+1)-INDEX(Кс_Р,{m})))")


def apply_kc(WD, template):
    X.load_sst(WD)
    D_RED = add_dxf(WD, '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>')
    D_IN = add_dxf(WD, '<dxf><fill><patternFill><bgColor rgb="FFFFF2CC"/></patternFill></fill></dxf>')

    # ============================================================ 1. «Справочная»: таблица Кс, Ко, группы, группа по коду
    spr = X.Sheet(X.sheet_path(WD, "Справочная"))
    R_T = spr.style("BE2"); R_H = spr.style("BE3"); R_L = spr.style("BE4"); R_V = spr.style("BF8"); R_IN = spr.style("BF4")
    c = {}
    c["AZ2"] = ct("AZ2", "Кс", R_H)
    c["AZ3"] = ct("AZ3", "Группа Кс", R_H)
    for r in range(4, 52):
        code = X.value(spr, f"P{r}")
        key = norm(code) if code else None
        if key and key in GBC:
            c[f"AZ{r}"] = ct(f"AZ{r}", GBC[key], R_IN)
        else:
            if code: print("  код без группы (будет Быт):", code)
            c[f"AZ{r}"] = ce(f"AZ{r}", R_IN)
    c["BP2"] = ct("BP2", "КОЭФФИЦИЕНТ СПРОСА (СП 256.1325800.2016, расчётные нагрузки)", R_T)
    c["BP3"] = ct("BP3", "Руст быт, кВт", R_H); c["BQ3"] = ct("BQ3", "Кс", R_H)
    c["BR3"] = ct("BR3", "Квартиры/дома повышенной комфортности (СП 31-110-2003 табл. 6.2, те же значения в СП 256). Между точками — линейная интерполяция.", R_H)
    for i, (p, k) in enumerate(KS_TABLE):
        r = 4 + i
        c[f"BP{r}"] = cn(f"BP{r}", p, R_IN); c[f"BQ{r}"] = cn(f"BQ{r}", k, R_IN)
    c["BR4"] = ct("BR4", "до 14 кВт — 0,8 (строка 0 кВт нужна для интерполяции)", R_L)
    c["BR11"] = ct("BR11", "70 кВт и более — 0,45", R_L)
    c["BP13"] = ct("BP13", "НАСТРОЙКИ", R_H); c["BQ13"] = ce("BQ13", R_H)
    c["BP14"] = ct("BP14", "Ко сил.", R_L); c["BQ14"] = cn("BQ14", 0.9, R_IN)
    c["BR14"] = ct("BR14", "Коэффициент участия силовой нагрузки в максимуме: Рр = Рр.быт + 0,9·Рр.сил (СП 256, расчёт жилого дома)", R_L)
    c["BP16"] = ct("BP16", "ГРУППЫ", R_H); c["BQ16"] = ce("BQ16", R_H)
    c["BR16"] = ct("BR16", "«Исходные данные»: AZ — группа (пусто = по коду, столбец AZ этой страницы), BA — Кс для «Сил», BB — секция (1/2)", R_H)
    GR = [("Быт", "Бытовая нагрузка дома/квартиры: розетки, свет, тёплый пол, кондиционеры, бытовые щиты ЩР. Кс — по таблице выше по СУММЕ Руст быт щита (не перемножать с Кс нижних щитов)."),
          ("Сил", "Силовая: вентиляция, насосы, пищеблок, бассейн, хаммам… Кс — свой у каждой линии («Исходные данные» BA; пусто = 1). Вентиляция и насосы — по табл. СП 256 для сантехнического и вентиляционного оборудования (по числу двигателей, резервные не учитывать); пищеблок — по табл. для теплового оборудования."),
          ("Пост", "Постоянная, Кс = 1 и без Ко: обогрев кровли/воронок, слаботочные шкафы, АСУ, питание KNX, СПЗ."),
          ("Нет", "Не учитывается в Руст и Рр: вводы, резервные линии и двигатели, транзит.")]
    for i, (g, t) in enumerate(GR):
        r = 17 + i
        c[f"BP{r}"] = ct(f"BP{r}", g, R_V); c[f"BR{r}"] = ct(f"BR{r}", t, R_L)
    c["BP22"] = ct("BP22", "ФОРМУЛА", R_H); c["BQ22"] = ce("BQ22", R_H)
    NOTES = ["Рр = Руст.быт · Кс(табл.) + Ко · Σ(Pсил · Кс) + Σ Pпост;   Кс итоговый = Рр / Руст — он идёт в «Нагрузка_щитов» V9 и в «В Акад».",
             "Iр — по наиболее нагруженной фазе: Iр = Pф.max · Кс итог / (0,22 · cos φ), но не меньше симметричного Рр / (√3 · 0,38 · cos φ).",
             "Секции: «Весь щит» = аварийный режим (все линии на одном вводе) — по нему выбирать вводные аппараты при АВР/секционном; «Секция 1» и «Секция 2» — нормальный режим.",
             "Многоуровневое питание: в ВРУ/ГРЩ линия к бытовому ЩР — группа «Быт» с Руст этого щита; Кс считается заново по сумме, Кс нижних щитов не перемножается."]
    for i, t in enumerate(NOTES):
        c[f"BR{23 + i}"] = ct(f"BR{23 + i}", t, R_L)
    bulk(spr, c)
    set_cols(spr, {"AZ": 10, "BO": 3, "BP": 13, "BQ": 8, "BR": 110})
    spr.save()
    add_name(WD, "Кс_Р", "Справочная!$BP$4:$BP$11")
    add_name(WD, "Кс_К", "Справочная!$BQ$4:$BQ$11")
    add_name(WD, "Ко_сил", "Справочная!$BQ$14")

    # ============================================================ 2. «Исходные данные»: ввод группы, Кс, секции
    isx = X.Sheet(X.sheet_path(WD, "Исходные данные"))
    S_H = isx.style("R2"); S_IN = isx.style("R3"); S_F = isx.style("AT3"); S_1 = isx.style("AT1") or isx.style("A1")
    c = {}
    for i, (col, t) in enumerate([("AZ", "Группа Кс (Быт/Сил/Пост/Нет; пусто = по коду)"), ("BA", "Кс (для Сил)"),
                                  ("BB", "Секция (1/2; пусто = 1)"), ("BC", "Группа Кс (итог)")]):
        c[f"{col}1"] = cn(f"{col}1", 52 + i, S_1)
        c[f"{col}2"] = ct(f"{col}2", t, S_H)
    for mm in re.finditer(r'<c r="P(\d+)"[^>]*?><f[^>]*>([^<]+)</f>', isx.s):
        assert html.unescape(mm.group(2)) == P_PAT.format(r=mm.group(1)), ("P", mm.group(1))
    look = "INDEX(Справочная!$AZ$4:$AZ$51,MATCH(AH{r},Справочная!$P$4:$P$51,0))&\"\""
    n_p = 0
    for r in range(3, 1001):
        for col in ("AZ", "BA", "BB"):
            if isx.is_empty(f"{col}{r}"): c[f"{col}{r}"] = ce(f"{col}{r}", S_IN)
        lk = look.format(r=r)
        c[f"BC{r}"] = cf(f"BC{r}", f'IF(D{r}="","",IF(AZ{r}<>"",AZ{r},IF(ISNUMBER(BA{r}),"Сил",IF(IFERROR({lk},"")="","Быт",{lk}))))', S_F)
        # cos φ: неизвестный код → 0,95 вместо #Н/Д
        # (в столбце P общие формулы t="shared" — переписываем ВСЕ ячейки P3:P1000 явно, иначе группы сломаются)
        m = isx.get(f"P{r}")
        if m and "<f" in m.group(0):
            c[f"P{r}"] = cf(f"P{r}", "IFERROR(" + P_PAT.format(r=r) + ",0.95)", isx.style(f"P{r}")); n_p += 1
    bulk(isx, c)
    add_dv(isx, [("AZ3:AZ1000", '"Быт,Сил,Пост,Нет"'), ("BB3:BB1000", '"1,2"')])
    add_cf(isx, "BA3:BA1000", 'AND(BA3<>"",NOT(ISNUMBER(BA3)))', D_RED)
    set_cols(isx, {"AZ": 16, "BA": 9, "BB": 10, "BC": 11})
    isx.save()
    print("  cos φ обёрнут в IFERROR:", n_p)

    # ============================================================ 3. «Нагрузка_щитов»
    ld = X.Sheet(X.sheet_path(WD, "Нагрузка_щитов"))
    L_H = ld.style("C2"); L_V = ld.style("H3"); L_T = ld.style("G7") or ld.style("C3")
    B_H = ld.style("O4"); B_L = ld.style("U5"); B_V = ld.style("V5")
    c = {}
    c["A2"] = ct("A2", "Группа Кс", L_H); c["B2"] = ct("B2", "Кс", L_H)
    c["AJ2"] = ct("AJ2", "Секция", L_H); c["AK2"] = ct("AK2", "Рр линии (P·Кс), кВт", L_H)
    BLK = [4, 13, 22]
    # в строках 3–6 были стёрты формулы F, G, H (фаза и мощность первых четырёх линий не попадали в расчёт) — восстановить
    n_fix = 0
    for r in range(3, 1001):
        for col, f in (("F", f"VLOOKUP(ROW({r - 2}:{r - 2}),Разбивка_нумерация_5,16,FALSE)"),
                       ("G", f'IF(E{r}="упр","",IF(E{r}=0,F{r},IF(E{r}="1,2,3","L1,L2,L3","L"&E{r})))'),
                       ("H", f"VLOOKUP(D{r},Исходные_данные,5,FALSE)")):
            m = ld.get(f"{col}{r}")
            if not (m and "<f" in m.group(0)):
                c[f"{col}{r}"] = cf(f"{col}{r}", f, ld.style(f"{col}7")); n_fix += 1
    print("  восстановлено формул F/G/H в «Нагрузка_щитов»:", n_fix)
    kbyt = lambda r: (f'IF(C{r}=$R$4,$AF$7,IF(C{r}=$R$13,$AF$16,IF(C{r}=$R$22,$AF$25,1)))')
    for r in range(3, 1001):
        m_ = f"MATCH(D{r},{ISX}$D$1:$D$1000,0)"
        c[f"A{r}"] = cf(f"A{r}", f'IFERROR(INDEX({ISX}$BC$1:$BC$1000,{m_})&"","")', L_T)
        c[f"B{r}"] = cf(f"B{r}", (f'IF(A{r}="","",IF(A{r}="Быт",{kbyt(r)},IF(A{r}="Сил",'
                                  f'IFERROR(IF(ISNUMBER(INDEX({ISX}$BA$1:$BA$1000,{m_})),INDEX({ISX}$BA$1:$BA$1000,{m_}),1),1),'
                                  f'IF(A{r}="Пост",1,0))))'), L_V)
        c[f"AJ{r}"] = cf(f"AJ{r}", f'IF(A{r}="","",IFERROR(IF(INDEX({ISX}$BB$1:$BB$1000,{m_})&""="2",2,1),1))', L_T)
        c[f"AK{r}"] = cf(f"AK{r}", f'IFERROR(H{r}*B{r},0)', L_V)
    Cc, A, H, N, AJ, AK, L = ("$C$3:$C$1000", "$A$3:$A$1000", "$H$3:$H$1000", "$N$3:$N$1000",
                              "$AJ$3:$AJ$1000", "$AK$3:$AK$1000", "$L$3:$L$1000")
    for r in BLK:
        nm = f"$R${r}"
        c[f"AE{r}"] = ct(f"AE{r}", "РАСЧЁТ Рр (СП 256)", B_H)
        for col, t in (("AF", "Весь щит (аварийн.)"), ("AG", "Секция 1"), ("AH", "Секция 2")):
            c[f"{col}{r}"] = ct(f"{col}{r}", t, B_H)
        labels = ["Руст (без «Нет»), кВт", "Руст быт, кВт", "Кс быт (табл. СП 256)", "Рр сил = Σ P·Кс, кВт",
                  "Рр пост (Кс = 1), кВт", "Рр = быт·Кс + Ко·сил + пост, кВт", "Кс итоговый = Рр / Руст",
                  "Iр по наиб. нагруж. фазе, А"]
        for i, t in enumerate(labels):
            c[f"AE{r + 1 + i}"] = ct(f"AE{r + 1 + i}", t, B_L)
        cos = f"$V${r + 4}"
        for col, sec in (("AF", ""), ("AG", f",{AJ},1"), ("AH", f",{AJ},2")):
            sif = lambda rng, g: f'SUMIFS({rng},{Cc},{nm},{A},"{g}"{sec})'
            R = lambda k: f"{col}{r + k}"
            g_ = lambda f: f'IF(ISERROR({nm}),"",{f})'      # щита нет в списке → пусто, а не #Н/Д
            c[R(1)] = cf(R(1), g_(f'{sif(H, "Быт")}+{sif(H, "Сил")}+{sif(H, "Пост")}'), B_V)
            c[R(2)] = cf(R(2), g_(sif(H, "Быт")), B_V)
            c[R(3)] = cf(R(3), g_(ks_interp(R(2))), B_V)
            c[R(4)] = cf(R(4), g_(sif(AK, "Сил")), B_V)
            c[R(5)] = cf(R(5), g_(sif(H, "Пост")), B_V)
            c[R(6)] = cf(R(6), g_(f"{R(2)}*{R(3)}+Ко_сил*{R(4)}+{R(5)}"), B_V)
            c[R(7)] = cf(R(7), g_(f"IF({R(1)}=0,0,{R(6)}/{R(1)})"), B_V)
            ph = lambda k: (f'(SUMIFS({H},{N},{nm}&"L1,L2,L3",{A},"<>Нет"{sec})/3'
                            f'+SUMIFS({H},{N},{nm}&"L{k}",{A},"<>Нет"{sec}))')
            c[R(8)] = cf(R(8), g_(f"IFERROR(MAX(MAX({ph(1)},{ph(2)},{ph(3)})*{R(7)}/(0.22*{cos}),"
                                f"{R(6)}/(1.732*0.38*{cos})),0)"), B_V)
        # основной блок (его читает «В Акад»)
        c[f"V{r + 1}"] = cf(f"V{r + 1}", f"AF{r + 1}", ld.style(f"V{r + 1}"))
        c[f"V{r + 2}"] = cf(f"V{r + 2}", f"AF{r + 6}", ld.style(f"V{r + 2}"))
        c[f"V{r + 3}"] = cf(f"V{r + 3}", f"AF{r + 8}", ld.style(f"V{r + 3}"))
        c[f"V{r + 4}"] = cf(f"V{r + 4}", f"IFERROR(SUMIF({Cc},{nm},{L})/SUMIF({Cc},{nm},{H}),0.95)", ld.style(f"V{r + 4}"))
        c[f"V{r + 5}"] = cf(f"V{r + 5}", f"AF{r + 7}", ld.style(f"V{r + 5}"))
        c[f"U{r + 3}"] = ct(f"U{r + 3}", "Iр=", ld.style(f"U{r + 3}"))
    bulk(ld, c)
    # столбец A был скрыт — открыть
    ld.s = re.sub(r'<col min="1" max="1"[^>]*/>', "", ld.s)
    set_cols(ld, {"A": 9, "B": 6, "AD": 2, "AE": 33, "AF": 18, "AG": 10, "AH": 10, "AI": 2, "AJ": 8, "AK": 12})
    add_cf(ld, "B3:B1000", 'AND($A3="Сил",$B3=1)', D_IN)
    ld.save()

    # ============================================================ 4. «Краткое руководство»
    gd = X.Sheet(X.sheet_path(WD, "Краткое руководство"))
    sE = gd.style("E13"); sF = gd.style("F13"); sB = gd.style("B18"); sC = gd.style("C18")
    c = {"E14": ct("E14", "v3.9 (Кс)", sE),
         "F14": ct("F14", "Коэффициент спроса по СП 256: группы Быт/Сил/Пост/Нет, Кс быт — по таблице по сумме Руст, Ко = 0,9 для силовой, "
                          "секции 1/2 и аварийный режим, Iр по наиболее нагруженной фазе. Исправлены двойной Кс в Iр и cos φ = 0,95 константой.", sF),
         "B19": ct("B19", "18", sB),
         "C19": ct("C19", "КС (v3.9): «Исходные данные» AZ — группа (пусто = по коду), BA — Кс для «Сил» (числом, через запятую!), BB — секция 1/2. "
                          "Расчёт — «Нагрузка_щитов» AE:AH рядом с каждым щитом, пояснения и таблица Кс — «Справочная» BP:BR.", sC)}
    bulk(gd, c)
    gd.save()


def unpack(src, WD):
    shutil.rmtree(WD, ignore_errors=True)
    with zipfile.ZipFile(src) as z:
        z.extractall(WD); return z.namelist()


def fix_dims(WD):
    import glob
    for p in glob.glob(WD + "/xl/worksheets/sheet*.xml"):
        s_ = rd(p)
        refs = re.findall(r'<c r="([A-Z]+)(\d+)"', s_)
        if not refs: continue
        mc = max(X.col2n(a) for a, _ in refs); mr = max(int(b) for _, b in refs)
        col = ""
        while mc: mc, rem = divmod(mc - 1, 26); col = chr(65 + rem) + col
        s2 = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:{col}{mr}"/>', s_, count=1)
        if s2 != s_: wr(p, s2)


def pack(WD, order, out):
    fix_dims(WD); check_shared(WD)
    if os.path.exists(out): os.remove(out)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n_ in order:
            p_ = os.path.join(WD, n_)
            if os.path.exists(p_): z.write(p_, n_)


def title(WD, text):
    gd = X.Sheet(X.sheet_path(WD, "Краткое руководство"))
    bulk(gd, {"C1": ct("C1", text, gd.style("C1") or gd.style("C2"))}); gd.save()


def _shift(f, dr, dc):
    """сдвинуть относительные ссылки A1 в формуле (вне строковых литералов)"""
    parts = f.split('"')
    rx = re.compile(r'(?<![A-Za-zА-Яа-яЁё_\d.!$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-zА-Яа-я_])')
    def rep(m):
        c1, col, r1, row = m.groups()
        n = X.col2n(col) + (0 if c1 else dc); rr = int(row) + (0 if r1 else dr)
        cs = ""
        while n: n, rem = divmod(n - 1, 26); cs = chr(65 + rem) + cs
        return f"{c1}{cs}{r1}{rr}"
    for i in range(0, len(parts), 2):
        parts[i] = rx.sub(rep, parts[i])
    return '"'.join(parts)


def unshare(sh_, cols):
    """общие формулы (t="shared") в столбцах cols → обычные, чтобы можно было писать поверх любой ячейки"""
    cols = set(cols); masters = {}
    for m in re.finditer(r'<c r="([A-Z]+)(\d+)"[^>]*><f t="shared" ref="([^"]*)" si="(\d+)">([^<]*)</f>', sh_.s):
        masters[m.group(4)] = (m.group(1), int(m.group(2)), html.unescape(m.group(5)))
    todo = {si for si, (c_, r_, f_) in masters.items() if c_ in cols}
    def fix(m):
        cell = m.group(0); col, row = m.group(1), int(m.group(2))
        fm = re.search(r'<f t="shared"(?: ref="[^"]*")? si="(\d+)"\s*(?:/>|>[^<]*</f>)', cell)
        if not fm or fm.group(1) not in todo: return cell
        mc, mr, mf = masters[fm.group(1)]
        f = _shift(mf, row - mr, X.col2n(col) - X.col2n(mc))
        return cell.replace(fm.group(0), f"<f>{esc(f)}</f>")
    sh_.s = re.sub(r'<c r="([A-Z]+)(\d+)"[^>]*?(?:/>|>.*?</c>)', fix, sh_.s, flags=re.S)
    return len(todo)


def check_shared(WD):
    import glob
    for p in glob.glob(WD + "/xl/worksheets/sheet*.xml"):
        s_ = rd(p)
        masters = set(re.findall(r'<f t="shared" ref="[^"]*" si="(\d+)"', s_))
        orphan = [x for x in re.findall(r'<f t="shared" si="(\d+)"\s*/>', s_) if x not in masters]
        assert not orphan, (p, "общие формулы без мастера", orphan[:5])


# ------------------------------------------------------------------ ГРЩ-ЖД из файла Марата
def fill_grsh(WD, src):
    X.load_sst(WD)
    wb = openpyxl.load_workbook(src)
    si = wb["Исходные данные"]; sr = wb["Разбивка_по_щитам"]
    # его «Исходные данные» имеют лишний столбец Q «Коэффициент спроса» → R..T сдвинуты на 1
    MAP = {"C": "C", "D": "D", "E": "E", "F": "F", "G": "G", "H": "H", "I": "I", "J": "J", "K": "K", "L": "L",
           "N": "N", "O": "O", "R": "Q", "S": "R", "T": "S"}
    GROUP = {"Ввод.1": "Нет", "Ввод.2": "Нет",
             "ЩР-001": "Быт", "ЩР-01": "Быт", "ЩР-1": "Быт", "ЩР-2": "Быт", "ЩО-2": "Быт",
             "ЩОВ-01": "Сил", "ЩОВ-001": "Сил", "ЩОВ-002": "Сил", "ЩОВ-2": "Сил", "ЩС-Б": "Сил",
             "ЩС-ВК.001": "Сил", "ЩС-ВК.002": "Сил", "ЩС-ВК.003": "Сил", "ЩР-ПК": "Сил", "РЩ-1": "Сил", "Э.4": "Сил",
             "ЩУОВ": "Пост", "SEC1-21": "Пост", "ЩСПЗ-ЖД.1": "Пост", "ЩСПЗ-ЖД.2": "Пост"}
    # секции — из его «Нагрузка_щитов» E (там по строкам листа нагрузок)
    sl = openpyxl.load_workbook(src, data_only=True)["Нагрузка_щитов"]
    SEC = {}
    for r in range(3, 200):
        ln = sl[f"D{r}"].value; t = sl[f"E{r}"].value
        if ln and ln != "#N/A" and t:
            SEC[ln] = 2 if "2" in str(t) else 1
    SEC["ЩСПЗ-ЖД.2"] = 2
    isx = X.Sheet(X.sheet_path(WD, "Исходные данные"))
    print("  общих формул развёрнуто:", unshare(isx, set(MAP.values())))
    c = {}
    lines = []
    for r in range(3, 60):
        ln = si[f"D{r}"].value
        for sc, dc in MAP.items():
            v = si[f"{sc}{r}"].value
            ref = f"{dc}{r}"; st = isx.style(ref)
            if v is None:
                if ln: c[ref] = ce(ref, st)      # как в его файле: пусто (а не формула-заготовка шаблона)
                continue
            if isinstance(v, str) and v.startswith("="):
                f = v[1:]
                # в его книге ссылки на свои столбцы: сдвиг R..T не затрагивает формулы J/K (ссылаются на J, M)
                assert not re.search(r"\b[QRST]\d", f), (ref, f)
                c[ref] = cf(ref, f, st)
            elif isinstance(v, (int, float)):
                c[ref] = cn(ref, v, st)
            else:
                c[ref] = ct(ref, v, st)
        if ln:
            lines.append(ln)
            g = GROUP.get(ln)
            if g: c[f"AZ{r}"] = ct(f"AZ{r}", g, isx.style(f"AZ{r}"))
            kc = si[f"Q{r}"].value
            if g == "Сил" and kc not in (None, ""):
                c[f"BA{r}"] = cn(f"BA{r}", float(str(kc).replace(",", ".")), isx.style(f"BA{r}"))
            if ln in SEC:
                c[f"BB{r}"] = cn(f"BB{r}", SEC[ln], isx.style(f"BB{r}"))
    bulk(isx, c); isx.save()
    print("  линий перенесено:", len(lines), lines)
    rz = X.Sheet(X.sheet_path(WD, "Разбивка_по_щитам"))
    c = {}
    for r in range(3, 1001):
        for col in ("F", "G", "H"):
            v = sr[f"{col}{r}"].value
            if v not in (None, ""):
                c[f"{col}{r}"] = ct(f"{col}{r}", v, rz.style(f"{col}{r}"))
    bulk(rz, c); rz.save()
    # марки кабеля, которых нет в шаблоне
    spr = X.Sheet(X.sheet_path(WD, "Справочная"))
    have = {X.value(spr, f"L{r}") for r in range(4, 52)}
    free = [r for r in range(4, 52) if not X.value(spr, f"L{r}")]
    sp = wb["Справочная"]; c = {}
    for r in range(4, 52):
        code = sp[f"L{r}"].value
        if code and code not in have:
            rr = free.pop(0)
            c[f"L{rr}"] = ct(f"L{rr}", code, spr.style(f"L{rr}") or spr.style("L16"))
            c[f"M{rr}"] = ct(f"M{rr}", sp[f"M{r}"].value, spr.style(f"M{rr}") or spr.style("M16"))
            if sp[f"N{r}"].value is not None:
                c[f"N{rr}"] = cn(f"N{rr}", sp[f"N{r}"].value, spr.style(f"N{rr}") or spr.style("N16"))
            print("  марка кабеля добавлена:", code)
    bulk(spr, c); spr.save()


if __name__ == "__main__":
    # --- пример
    WD = os.path.join(HERE, "w_ex"); order = unpack(SRC_EX, WD)
    print("пример"); apply_kc(WD, template=False)
    pack(WD, order, os.path.join(OUT, "Однолинейка_пример_ЖК-Остров.xlsx"))
    # --- шаблон
    WD = os.path.join(HERE, "w_tpl"); order = unpack(SRC_TPL, WD)
    print("шаблон"); apply_kc(WD, template=True)
    title(WD, "ШАБЛОН однолинейки v3.9 — копируйте на каждый объект/щит (имя файла без номера версии). Пример заполнения — «Однолинейка_пример_ЖК-Остров.xlsx».")
    pack(WD, order, os.path.join(OUT, "Однолинейка_шаблон.xlsx"))
    # --- ГРЩ-ЖД
    if SRC_GRSH:
        WD2 = os.path.join(HERE, "w_grsh"); shutil.rmtree(WD2, ignore_errors=True); shutil.copytree(WD, WD2)
        print("ГРЩ-ЖД"); fill_grsh(WD2, SRC_GRSH)
        title(WD2, "ГРЩ-ЖД (резиденция) — однолинейка v3.9. Данные перенесены из вашей таблицы ГРЩ 29.09.2026; Кс — по СП 256 (лист «Нагрузка_щитов» AE:AH).")
        pack(WD2, order, os.path.join(OUT, "Однолинейка ГРЩ-ЖД.xlsx"))
    print("готово")
