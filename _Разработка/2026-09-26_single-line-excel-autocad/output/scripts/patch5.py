"""v3.7: спецификация собирается сама из Каталога (авто + вручную), понятный «Справочник_схемы»,
УГО по-русски + свои блоки, клеммы 3L на линию по коду (шторы SH = 2×3L)."""
import sys, os, re, shutil, zipfile, html
SP = "/tmp/claude-0/-home-claude/5e645175-0e52-5350-b586-c0cb5927e030/scratchpad/"
sys.path.insert(0, SP + "ost")
import xlsxlib as X

SRC = SP + "ost/Однолинейка ЩР ЭОМ_v3.6.xlsx"
OLD = SP + "ost/Однолинейка ЩР ЭОМ_v3.5.1.xlsx"      # старая ручная спецификация
OUT = SP + "ost/Однолинейка ЩР ЭОМ_v3.7.xlsx"
WD = SP + "ost/w7"; WO = SP + "ost/w7old"
for d in (WD, WO): shutil.rmtree(d, ignore_errors=True)
with zipfile.ZipFile(SRC) as z:
    z.extractall(WD); order = z.namelist()
with zipfile.ZipFile(OLD) as z:
    z.extractall(WO)
X.load_sst(WD)
esc = lambda s: html.escape(str(s), quote=False)
O = "Однолинейка!"; K_ = "Каталог!"
CL = 500

# ---------------------------------------------------------------- стили
stp = WD + "/xl/styles.xml"; st = open(stp, encoding="utf-8").read()
def add(tag, xml):
    global st
    m = re.search(r'<%s count="(\d+)"' % tag, st); n = int(m.group(1))
    st = st.replace(m.group(0), '<%s count="%d"' % (tag, n + 1), 1)
    k = st.find("</%s>" % tag); st = st[:k] + xml + st[k:]
    return n
D_RED = add("dxfs", '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>')
D_YEL = add("dxfs", '<dxf><font><color rgb="FF7F6000"/></font><fill><patternFill><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>')
open(stp, "w", encoding="utf-8").write(st)

cat_old = X.Sheet(X.sheet_path(WD, "Каталог"))
S_TITLE = cat_old.style("A1"); S_HEAD = cat_old.style("A3"); S_IN = cat_old.style("A4"); S_CALC = cat_old.style("U4")
S_LBL = cat_old.style("O4")


class Grid:
    def __init__(self): self.rows = {}
    def _p(self, ref, xml):
        r = int(re.search(r"\d+", ref).group(0)); self.rows.setdefault(r, {})[re.match(r"[A-Z]+", ref).group(0)] = xml
    def t(self, ref, v, s): self._p(ref, f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{esc(v)}</t></is></c>')
    def n(self, ref, v, s): self._p(ref, f'<c r="{ref}" s="{s}"><v>{v}</v></c>')
    def f(self, ref, v, s): self._p(ref, f'<c r="{ref}" s="{s}"><f>{esc(v)}</f></c>')
    def e(self, ref, s): self._p(ref, f'<c r="{ref}" s="{s}"/>')
    def v(self, ref, v, s):
        if v in (None, ""): self.e(ref, s)
        elif isinstance(v, (int, float)): self.n(ref, v, s)
        else: self.t(ref, v, s)
    def xml(self):
        out = []
        for r in sorted(self.rows):
            cells = self.rows[r]
            out.append(f'<row r="{r}">' + "".join(cells[c] for c in sorted(cells, key=X.col2n)) + "</row>")
        return "".join(out)


def cols_xml(widths, hidden=()):
    out = []
    for c, w in sorted(widths.items(), key=lambda t: X.col2n(t[0].split(":")[0])):
        a, b = (c.split(":") + [c])[:2] if ":" in c else (c, c)
        hid = ' hidden="1" outlineLevel="1"' if c in hidden else ""
        out.append(f'<col min="{X.col2n(a)}" max="{X.col2n(b)}" width="{w}" customWidth="1"{hid}/>')
    return "<cols>" + "".join(out) + "</cols>"


def sheet_xml(grid, cols, extra_after="", views='<sheetViews><sheetView workbookViewId="0"/></sheetViews>', pr=""):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'{pr}{views}<sheetFormatPr defaultRowHeight="15"/>{cols}<sheetData>{grid.xml()}</sheetData>{extra_after}'
            '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>')


# ================================================================= данные: Каталог + ручные строки старой спецификации
def rows_of(sh, cols, r1, r2):
    return [[X.value(sh, f"{c}{r}") for c in cols] for r in range(r1, r2 + 1)]
CAT = [r for r in rows_of(cat_old, "ABCDEFGHIJKLM", 4, 200) if r[0] or r[2]]
for r in CAT:                          # числа из текста
    for i in (4, 9, 11):
        try: r[i] = float(r[i]) if r[i] not in (None, "") and "." in str(r[i]) else (int(r[i]) if r[i] not in (None, "") else r[i])
        except ValueError: pass
AUTO_CATS = ("Автомат", "Дифавтомат", "Аппарат", "Модуль")
AUTO_TERM = ("VPR-2.5-3L-GY", "VPR-2.5-QUATTRO-GY")
old_sp = X.Sheet(X.sheet_path(WO, "Спецификация"))
# ssts старого файла (другие индексы)
X.load_sst(WO)
OLDSPEC = []
for r in range(4, 95):
    C_, E_, F_, G_, H_ = (X.value(old_sp, f"{c}{r}") for c in "CEFGH")
    if C_: OLDSPEC.append((r, C_, E_, F_, G_, H_))
X.load_sst(WD)
qty_manual = {}           # артикул каталога -> кол-во вручную (для неавтоматических клемм)
by_art = {str(r[7]).strip(): r for r in CAT if r[7]}
migr = []
def cat_of(r):
    if r <= 19: return "Корпус"
    if r <= 44: return "Шины"
    if r <= 61: return "Клемма"
    if r <= 81: return "Провод"
    return "Прочее"
for (r, C_, E_, F_, G_, H_) in OLDSPEC:
    art = str(E_).strip() if E_ is not None else ""
    try: q = float(H_); q = int(q) if q == int(q) else q
    except (TypeError, ValueError): q = H_
    if art in by_art:
        row = by_art[art]
        if row[0] not in AUTO_CATS and art not in AUTO_TERM:
            qty_manual[art] = q
        continue
    migr.append([cat_of(r), F_ or "", art or C_[:30], "", "", "", C_, art, "", "", "", "", G_ or "шт.", q])
print("catalog", len(CAT), "migrated", len(migr), "manual qty for", qty_manual)

# ================================================================= 1. КАТАЛОГ (пересобрать)
g = Grid()
g.t("A1", "КАТАЛОГ ОБОРУДОВАНИЯ. Спецификация собирается отсюда: N — посчитано по однолинейке, O — впишите вручную (корпус, провод, запас). Добавляйте строки вниз.", S_TITLE)
HEAD = ["Категория (раздел)", "Производитель", "Модель (как в таблице)", "Префикс модуля", "Каналов", "Тип каналов",
        "Наименование для спецификации", "Артикул", "Полюса", "Номинал, А", "30 мА", "Ширина, мод. DIN", "Ед.",
        "По однолинейке (авто)", "Вручную / запас (+)", "ИТОГО в спецификацию"]
for i, h in enumerate(HEAD): g.t(f"{'ABCDEFGHIJKLMNOP'[i]}3", h, S_HEAD)
data = [r + [None] for r in CAT] + migr
for i, rec in enumerate(data):
    k = 4 + i
    for j, c in enumerate("ABCDEFGHIJKLM"): g.v(f"{c}{k}", rec[j], S_IN)
    mq = rec[13] if rec[13] is not None else qty_manual.get(str(rec[7]).strip() if rec[7] else "")
    g.v(f"O{k}", mq, S_IN)
for k in range(4 + len(data), CL + 1):
    for c in "ABCDEFGHIJKLMO": g.e(f"{c}{k}", S_IN)
# настройки R:S
SETT = [("Производитель автоматов", "Chint"), ("Производитель аппаратов", "(все)"), ("Производитель модулей", "(все)"),
        ("Клемма трёхуровневая (модель)", "VPR-2.5-3L-GY"), ("Клемма проходная доп. (модель)", "VPR-2.5-QUATTRO-GY"),
        ("Трёхуровневых на 4–5-жильную линию, если по коду не задано", 1)]
g.t("R3", "НАСТРОЙКИ", S_HEAD); g.t("S3", "Значение", S_HEAD)
for i, (a, b) in enumerate(SETT):
    g.t(f"R{4+i}", a, S_LBL); g.v(f"S{4+i}", b, S_IN)
g.t("R11", "ПОРЯДОК РАЗДЕЛОВ СПЕЦИФИКАЦИИ (можно менять и дописывать)", S_HEAD); g.e("S11", S_HEAD)
SECT = ["Корпус", "Автомат", "Дифавтомат", "Аппарат", "Модуль", "Шины", "Клемма", "Провод", "Прочее"]
for i in range(14):
    g.v(f"R{12+i}", SECT[i] if i < len(SECT) else None, S_IN)
g.t("R27", "ЕСТЬ В ОДНОЛИНЕЙКЕ, НО НЕТ В КАТАЛОГЕ — добавьте строку:", S_HEAD); g.e("S27", S_HEAD)
for i in range(10):
    g.f(f"R{28+i}", f'IFERROR(INDEX($AG$4:$AG$43,MATCH({i+1},$AH$4:$AH$43,0)),"")', S_CALC)
# вычисляемые N, P и служебные U..AH
for k in range(4, CL + 1):
    p = k - 1
    g.f(f"N{k}", (f'IF(OR(A{k}="Автомат",A{k}="Дифавтомат"),IF(W{k}="","",COUNTIF({O}$H$3:$H$1000,U{k})),'
                  f'IF(A{k}="Аппарат",IF(C{k}="","",COUNTIF({O}$I$3:$I$1000,C{k})+COUNTIF({O}$J$3:$J$1000,C{k})+COUNTIF({O}$K$3:$K$1000,C{k})),'
                  f'IF(A{k}="Модуль",IF(C{k}="","",COUNTIFS({O}$AW$3:$AW$1000,"?*",{O}$AI$3:$AI$1000,C{k})),'
                  f'IF(A{k}="Клемма",IF(C{k}="","",IF(C{k}=$S$7,SUM({O}$BW$3:$BW$1000),IF(C{k}=$S$8,SUM({O}$BX$3:$BX$1000),""))),""))))'), S_CALC)
    g.f(f"P{k}", f'IF(COUNT(N{k},O{k})=0,"",SUM(N{k},O{k}))', S_CALC)
    F = {
     "U": f'IF(OR(A{k}="Автомат",A{k}="Дифавтомат"),I{k}&" С"&J{k}&"А"&IF(K{k}="да"," 30мА",""),"")',
     "V": f'IF(U{k}="","",B{k}&"|"&U{k})',
     "W": (f'IF(AND(U{k}<>"",OR($S$4="(все)",B{k}=$S$4),IF($S$4="(все)",COUNTIF($U$4:U{k},U{k}),'
           f'COUNTIFS($U$4:U{k},U{k},$B$4:B{k},B{k}))=1),MAX($W$3:W{p})+1,"")'),
     "X": f'IFERROR(INDEX($U$4:$U${CL},MATCH(ROW()-3,$W$4:$W${CL},0)),"")',
     "Y": f'IF(AND(A{k}="Аппарат",C{k}<>"",OR($S$5="(все)",B{k}=$S$5)),MAX($Y$3:Y{p})+1,"")',
     "Z": f'IFERROR(INDEX($C$4:$C${CL},MATCH(ROW()-3,$Y$4:$Y${CL},0)),"")',
     "AA": f'IF(AND(A{k}="Модуль",C{k}<>"",OR($S$6="(все)",B{k}=$S$6)),MAX($AA$3:AA{p})+1,"")',
     "AB": f'IFERROR(INDEX($C$4:$C${CL},MATCH(ROW()-3,$AA$4:$AA${CL},0)),"")',
     "AC": f'IF(OR(B{k}="",COUNTIF($B$4:B{k},B{k})>1),"",MAX($AC$3:AC{p})+1)',
     "AD": f'IFERROR(INDEX($B$4:$B${CL},MATCH(ROW()-3,$AC$4:$AC${CL},0)),"")',
     "AE": f'IF(AND(A{k}<>"",N(P{k})>0),IFERROR(MATCH(A{k},$R$12:$R$25,0),99)*1000+ROW(),"")',
     "AF": f'IF(AE{k}="","",COUNTIF($AE$4:$AE${CL},"<="&AE{k}))',
    }
    for c, f in F.items(): g.f(f"{c}{k}", f, S_CALC)
# «нет в каталоге»: 40 позиций подсчёта (Спецификация L4:L23, L27:L36, L39:L48)
SPROWS = list(range(4, 24)) + list(range(27, 37)) + list(range(39, 49))
for i, sr in enumerate(SPROWS):
    k = 4 + i
    it = f"Спецификация!$L${sr}"
    if sr <= 23:
        miss = f'IF($S$4="(все)",COUNTIF($U$4:$U${CL},{it}),COUNTIF($V$4:$V${CL},$S$4&"|"&{it}))=0'
    else:
        miss = f'COUNTIF($C$4:$C${CL},{it})=0'
    g.f(f"AG{k}", f'IF({it}="","",IF({miss},{it},""))', S_CALC)
    g.f(f"AH{k}", f'IF(AG{k}="","",COUNTIF($AG$4:AG{k},"?*"))', S_CALC)
for c, t in (("U", "обозначение"), ("V", "ключ"), ("W", "№ авт."), ("Y", "№ апп."), ("AA", "№ мод."), ("AC", "№ произв."),
             ("AE", "ключ сортировки"), ("AF", "№ в спецификации"), ("AG", "нет в каталоге"), ("AH", "№")):
    g.t(f"{c}3", t, S_HEAD)
for c, t in (("X", "список автоматов"), ("Z", "список аппаратов"), ("AB", "список модулей"), ("AD", "список произв.")):
    g.t(f"{c}2", t, S_HEAD)
g.t("X3", "авто", S_CALC); g.t("Z3", "список аппаратов", S_HEAD); g.t("AB3", "-", S_CALC); g.t("AD3", "(все)", S_CALC)

widths = {"A": 13, "B": 12, "C": 24, "D": 8, "E": 8, "F": 11, "G": 56, "H": 18, "I": 7, "J": 8, "K": 6, "L": 8, "M": 7,
          "N": 11, "O": 11, "P": 11, "Q": 2, "R": 46, "S": 22, "T": 2, "U:AH": 10}
catdv = ('<dataValidations count="4">'
         '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="A4:A500"><formula1>$R$12:$R$25</formula1></dataValidation>'
         '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="K4:K500"><formula1>"да,нет"</formula1></dataValidation>'
         '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="I4:I500"><formula1>"1P,2P,3P,4P"</formula1></dataValidation>'
         '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="S4:S6"><formula1>Список_произв</formula1></dataValidation>'
         '</dataValidations>')
catcf = (f'<conditionalFormatting sqref="R28:R37"><cfRule type="expression" dxfId="{D_RED}" priority="1"><formula>$R28&lt;&gt;""</formula></cfRule></conditionalFormatting>'
         f'<conditionalFormatting sqref="P4:P500"><cfRule type="expression" dxfId="{D_YEL}" priority="2"><formula>N($P4)&gt;0</formula></cfRule></conditionalFormatting>')
cat_xml = sheet_xml(g, cols_xml(widths, hidden=("U:AH",)),
                    extra_after=f'<autoFilter ref="A3:P{CL}"/>{catcf}{catdv}',
                    views='<sheetViews><sheetView workbookViewId="0"><pane xSplit="3" ySplit="3" topLeftCell="D4" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>',
                    pr='<sheetPr><tabColor rgb="FF4472C4"/><outlinePr summaryRight="1"/></sheetPr>')
open(cat_old.path, "w", encoding="utf-8").write(cat_xml)

# ================================================================= 2. СПРАВОЧНИК_СХЕМЫ (пересобрать понятно)
ref_old = X.Sheet(X.sheet_path(WD, "Справочник_схемы"))
R_T = ref_old.style("J1"); R_H = ref_old.style("M3"); R_L = ref_old.style("J3"); R_V = ref_old.style("K3")
NOTES = [X.value(ref_old, f"J{r}") for r in range(17, 27)]
g = Grid()
g.t("A1", "СПРАВОЧНИК СХЕМЫ — настройки, проверка и свободные каналы (заполнять только жёлтые ячейки B4 и B5)", R_T)
g.t("A3", "НАСТРОЙКИ", R_H); g.e("B3", R_H)
g.t("A4", "Мест (столбцов схемы) на лист A3", R_L); g.n("B4", 17, S_IN)
g.t("A5", "Мин. автомат перед БП / трансформатором, А", R_L); g.n("B5", 10, S_IN)
g.t("A7", "ИТОГИ И ПРОВЕРКА ОДНОЛИНЕЙКИ", R_H); g.e("B7", R_H)
ITOG = [("Автоматов всего", f'COUNTIF({O}$E$3:$E$1000,"?*")'),
        ("  из них с 30 мА (QFD)", f'COUNTIF({O}$AJ$3:$AJ$1000,"QFD*")'),
        ("Клемм XT (номеров)", f'COUNTIF({O}$AY$3:$AY$1000,1)'),
        ("  трёхуровневых блоков", f'SUM({O}$BW$3:$BW$1000)'),
        ("  доп. проходных", f'SUM({O}$BX$3:$BX$1000)'),
        ("Модулей всего", f'COUNTIF({O}$AW$3:$AW$1000,"?*")'),
        ("Строк с ошибками (красные в AE)", f'COUNTIF({O}$BU$3:$BU$1000,2)'),
        ("Строк с предупреждениями (жёлтые в AE)", f'COUNTIF({O}$BU$3:$BU$1000,1)')]
for i, (a, f) in enumerate(ITOG):
    g.t(f"A{8+i}", a, R_L); g.f(f"B{8+i}", f, R_V)
g.t("D3", "МОДУЛИ В ПРОЕКТЕ — СВОБОДНЫЕ КАНАЛЫ", R_H)
for c in "EFGHIJ": g.e(f"{c}3", R_H)
for c, t in zip("DEFGHIJ", ["Щит", "Модуль", "Модель", "Каналов", "Занято (посл. №)", "Резерв «р»", "Свободно"]):
    g.t(f"{c}4", t, R_H)
AG = lambda k: f'(MATCH({k},{O}$BS$1:$BS$1000,0))'
for k in range(1, 36):
    r = 4 + k
    g.f(f"D{r}", f'IFERROR(INDEX({O}$A$1:$A$1000,{AG(k)}),"")', R_V)
    g.f(f"E{r}", f'IFERROR(INDEX({O}$AW$1:$AW$1000,{AG(k)}),"")', R_V)
    g.f(f"F{r}", f'IF(E{r}="","",INDEX({O}$AI$1:$AI$1000,{AG(k)}))', R_V)
    g.f(f"G{r}", f'IF(E{r}="","",IFERROR(VLOOKUP(F{r},Модули_спр,3,FALSE),""))', R_V)
    g.f(f"H{r}", f'IF(E{r}="","",_xlfn.MAXIFS({O}$O$3:$O$1000,{O}$AH$3:$AH$1000,E{r},{O}$A$3:$A$1000,D{r}))', R_V)
    g.f(f"I{r}", f'IF(E{r}="","",COUNTIFS({O}$AH$3:$AH$1000,E{r},{O}$A$3:$A$1000,D{r},{O}$N$3:$N$1000,"р"))', R_V)
    g.f(f"J{r}", f'IF(OR(E{r}="",G{r}=""),"",G{r}-H{r})', R_V)
g.t("D41", "Свободных каналов всего", R_H)
for c in "EFGHI": g.e(f"{c}41", R_H)
g.f("J41", "SUM(J5:J39)", R_H)
g.t("A18", "ГДЕ ЧТО ЛЕЖИТ", R_H); g.e("B18", R_H)
WHERE = ["Оборудование, артикулы, кол-во вручную → лист «Каталог»",
         "Коды потребителей → УГО, 30 мА, мин. автомат, клеммы 3L → «Справочная», столбцы AU:AY",
         "Значки УГО и имена блоков AutoCAD → «Справочная», столбцы BA:BC",
         "Статус каждой строки → «Однолинейка», столбец AE"]
for i, t in enumerate(WHERE): g.t(f"A{19+i}", t, R_L)
g.t("A24", "КАК ПОДБИРАЕТСЯ АВТОМАТ («авто»)", R_H); g.e("B24", R_H)
NOTES = [
    "1. Ток группы = сумма токов линий группы (линия в нескольких строках считается один раз).",
    "2. Номинал по току — по таблице «Автомат_номиналы» (Справочная F:G).",
    "3. Итог = наибольший из: по току, «Мин. автомат» по коду потребителя (Справочная AX: розетки 16 А, свет 10 А) и ручного автомата из «Исходных данных».",
    "4. «Макс. автомат по кабелю» — по сечению (Справочная I:J). Если итог больше — «КАБЕЛЬ НЕ ЗАЩИЩЁН».",
    "5. Если после автомата БП/трансформатор — берётся ток и мин. номинал B5.",
    "6. Полюса: 3 фазы → 3P (4P с 30 мА), 1 фаза → 1P (2P с 30 мА).",
    "7. 30 мА: столбец F «Однолинейки» или по коду (Справочная AW).",
    "8. Свой текст в «Автомат» (например «2P С20А 30мА») — используется как есть и проверяется."]
for i, t in enumerate(NOTES): g.t(f"A{25+i}", t, R_L)
ref_xml = sheet_xml(g, cols_xml({"A": 46, "B": 10, "C": 3, "D": 8, "E": 9, "F": 14, "G": 9, "H": 11, "I": 10, "J": 10}),
                    pr='<sheetPr><tabColor rgb="FF70AD47"/></sheetPr>')
open(ref_old.path, "w", encoding="utf-8").write(ref_xml)

# ================================================================= 3. СПРАВОЧНАЯ: УГО по-русски, 3L на линию, таблица значков
spr = X.Sheet(X.sheet_path(WD, "Справочная"))
RUS = {"SOCKET": "Розетка", "LAMP": "Светильник", "SERVO": "Сервопривод", "VALVE": "Клапан",
       "MOTOR": "Привод (двигатель)", "HEAT": "Нагреватель", "BOX": "Прочее (прямоуг. с подписью)"}
for r in range(4, 73):
    v = X.value(spr, f"AV{r}")
    if v in RUS: spr.text(f"AV{r}", RUS[v])
s_h = spr.style("AU3"); s_in = spr.style("AV7") or spr.style("AU4"); s_v = spr.style("AU4")
spr.text("AY3", "Трёхуровн. клемм на линию", st=s_h)
for r in range(4, 73):
    code = X.value(spr, f"AU{r}") if r >= 53 else X.value(spr, f"P{r}")
    if code == "SH": spr.num(f"AY{r}", 2, st=s_in)
    elif r != 53: spr.clear(f"AY{r}") if not spr.is_empty(f"AY{r}") else spr.put(f"AY{r}", f'<c r="AY{r}" s="{s_in}"/>')
spr.text("AU75", "УГО — выбирается из списка (таблица значков BA:BC). Пустое УГО (красное) — у линии нет значка. "
                 "AY: сколько трёхуровневых клемм занимает одна линия (шторы SH = 2); пусто — по жильности кабеля.", st=s_v)
# таблица значков BA:BC
spr.text("BA2", "ЗНАЧКИ УГО ДЛЯ ОДНОЛИНЕЙКИ", st=spr.style("AU2"))
for c, t in zip(["BA", "BB", "BC"], ["УГО (выбирается в AV)", "Блок AutoCAD (из 02_Однолинейка.dwg)", "Код программы (не менять)"]):
    spr.text(f"{c}3", t, st=s_h)
ICONS = [("Розетка", "Розетка1", "SOCKET"), ("Светильник", "$RECOVER_230831092053-0", "LAMP"),
         ("Сервопривод", "Сервопривод-1", "SERVO"), ("Клапан", "EK.1", "VALVE"), ("Привод (двигатель)", "Привод", "MOTOR"),
         ("Нагреватель", "(условный значок программы)", "HEAT"), ("Прочее (прямоуг. с подписью)", "(условный значок программы)", "BOX")]
for i in range(27):
    r = 4 + i
    if i < len(ICONS):
        a, b, c = ICONS[i]
        spr.text(f"BA{r}", a, st=s_v); spr.text(f"BB{r}", b, st=s_v); spr.text(f"BC{r}", c, st=s_v)
    else:
        for c in ("BA", "BB", "BC"): spr.put(f"{c}{r}", f'<c r="{c}{r}" s="{s_in}"/>')
spr.text("BA32", "Свой значок: впишите название в BA и точное имя блока в BB (код оставьте пустым) — блок возьмётся из библиотеки DWG.", st=s_v)
# DV: AV — список значков, AY — числа
spr.s = re.sub(r'<dataValidation [^>]*sqref="AV4:AV72">.*?</dataValidation>',
               '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="AV4:AV72"><formula1>$BA$4:$BA$30</formula1></dataValidation>', spr.s, flags=re.S)
def colset(sh, spec):
    m = re.search(r"<cols>(.*?)</cols>", sh.s, flags=re.S)
    cols = re.findall(r"<col [^>]*/>", m.group(1))
    keep = [c for c in cols if not any(X.col2n(a) <= int(re.search(r'max="(\d+)"', c).group(1)) and int(re.search(r'min="(\d+)"', c).group(1)) <= X.col2n(a) for a in spec)]
    new = keep + [f'<col min="{X.col2n(a)}" max="{X.col2n(a)}" width="{w}" customWidth="1"/>' for a, w in spec.items()]
    new.sort(key=lambda c: int(re.search(r'min="(\d+)"', c).group(1)))
    sh.s = sh.s[:m.start()] + "<cols>" + "".join(new) + "</cols>" + sh.s[m.end():]
colset(spr, {"AV": 26, "AY": 12, "AZ": 3, "BA": 28, "BB": 30, "BC": 14})
spr.save()


def bulk(sh_, cells):
    """cells: {ref: cellxml} — вставить/заменить за один проход по строкам"""
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
        return attrs + ">" + "".join(cur[c] for c in sorted(cur, key=X.col2n)) + "</row>"
    sh_.s = re.sub(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', fix, sh_.s, flags=re.S)
    assert not byrow, ("rows missing", sorted(byrow)[:5])

def fx(ref, f, s): return f'<c r="{ref}" s="{s}"><f>{esc(f)}</f></c>'

# ================================================================= 4. ОДНОЛИНЕЙКА: BW 3L, BX проходные, BY УГО(рус) → BK код/блок
sh = X.Sheet(X.sheet_path(WD, "Однолинейка"))
s_hid = sh.style("BU3"); s_hd = sh.style("BU2")
sh.text("BW2", "3L клемм", st=s_hd); sh.text("BX2", "доп. проходных", st=s_hd); sh.text("BY2", "УГО (рус)", st=s_hd)
look = lambda col, r: (f'MAX(IFERROR(--VLOOKUP(R{r},УГО_спр,{col},FALSE),0),IFERROR(--VLOOKUP(AL{r},УГО_спр,{col},FALSE),0),'
                       f'IFERROR(--VLOOKUP(LEFT(AL{r},2),УГО_спр,{col},FALSE),0))')
cores = lambda r: f'IFERROR(--LEFT(SUBSTITUTE(BQ{r},"x","х"),FIND("х",SUBSTITUTE(BQ{r},"x","х"))-1),3)'
cells = {}
for r in range(3, 1001):
    cells[f"BW{r}"] = fx(f"BW{r}", f'IF(AY{r}<>1,0,IF({look(5, r)}>0,{look(5, r)},IF({cores(r)}>3,Каталог!$S$9,1)))', s_hid)
    cells[f"BX{r}"] = fx(f"BX{r}", f'IF(AY{r}<>1,0,MAX(0,{cores(r)}-3*BW{r}))', s_hid)
    cells[f"BY{r}"] = fx(f"BY{r}", (f'IF(R{r}="","",IFERROR(VLOOKUP(R{r},УГО_спр,2,FALSE),IFERROR(VLOOKUP(AL{r},УГО_спр,2,FALSE),'
                          f'IFERROR(VLOOKUP(LEFT(AL{r},2),УГО_спр,2,FALSE),""))))'), s_hid)
    cells[f"BK{r}"] = fx(f"BK{r}", (f'IF(R{r}="","",IF(OR(BY{r}="",BY{r}=0),"BOX",IFERROR(IF(VLOOKUP(BY{r},УГО_блоки,3,FALSE)&""<>"",'
                          f'VLOOKUP(BY{r},УГО_блоки,3,FALSE),VLOOKUP(BY{r},УГО_блоки,2,FALSE)),BY{r})))'), sh.style(f"BK{r}") or s_hid)
bulk(sh, cells)
m = re.search(r'<col min="73" max="74"[^>]*/>', sh.s)
sh.s = sh.s.replace(m.group(0), m.group(0).replace('max="74"', 'max="77"'), 1)
sh.save()

# ================================================================= 5. СПЕЦИФИКАЦИЯ: слева собирается сама, подсчёт справа скрыт
sp = X.Sheet(X.sheet_path(WD, "Спецификация"))
# убрать столбцы K (было вручную), P:S (сверка) и их правила; заметки L64:L66
sp.s = re.sub(r'<c r="(?:K|P|Q|R|S)\d+"[^>]*?(?:/>|>.*?</c>)', "", sp.s, flags=re.S)
sp.s = re.sub(r'<conditionalFormatting sqref="(?:K4|S4)[^"]*">.*?</conditionalFormatting>', "", sp.s, flags=re.S)
for r in (65, 66):
    if sp.get(f"L{r}"): sp.clear(f"L{r}")
st_col = {c: sp.style(f"{c}20") for c in "BCDEFGHIJ"}
X.load_sst(WD)
def ensure_row(sh_, r):
    if re.search(r'<row r="%d"[ >/]' % r, sh_.s): return
    rows = [(int(mm.group(1)), mm.start()) for mm in re.finditer(r'<row r="(\d+)"', sh_.s)]
    pos = next((s_ for n_, s_ in rows if n_ > r), None)
    if pos is None: pos = sh_.s.find("</sheetData>")
    sh_.s = sh_.s[:pos] + f'<row r="{r}"/>' + sh_.s[pos:]
LASTSP = 160
idx = lambda r, col: f'INDEX(Каталог!${col}$4:${col}$500,$A{r})'
for r in range(4, LASTSP + 1):
    ensure_row(sp, r)
    sp.formula(f"A{r}", f'IFERROR(MATCH(ROW()-3,Каталог!$AF$4:$AF$500,0),"")', st=st_col["B"])
    sp.formula(f"B{r}", f'IF($A{r}="","",$B$3&"."&(ROW()-3))', st=st_col["B"])
    sp.formula(f"C{r}", f'IF($A{r}="","",{idx(r, "G")}&"")', st=st_col["C"])
    sp.put(f"D{r}", f'<c r="D{r}" s="{st_col["D"]}"/>')
    sp.formula(f"E{r}", f'IF($A{r}="","",{idx(r, "H")}&"")', st=st_col["E"])
    sp.formula(f"F{r}", f'IF($A{r}="","",{idx(r, "B")}&"")', st=st_col["F"])
    sp.formula(f"G{r}", f'IF($A{r}="","",{idx(r, "M")}&"")', st=st_col["G"])
    sp.formula(f"H{r}", f'IF($A{r}="","",{idx(r, "P")})', st=st_col["H"])
    for c in "IJ": sp.put(f"{c}{r}", f'<c r="{c}{r}" s="{st_col[c]}"/>')
# подсчёт клемм по кабелю: с учётом клемм на линию
for r in range(51, 61):
    sp.formula(f"N{r}", f'IF(L{r}="","",SUMIF(Однолинейка!$BQ$3:$BQ$1000,L{r},Однолинейка!$BW$3:$BW$1000))')
    sp.formula(f"O{r}", f'IF(L{r}="","",SUMIF(Однолинейка!$BQ$3:$BQ$1000,L{r},Однолинейка!$BX$3:$BX$1000))')
sp.formula("M62", "SUM(Однолинейка!$BW$3:$BW$1000)"); sp.formula("M63", "SUM(Однолинейка!$BX$3:$BX$1000)")
# предупреждение над таблицей
ensure_row(sp, 1)
sp.formula("C1", ('IF(Каталог!$R$28="","Спецификация собирается сама из листа «Каталог» (столбец P). Менять количество — в Каталоге, столбец O.",'
                  '"ВНИМАНИЕ: в однолинейке есть позиции, которых нет в Каталоге: "&Каталог!$R$28&IF(Каталог!$R$29="","",", "&Каталог!$R$29)'
                  '&IF(Каталог!$R$30="","",", "&Каталог!$R$30)&" — добавьте их в Каталог.")'), st=st_col["C"])
p = max([int(x) for x in re.findall(r'priority="(\d+)"', sp.s)] or [0])
cf = (f'<conditionalFormatting sqref="C1"><cfRule type="expression" dxfId="{D_RED}" priority="{p+1}">'
      f'<formula>Каталог!$R$28&lt;&gt;""</formula></cfRule></conditionalFormatting>')
i = sp.s.find("<dataValidations"); i = i if i >= 0 else sp.s.find("<pageMargins")
sp.s = sp.s[:i] + cf + sp.s[i:]
# столбцы: A скрыт, K:S свернуты
m = re.search(r"<cols>(.*?)</cols>", sp.s, flags=re.S)
cols = [c for c in re.findall(r"<col [^>]*/>", m.group(1)) if int(re.search(r'min="(\d+)"', c).group(1)) <= 10]
cols = [re.sub(r'max="(\d+)"', lambda mm: f'max="{min(int(mm.group(1)), 10)}"', c) for c in cols]
cols.insert(0, '<col min="1" max="1" width="4" hidden="1" customWidth="1"/>')
cols.append('<col min="11" max="11" width="3" customWidth="1"/>')
cols.append('<col min="12" max="12" width="40" hidden="1" outlineLevel="1" customWidth="1"/>')
cols.append('<col min="13" max="19" width="11" hidden="1" outlineLevel="1" customWidth="1"/>')
sp.s = sp.s[:m.start()] + "<cols>" + "".join(cols) + "</cols>" + sp.s[m.end():]
sp.s = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:S{LASTSP}"/>', sp.s)
sp.save()

# ================================================================= 6. старая ручная спецификация — скрытый лист
old_xml = open(old_sp.path, encoding="utf-8").read()
old_xml = re.sub(r'<c r="(?:[K-Z]|A[A-Z])\d+"[^>]*?(?:/>|>.*?</c>)', "", old_xml, flags=re.S)
old_xml = re.sub(r'<conditionalFormatting.*?</conditionalFormatting>', "", old_xml, flags=re.S)
# общие строки старого файла → inline
X.load_sst(WO)
OLD_SST = X.SST if hasattr(X, "SST") else None
def inl(mm):
    cell = mm.group(0)
    v = re.search(r"<v>(\d+)</v>", cell)
    txt = X.SST[int(v.group(1))]
    cell = re.sub(r'\st="s"', ' t="inlineStr"', cell)
    return re.sub(r"<v>\d+</v>", f'<is><t xml:space="preserve">{esc(txt)}</t></is>', cell)
old_xml = re.sub(r'<c [^>]*t="s"[^>]*>.*?</c>', inl, old_xml, flags=re.S)
old_xml = re.sub(r"<f>.*?</f>", "", old_xml, flags=re.S)
old_xml = re.sub(r'<c r="([A-Z]+\d+)"([^>]*)></c>', r'<c r="\1"\2/>', old_xml)
X.load_sst(WD)
n = 1
while os.path.exists(f"{WD}/xl/worksheets/sheet{n}.xml"): n += 1
fn = f"sheet{n}.xml"
# без связей на печать старого листа
old_xml = re.sub(r"<pageSetup[^>]*/>", "", old_xml)
old_xml = re.sub(r'\sxr:uid="[^"]*"', "", old_xml)
old_xml = re.sub(r'\stabSelected="1"', "", old_xml)
old_xml = re.sub(r"<(drawing|legacyDrawing|legacyDrawingHF|picture)\b[^>]*/>", "", old_xml)
assert "r:id" not in old_xml, re.findall(r"<[^>]*r:id[^>]*>", old_xml)[:3]
open(f"{WD}/xl/worksheets/{fn}", "w", encoding="utf-8").write(old_xml)
rp = WD + "/xl/_rels/workbook.xml.rels"; rels = open(rp, encoding="utf-8").read()
ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]; rid = f"rId{max(ids) + 1}"
rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{fn}"/></Relationships>')
open(rp, "w", encoding="utf-8").write(rels)
ctp = WD + "/[Content_Types].xml"; ct = open(ctp, encoding="utf-8").read()
open(ctp, "w", encoding="utf-8").write(ct.replace("</Types>", f'<Override PartName="/xl/worksheets/{fn}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'))
order.append(f"xl/worksheets/{fn}")

# ================================================================= 7. книга: имена, новый лист
wbp = WD + "/xl/workbook.xml"; wb = open(wbp, encoding="utf-8").read()
sids = [int(x) for x in re.findall(r'sheetId="(\d+)"', wb)]
wb = wb.replace("</sheets>", f'<sheet name="Спецификация_до_v3.7" sheetId="{max(sids) + 1}" state="hidden" r:id="{rid}"/></sheets>')
def setname(name, val):
    global wb
    pat = r'(<definedName name="%s"[^>]*>)[^<]*(</definedName>)' % re.escape(name)
    if re.search(pat, wb): wb = re.sub(pat, lambda mm: mm.group(1) + esc(val) + mm.group(2), wb)
    else: wb = wb.replace("</definedNames>", f'<definedName name="{name}">{esc(val)}</definedName></definedNames>')
setname("Список_автоматов", "OFFSET(Каталог!$X$3,0,0,1+MAX(Каталог!$W$4:$W$500),1)")
setname("Список_аппаратов", "OFFSET(Каталог!$Z$4,0,0,MAX(1,MAX(Каталог!$Y$4:$Y$500)),1)")
setname("Список_модулей", "OFFSET(Каталог!$AB$3,0,0,1+MAX(Каталог!$AA$4:$AA$500),1)")
setname("Модели_список", "OFFSET(Каталог!$AB$3,0,0,1+MAX(Каталог!$AA$4:$AA$500),1)")
setname("Список_произв", "OFFSET(Каталог!$AD$3,0,0,1+MAX(Каталог!$AC$4:$AC$500),1)")
setname("УГО_спр", "Справочная!$AU$4:$AY$72")
setname("УГО_блоки", "Справочная!$BA$4:$BC$30")
setname("Столбцов_на_лист", "Справочник_схемы!$B$4")
setname("Мин_автомат_БП", "Справочник_схемы!$B$5")
names = re.findall(r'<sheet name="([^"]+)"', wb); ci = names.index("Каталог")
wb = re.sub(r'(<definedName name="_xlnm._FilterDatabase" localSheetId="%d" hidden="1">)[^<]*' % ci, r"\1Каталог!$A$3:$P$500", wb)
# имена по алфавиту (как пишет Excel)
dn = re.search(r"<definedNames>(.*?)</definedNames>", wb, flags=re.S)
items = re.findall(r"<definedName .*?</definedName>", dn.group(1), flags=re.S)
items.sort(key=lambda s_: (re.search(r'name="([^"]+)"', s_).group(1).lower(), re.search(r'localSheetId="(\d+)"', s_).group(1) if "localSheetId" in s_ else ""))
wb = wb[:dn.start()] + "<definedNames>" + "".join(items) + "</definedNames>" + wb[dn.end():]
open(wbp, "w", encoding="utf-8").write(wb)

# ================================================================= 8. руководство
gd = X.Sheet(X.sheet_path(WD, "Краткое руководство"))
sE = gd.style("E11"); sF = gd.style("F11"); sB = gd.style("B12"); sC = gd.style("C12")
gd.text("E12", "v3.7 (схема)", st=sE)
gd.text("F12", "Спецификация собирается сама из «Каталога» (без пустых строк, по разделам). Понятный «Справочник_схемы». "
               "УГО по-русски + свои блоки. Клеммы 3L на линию по коду (шторы SH = 2). Старая ручная спецификация — скрытый лист «Спецификация_до_v3.7».", st=sF)
RULES = {
    15: "СПЕЦИФИКАЦИЯ собирается сама из листа «Каталог»: N — посчитано по однолинейке, O — впишите вручную (корпус, провод, шины) или запас. Новая позиция — строкой вниз в «Каталоге».",
    17: "Новый код потребителя: строка в «Справочной» (P…Y) + в той же строке УГО (список), 30 мА, мин. автомат, клемм 3L на линию (AU:AY).",
    18: "Свой значок УГО: «Справочная» BA:BB — название + имя блока из 02_Однолинейка.dwg. Лист «В Акад» — старый способ (Data Link), оставлен для коллег.",
}
for r, t in RULES.items():
    gd.text(f"B{r}", str(r - 1), st=sB); gd.text(f"C{r}", t, st=sC)
gd.save()

# ---------------------------------------------------------------- сборка
if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n_ in order:
        p_ = os.path.join(WD, n_)
        if os.path.exists(p_): z.write(p_, n_)
print("ok")
