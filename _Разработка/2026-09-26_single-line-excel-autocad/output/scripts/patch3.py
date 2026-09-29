"""v3.5: Каталог + двухуровневые списки, светофор, защита формул, бирки, артикулы в подсчёте спецификации."""
import sys, os, re, shutil, zipfile, html
SP = "/tmp/claude-0/-home-claude/5e645175-0e52-5350-b586-c0cb5927e030/scratchpad/"
sys.path.insert(0, SP + "ost")
import xlsxlib as X

SRC = SP + "ost/Однолинейка ЩР ЭОМ_v3.4.xlsx"
OUT = SP + "ost/Однолинейка ЩР ЭОМ_v3.5.xlsx"
WD = SP + "ost/w4"
PASSWORD = "sx"
shutil.rmtree(WD, ignore_errors=True)
with zipfile.ZipFile(SRC) as z:
    z.extractall(WD); order = z.namelist()
X.load_sst(WD)
esc = X.esc
O = "Однолинейка!"
FIRST, LAST = 3, 1000

# ================================================================= стили
stp = WD + "/xl/styles.xml"; st = open(stp, encoding="utf-8").read()
def add(tag, xml):
    global st
    m = re.search(r'<%s count="(\d+)"' % tag, st); n = int(m.group(1))
    st = st.replace(m.group(0), '<%s count="%d"' % (tag, n + 1), 1)
    k = st.find("</%s>" % tag); st = st[:k] + xml + st[k:]
    return n
xfs_block = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", st, flags=re.S).group(1)
XFS = re.findall(r"<xf [^>]*?(?:/>|>.*?</xf>)", xfs_block, flags=re.S)
_unlock_cache = {}
def unlocked(sid):
    sid = int(sid or 0)
    if sid in _unlock_cache: return _unlock_cache[sid]
    xf = XFS[sid]
    xf = re.sub(r"<protection[^>]*/>", "", xf)
    if xf.endswith("/>"):
        xf = xf[:-2] + ' applyProtection="1"><protection locked="0"/></xf>'
    else:
        xf = xf.replace("</xf>", '<protection locked="0"/></xf>')
        if "applyProtection" not in xf: xf = xf.replace("<xf ", '<xf applyProtection="1" ', 1)
    n = add("cellXfs", xf); XFS.append(xf)
    _unlock_cache[sid] = n
    return n
D_RED = add("dxfs", '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>')
D_YEL = add("dxfs", '<dxf><font><color rgb="FF7F6000"/></font><fill><patternFill><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>')
D_GRN = add("dxfs", '<dxf><font><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>')

# ================================================================= лист «Каталог» (новый)
ref = X.Sheet(X.sheet_path(WD, "Справочник_схемы"))
S_HEAD = ref.style("A3") or "0"; S_VAL = ref.style("A4") or "0"; S_TITLE = ref.style("A1") or S_HEAD
S_VAL_U = unlocked(S_VAL)

CAT = [  # Категория, Производитель, Модель(в таблицу), Префикс, Каналов, Тип каналов, Наименование, Артикул, Полюса, Номинал, 30мА, Ширина, Ед.
 ("Автомат", "Chint", "NB1-63 1P C10", "", "", "", "Автоматический выключатель NB1-63 1P 10А 6кА х-ка C (DB)", "179614", "1P", 10, "нет", 1, "шт."),
 ("Автомат", "Chint", "NB1-63 1P C16", "", "", "", "Автоматический выключатель NB1-63 1P 16А 6кА х-ка C (DB)", "179616", "1P", 16, "нет", 1, "шт."),
 ("Автомат", "Chint", "NB1-63 3P C20", "", "", "", "Автоматический выключатель NB1-63 3P 20А 6кА х-ка C (DB)", "179702", "3P", 20, "нет", 3, "шт."),
 ("Автомат", "Chint", "NB1-63 4P C50", "", "", "", "Автоматический выключатель NB1-63 4P 50А 6кА х-ка C (DB)", "179750", "4P", 50, "нет", 4, "шт."),
 ("Автомат", "Chint", "NXB-125 3P C80", "", "", "", "Автоматический выключатель NXB-125 3P C80 10кА х-ка C", "816139", "3P", 80, "нет", 4.5, "шт."),
 ("Дифавтомат", "Chint", "NB1L 1P+N C16 30мА", "", "", "", "Дифавтомат (АВДТ) NB1L 1P+N C16 30мА тип A 6кА", "203019", "2P", 16, "да", 2, "шт."),
 ("Дифавтомат", "Chint", "NB1L 1P+N C20 30мА", "", "", "", "Дифавтомат (АВДТ) NB1L 1P+N C20 30мА тип A 6кА", "203296", "2P", 20, "да", 2, "шт."),
 ("Дифавтомат", "Chint", "NB1L 1P+N C32 30мА", "", "", "", "Дифавтомат (АВДТ) NB1L 1P+N C32 30мА тип A 6кА", "982114", "2P", 32, "да", 2, "шт."),
 ("Аппарат", "Россия", "МК-5-1", "", "", "", "Модуль защиты контактов MK-5-1", "EA06.002.001", "", "", "", 1, "шт."),
 ("Аппарат", "Chint", "КМ NCH8-63/40", "", "", "", "Контактор модульный NCH8-63/40 63А 4НО AC220/230В", "256101", "", "", "", 3, "шт."),
 ("Аппарат", "Chint", "КМ NCH8-40/40", "", "", "", "Контактор модульный NCH8-40/40 40А 4НО AC220/230В", "256099", "", "", "", 3, "шт."),
 ("Аппарат", "Chint", "NJVA1-100", "", "", "", "Реле контроля напряжения и тока NJVA1-100 3P+N 100А", "508572", "", "", "", "", "шт."),
 ("Аппарат", "Chint", "NZK1-32 2P", "", "", "", "Модульный переключатель NZK1-32 2P 32А 3 положения", "643001", "", "", "", "", "шт."),
 ("Аппарат", "Arlight", "БП 24V 480W", "", "", "", "Блок питания ARV-DRP480-PFC-24 (24V, 20A, 480W)", "040231", "", "", "", "", "шт."),
 ("Аппарат", "Arlight", "БП 24V 75W", "", "", "", "Блок питания ARV-DRP75-24 (24V, 3.15A, 75W)", "040232", "", "", "", "", "шт."),
 ("Аппарат", "Arlight", "БП 12V 135W", "", "", "", "Блок питания ARV-DR135-12 (12V, 11.3A, 135W)", "035699", "", "", "", "", "шт."),
 ("Аппарат", "Arlight", "БП 12V 15W", "", "", "", "Блок питания ARV-DR15-12 (12V, 1.25A, 15W)", "034670", "", "", "", "", "шт."),
 ("Аппарат", "Zennio", "ZPSU640", "", "", "", "Универсальный источник питания KNX", "ZPSU640", "", "", "", "", "шт."),
 ("Модуль", "Zennio", "ZIOMB24V2", "R", 24, "реле", "Многофункциональный актуатор KNX", "ZIOMB24V2", "", "", "", "", "шт."),
 ("Модуль", "Zennio", "ZIO-MB24", "R", 24, "реле", "Многофункциональный актуатор KNX", "ZIO-MB24", "", "", "", "", "шт."),
 ("Модуль", "Zennio", "ZIOMB16V3", "R", 16, "реле", "Многофункциональный актуатор KNX", "ZIOMB16V3", "", "", "", "", "шт."),
 ("Модуль", "Zennio", "ZIOMB8V3", "R", 8, "реле", "Многофункциональный актуатор KNX", "ZIOMB8V3", "", "", "", "", "шт."),
 ("Модуль", "Zennio", "ZCL-8HT230", "R", 8, "клапаны 230В", "Актуатор клапанов отопления 230В", "ZCL-8HT230", "", "", "", "", "шт."),
 ("Модуль", "Zennio", "ZDINDX4", "Di", 4, "диммер", "Универсальный 4-канальный диммер 220V", "ZDINDX4", "", "", "", "", "шт."),
 ("Модуль", "Zennio", "ZDID64X2", "DA", 2, "DALI", "Интерфейс KNX-DALI", "ZDID64X2", "", "", "", "", "шт."),
 ("Модуль", "Zennio", "ZIO-RQUAD8", "IN", 8, "входы", "Модуль входов с 8 аналогово-цифровыми входами", "ZIO-RQUAD8", "", "", "", "", "шт."),
 ("Модуль", "Arlight", "SMART-DALI-104", "Di", 4, "диммер DALI", "Диммер SMART-DALI-104-62-ADDR-DIM-DT6-DIN White (12-24V, 4x5A)", "028422", "", "", "", "", "шт."),
 ("Модуль", "—", "уточнить-R", "R", 24, "реле", "Модель уточнить", "", "", "", "", "", "шт."),
 ("Модуль", "—", "уточнить-Z", "Z", 8, "датчики", "Модель уточнить", "", "", "", "", "", "шт."),
 ("Клемма", "DKC", "VPR-2.5-3L-GY", "", "", "", "Клемма трёхуровневая push-in, 6 точек, 2.5 кв.мм", "VPR-2.5-3L-GY", "", "", "", "", "шт."),
 ("Клемма", "DKC", "VPR-2.5-QUATTRO-GY", "", "", "", "Клемма проходная push-in, 4 точки, 2.5 кв.мм, серая", "VPR-2.5-QUATTRO-GY", "", "", "", "", "шт."),
 ("Клемма", "DKC", "D-VPR-2.5-3L", "", "", "", "Изолятор торцевой для клеммы VPR-2.5-3L", "D-VPR-2.5-3L", "", "", "", "", "шт."),
 ("Клемма", "DKC", "VPR-4-GY", "", "", "", "Клемма проходная push-in, 2 точки, 4 кв.мм, серая", "VPR-4-GY", "", "", "", "", "шт."),
 ("Клемма", "DKC", "VPR-6-GY", "", "", "", "Клемма проходная push-in, 2 точки, 6 кв.мм, серая", "VPR-6-GY", "", "", "", "", "шт."),
 ("Клемма", "DKC", "VPR-16-GY", "", "", "", "Клемма проходная push-in, 2 точки, 16 кв.мм, серая", "VPR-16-GY", "", "", "", "", "шт."),
]
HEAD = ["Категория", "Производитель", "Модель (как в таблице)", "Префикс модуля", "Каналов", "Тип каналов",
        "Наименование для спецификации", "Артикул", "Полюса", "Номинал, А", "30 мА", "Ширина, мод. DIN", "Ед."]
COLS = "ABCDEFGHIJKLM"
CL = 500
rows = {}
def cell_t(ref, t, s): return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{esc(str(t))}</t></is></c>'
def cell_n(ref, v, s): return f'<c r="{ref}" s="{s}"><v>{v}</v></c>'
def cell_f(ref, f, s): return f'<c r="{ref}" s="{s}"><f>{esc(f)}</f></c>'
def cell_e(ref, s): return f'<c r="{ref}" s="{s}"/>'
def put(r, xml): rows.setdefault(r, []).append(xml)
put(1, cell_t("A1", "КАТАЛОГ ОБОРУДОВАНИЯ — добавляйте строки вниз (до 500). Отсюда берутся выпадающие списки, артикулы спецификации и ёмкость модулей.", S_TITLE))
for i, h in enumerate(HEAD): put(3, cell_t(f"{COLS[i]}3", h, S_HEAD))
for k in range(4, CL + 1):
    rec = CAT[k - 4] if k - 4 < len(CAT) else None
    for i, c in enumerate(COLS):
        ref_ = f"{c}{k}"
        if rec is None or rec[i] in ("", None): put(k, cell_e(ref_, S_VAL_U))
        elif isinstance(rec[i], (int, float)): put(k, cell_n(ref_, rec[i], S_VAL_U))
        else: put(k, cell_t(ref_, rec[i], S_VAL_U))
# настройки O:P
SETT = [("Производитель автоматов", "Chint"), ("Производитель аппаратов", "(все)"), ("Производитель модулей", "(все)"),
        ("Клемма трёхуровневая (модель)", "VPR-2.5-3L-GY"), ("Клемма проходная доп. (модель)", "VPR-2.5-QUATTRO-GY")]
put(3, cell_t("O3", "НАСТРОЙКИ СПИСКОВ (выберите)", S_HEAD)); put(3, cell_t("P3", "Значение", S_HEAD))
for i, (a, b) in enumerate(SETT):
    put(4 + i, cell_t(f"O{4+i}", a, S_VAL)); put(4 + i, cell_t(f"P{4+i}", b, S_VAL_U))
# служебные столбцы R:AA (скрыты)
HS = [("R", "автомат (обозн.)"), ("S", "ключ автомата"), ("T", "№ в списке автоматов"), ("U", "список автоматов"),
      ("V", "№ аппарата"), ("W", "список аппаратов"), ("X", "№ модуля"), ("Y", "список модулей"), ("Z", "№ произв."), ("AA", "список произв.")]
for c, h in HS: put(2 if c in ("U", "Y", "AA") else 3, cell_t(f"{c}{2 if c in ('U', 'Y', 'AA') else 3}", h, S_HEAD))
put(3, cell_t("U3", "авто", S_VAL)); put(3, cell_t("Y3", "-", S_VAL)); put(3, cell_t("AA3", "(все)", S_VAL))
for k in range(4, CL + 1):
    p = k - 1
    F = {
     "R": f'IF(OR(A{k}="Автомат",A{k}="Дифавтомат"),I{k}&" С"&J{k}&"А"&IF(K{k}="да"," 30мА",""),"")',
     "S": f'IF(R{k}="","",B{k}&"|"&R{k})',
     "T": f'IF(AND(R{k}<>"",OR($P$4="(все)",B{k}=$P$4),COUNTIFS($R$4:R{k},R{k},$B$4:B{k},B{k})=1),MAX($T$3:T{p})+1,"")',
     "U": f'IFERROR(INDEX($R$4:$R${CL},MATCH(ROW()-3,$T$4:$T${CL},0)),"")',
     "V": f'IF(AND(A{k}="Аппарат",C{k}<>"",OR($P$5="(все)",B{k}=$P$5)),MAX($V$3:V{p})+1,"")',
     "W": f'IFERROR(INDEX($C$4:$C${CL},MATCH(ROW()-3,$V$4:$V${CL},0)),"")',
     "X": f'IF(AND(A{k}="Модуль",C{k}<>"",OR($P$6="(все)",B{k}=$P$6)),MAX($X$3:X{p})+1,"")',
     "Y": f'IFERROR(INDEX($C$4:$C${CL},MATCH(ROW()-3,$X$4:$X${CL},0)),"")',
     "Z": f'IF(OR(B{k}="",COUNTIF($B$4:B{k},B{k})>1),"",MAX($Z$3:Z{p})+1)',
     "AA": f'IFERROR(INDEX($B$4:$B${CL},MATCH(ROW()-3,$Z$4:$Z${CL},0)),"")',
    }
    for c, f in F.items(): put(k, cell_f(f"{c}{k}", f, S_VAL))
def colkey(ref):
    m = re.match(r"([A-Z]+)", ref); return X.col2n(m.group(1))
sd = []
for r in sorted(rows):
    cells = sorted(rows[r], key=lambda x: colkey(re.search(r'r="([A-Z]+)', x).group(1)))
    sd.append(f'<row r="{r}">' + "".join(cells) + "</row>")
widths = {"A": 12, "B": 13, "C": 24, "D": 9, "E": 8, "F": 12, "G": 60, "H": 18, "I": 8, "J": 9, "K": 7, "L": 9, "M": 6, "N": 3, "O": 30, "P": 22}
colxml = "".join(f'<col min="{X.col2n(c)}" max="{X.col2n(c)}" width="{w}" customWidth="1"/>' for c, w in widths.items())
colxml += f'<col min="{X.col2n("R")}" max="{X.col2n("AA")}" width="10" hidden="1" customWidth="1"/>'
dv = ('<dataValidations count="4">'
      '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="A4:A500"><formula1>"Автомат,Дифавтомат,Аппарат,Модуль,Клемма,Прочее"</formula1></dataValidation>'
      '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="K4:K500"><formula1>"да,нет"</formula1></dataValidation>'
      '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="I4:I500"><formula1>"1P,2P,3P,4P"</formula1></dataValidation>'
      '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="P4:P6"><formula1>Список_произв</formula1></dataValidation>'
      '</dataValidations>')
cat_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           '<sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
           '<sheetFormatPr defaultRowHeight="15"/>'
           f'<cols>{colxml}</cols><sheetData>{"".join(sd)}</sheetData>'
           f'<autoFilter ref="A3:M{CL}"/>{dv}'
           '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>')

def add_sheet(name, xml, after):
    n = 1
    while os.path.exists(f"{WD}/xl/worksheets/sheet{n}.xml"): n += 1
    fn = f"sheet{n}.xml"
    open(f"{WD}/xl/worksheets/{fn}", "w", encoding="utf-8").write(xml)
    rp = WD + "/xl/_rels/workbook.xml.rels"; rels = open(rp, encoding="utf-8").read()
    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]; rid = f"rId{max(ids) + 1}"
    rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{fn}"/></Relationships>')
    open(rp, "w", encoding="utf-8").write(rels)
    ctp = WD + "/[Content_Types].xml"; ct = open(ctp, encoding="utf-8").read()
    ct = ct.replace("</Types>", f'<Override PartName="/xl/worksheets/{fn}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    open(ctp, "w", encoding="utf-8").write(ct)
    wbp = WD + "/xl/workbook.xml"; wb = open(wbp, encoding="utf-8").read()
    sids = [int(x) for x in re.findall(r'sheetId="(\d+)"', wb)]
    m = re.search(r'<sheet name="%s"[^>]*/>' % re.escape(after), wb)
    wb = wb[:m.end()] + f'<sheet name="{name}" sheetId="{max(sids) + 1}" r:id="{rid}"/>' + wb[m.end():]
    open(wbp, "w", encoding="utf-8").write(wb)
    order.append(f"xl/worksheets/{fn}")
    return fn
add_sheet("Каталог", cat_xml, "Справочник_схемы")

# ================================================================= имена
wbp = WD + "/xl/workbook.xml"; wb = open(wbp, encoding="utf-8").read()
def setname(name, val):
    global wb
    pat = r'(<definedName name="%s"[^>]*>)[^<]*(</definedName>)' % re.escape(name)
    if re.search(pat, wb): wb = re.sub(pat, lambda m: m.group(1) + esc(val) + m.group(2), wb)
    else: wb = wb.replace("</definedNames>", f'<definedName name="{name}">{esc(val)}</definedName></definedNames>')
setname("Модули_спр", "Каталог!$C$4:$F$500")
setname("Список_автоматов", "OFFSET(Каталог!$U$3,0,0,1+MAX(Каталог!$T$4:$T$500),1)")
setname("Список_аппаратов", "OFFSET(Каталог!$W$4,0,0,MAX(1,MAX(Каталог!$V$4:$V$500)),1)")
setname("Список_модулей", "OFFSET(Каталог!$Y$3,0,0,1+MAX(Каталог!$X$4:$X$500),1)")
setname("Список_произв", "OFFSET(Каталог!$AA$3,0,0,1+MAX(Каталог!$Z$4:$Z$500),1)")
setname("Модели_список", "OFFSET(Каталог!$Y$3,0,0,1+MAX(Каталог!$X$4:$X$500),1)")
if "fullCalcOnLoad" not in wb:
    wb = re.sub(r"<calcPr([^>]*)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb)
open(wbp, "w", encoding="utf-8").write(wb)

# ================================================================= Справочник: старая таблица моделей -> примечание; итоги светофора
for r in range(3, 41):
    for c in "ABCDE":
        if not ref.is_empty(f"{c}{r}"): ref.clear(f"{c}{r}")
ref.text("A3", "Модели модулей теперь на листе «Каталог»", st=S_HEAD)
ref.text("J15", "Строк с ошибками (красные)", st=ref.style("J14")); ref.formula("K15", f'COUNTIF({O}$BU$3:$BU$1000,2)', st=ref.style("K14"))
ref.text("J16", "Строк с предупреждениями (жёлтые)", st=ref.style("J14")); ref.formula("K16", f'COUNTIF({O}$BU$3:$BU$1000,1)', st=ref.style("K14"))
ref.save()

# ================================================================= Однолинейка: светофор, бирки-индекс, списки, защита
sh = X.Sheet(X.sheet_path(WD, "Однолинейка"))
st_h = sh.style("BA21"); st_hd = sh.style("BA2"); st_ae = sh.style("AD21") or st_h
sh.text("AE2", "Статус строки\n(что проверить)", st=sh.style("AD2"))
sh.text("BU2", "уровень", st=st_hd); sh.text("BV2", "№ линии", st=st_hd)
AE = ('IF(AND(R{r}="",G{r}=""),"",_xlfn.TEXTJOIN("; ",TRUE,'
      'IF(AND(R{r}<>"",V{r}=""),"нет фазы",""),'
      'IF(AND(R{r}<>"",N(T{r})=0),"нет мощности",""),'
      'IF(AND(P{r}<>"",COUNTIFS($A$3:$A$1000,A{r},$P$3:$P$1000,P{r})>1),"дубль клеммы "&P{r},""),'
      'IF(AND(W{r}<>"",W{r}<>"ok"),W{r},""),'
      'IF(X{r}<>"",X{r},""),'
      'IF(AND(ISNUMBER(SEARCH("(5х",Q{r})),N(AK{r})=1),"5 жил на 1 фазе",""),'
      'IF(AND(AL{r}="R",E{r}<>"",N(AS{r})=0),"розетка без 30 мА",""),'
      'IF(IFERROR(_xlfn.NUMBERVALUE(MID(Q{r},FIND("-",Q{r},FIND("L=",Q{r}))+1,FIND("%",Q{r})-FIND("-",Q{r},FIND("L=",Q{r}))-1),","),0)>4,"потери > 4 %","")))')
BU = ('IF(AE{r}="",0,IF(OR(ISNUMBER(SEARCH("дубль",AE{r})),ISNUMBER(SEARCH("ПЕРЕГРУЗ",AE{r})),ISNUMBER(SEARCH("ЗАЩИЩ",AE{r})),'
      'ISNUMBER(SEARCH("КАНАЛ",AE{r})),ISNUMBER(SEARCH("без 30",AE{r}))),2,1))')
BV = 'IF(R{r}="","",COUNTIF($R$3:R{r},"?*"))'
for r in range(FIRST, LAST + 1):
    sh.formula(f"AE{r}", AE.format(r=r), st=st_ae)
    sh.formula(f"BU{r}", BU.format(r=r), st=st_h)
    sh.formula(f"BV{r}", BV.format(r=r), st=st_h)
# светофор: подсветка № места и статуса
cf = (f'<conditionalFormatting sqref="B3:B1000 AE3:AE1000">'
      f'<cfRule type="expression" dxfId="{D_RED}" priority="1"><formula>$BU3=2</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{D_YEL}" priority="2"><formula>$BU3=1</formula></cfRule>'
      f'<cfRule type="expression" dxfId="{D_GRN}" priority="3"><formula>AND($R3&lt;&gt;"",$BU3=0)</formula></cfRule>'
      f'</conditionalFormatting>')
# сдвинуть приоритеты существующих правил, чтобы наши были первыми
sh.s = re.sub(r'priority="(\d+)"', lambda m: f'priority="{int(m.group(1)) + 10}"', sh.s)
k = sh.s.find("<conditionalFormatting"); sh.s = sh.s[:k] + cf + sh.s[k:]
# выпадающие списки
newdv = ('<dataValidations count="6">'
         '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="E3:E1000"><formula1>Список_автоматов</formula1></dataValidation>'
         '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="F3:F1000"><formula1>"да,нет"</formula1></dataValidation>'
         '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="I3:K1000"><formula1>Список_аппаратов</formula1></dataValidation>'
         '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="L3:L1000"><formula1>Список_модулей</formula1></dataValidation>'
         '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="N3:N1000"><formula1>"+,=,р,б/к"</formula1></dataValidation>'
         '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="Y3:Y1000"><formula1>"вручную"</formula1></dataValidation>'
         '</dataValidations>')
sh.s = re.sub(r"<dataValidations.*?</dataValidations>", newdv, sh.s, flags=re.S)
# разблокировать столбцы ввода (D E F I J K L N Y), строки 3..1000
INP = ["D", "E", "F", "I", "J", "K", "L", "N", "P", "Y"]
seen = set()
def unl(m):
    col, r, attrs = m.group(1), int(m.group(2)), m.group(3)
    if col not in INP or not (FIRST <= r <= LAST): return m.group(0)
    seen.add(f"{col}{r}")
    sm = re.search(r'\ss="(\d+)"', attrs)
    new = unlocked(sm.group(1) if sm else 0)
    attrs = re.sub(r'\ss="\d+"', "", attrs)
    return f'<c r="{col}{r}" s="{new}"{attrs}'
sh.s = re.sub(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?=/>|>)', unl, sh.s)
for col in INP:
    for r in range(FIRST, LAST + 1):
        if f"{col}{r}" not in seen:
            sh.put(f"{col}{r}", f'<c r="{col}{r}" s="{unlocked(0)}"/>')
# защита листа (пароль — старый хэш Excel)
def xl_hash(pw):
    h = 0
    for ch in reversed(pw):
        h = ((h >> 14) & 0x01) | ((h << 1) & 0x7FFF); h ^= ord(ch)
    h = ((h >> 14) & 0x01) | ((h << 1) & 0x7FFF); h ^= len(pw); h ^= 0xCE4B
    return format(h, "X")
prot = (f'<sheetProtection password="{xl_hash(PASSWORD)}" sheet="1" objects="1" scenarios="1" '
        f'formatCells="0" formatColumns="0" formatRows="0" autoFilter="0" sort="0"/>')
sh.s = re.sub(r"<sheetProtection[^>]*/>", "", sh.s)
sh.s = sh.s.replace("</sheetData>", "</sheetData>" + prot, 1)
# скрыть BU:BV
a, b = X.col2n("BU"), X.col2n("BV")
sh.s = sh.s.replace("</cols>", f'<col min="{a}" max="{b}" width="8" hidden="1" outlineLevel="1" customWidth="1"/></cols>', 1)
# ширина AE
if not re.search(r'<col min="31" max="31"', sh.s):
    sh.s = sh.s.replace("</cols>", '<col min="31" max="31" width="34" customWidth="1"/></cols>', 1)
sh.save()

# ================================================================= лист «Бирки»
bk = {}
def bput(r, xml): bk.setdefault(r, []).append(xml)
BH = [("A", "№", 6), ("B", "Щит", 8), ("C", "Бирка автомат/клемма", 22), ("D", "Бирка кабеля", 18), ("E", "Наименование", 44), ("F", "Кабель", 34)]
bput(1, cell_t("A1", "БИРКИ ДЛЯ ПОДКЛЮЧЕНИЯ (считаются сами из листа «Однолинейка»). Печатайте столбцы C и D.", S_TITLE))
for c, h, w in BH: bput(2, cell_t(f"{c}2", h, S_HEAD))
NB = 600
for k in range(1, NB + 1):
    r = k + 2
    rowm = f'MATCH({k},{O}$BV$1:$BV$1000,0)'
    bput(r, cell_f(f"A{r}", f'IF(ISNUMBER({rowm}),{k},"")', S_VAL))
    bput(r, cell_f(f"B{r}", f'IFERROR(INDEX({O}$A$1:$A$1000,{rowm}),"")', S_VAL))
    bput(r, cell_f(f"C{r}", f'IFERROR(INDEX({O}$G$1:$G$1000,{rowm})&IF(INDEX({O}$P$1:$P$1000,{rowm})="","","/"&INDEX({O}$P$1:$P$1000,{rowm})),"")', S_VAL))
    bput(r, cell_f(f"D{r}", f'IFERROR(INDEX({O}$R$1:$R$1000,{rowm}),"")', S_VAL))
    bput(r, cell_f(f"E{r}", f'IFERROR(INDEX({O}$S$1:$S$1000,{rowm}),"")', S_VAL))
    bput(r, cell_f(f"F{r}", f'IFERROR(INDEX({O}$Q$1:$Q$1000,{rowm}),"")', S_VAL))
sd = [f'<row r="{r}">' + "".join(bk[r]) + "</row>" for r in sorted(bk)]
colxml = "".join(f'<col min="{X.col2n(c)}" max="{X.col2n(c)}" width="{w}" customWidth="1"/>' for c, h, w in BH)
bir_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           '<sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
           '<sheetFormatPr defaultRowHeight="15"/>'
           f'<cols>{colxml}</cols><sheetData>{"".join(sd)}</sheetData>'
           '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>')
add_sheet("Бирки", bir_xml, "Каталог")

# ================================================================= Спецификация: артикулы к подсчёту
sp = X.Sheet(X.sheet_path(WD, "Спецификация"))
st_head = sp.style("C2"); st_txt = sp.style("C4"); st_num = sp.style("H4")
for c, t in zip("PQR", ["Наименование по каталогу", "Артикул", "Производитель"]): sp.text(f"{c}2", t, st=st_head)
C = "Каталог!"
def by_model(r, col):
    miss = '"нет в каталоге"' if col == "G" else '""'
    return f'IF(L{r}="","",IFERROR(INDEX({C}${col}$4:${col}$500,MATCH(L{r},{C}$C$4:$C$500,0)),{miss}))'
for r in range(4, 24):   # автоматы: по производителю из настроек + обозначение
    key = f'{C}$P$4&"|"&L{r}'
    sp.formula(f"P{r}", f'IF(L{r}="","",IFERROR(INDEX({C}$G$4:$G$500,MATCH({key},{C}$S$4:$S$500,0)),"нет в каталоге"))', st=st_txt)
    sp.formula(f"Q{r}", f'IF(L{r}="","",IFERROR(INDEX({C}$H$4:$H$500,MATCH({key},{C}$S$4:$S$500,0)),""))', st=st_txt)
    sp.formula(f"R{r}", f'IF(OR(L{r}="",Q{r}=""),"",{C}$P$4)', st=st_txt)
for r in list(range(27, 37)) + list(range(39, 49)):
    sp.formula(f"P{r}", by_model(r, "G"), st=st_txt); sp.formula(f"Q{r}", by_model(r, "H"), st=st_txt); sp.formula(f"R{r}", by_model(r, "B"), st=st_txt)
sp.text("L62", "Клемма трёхуровневая (по настройке Каталога)", st=st_txt); sp.formula("M62", "SUM(N51:N60)", st=st_num)
sp.formula("P62", f'IFERROR(INDEX({C}$G$4:$G$500,MATCH({C}$P$7,{C}$C$4:$C$500,0)),"")', st=st_txt)
sp.formula("Q62", f'IFERROR(INDEX({C}$H$4:$H$500,MATCH({C}$P$7,{C}$C$4:$C$500,0)),"")', st=st_txt)
sp.text("L63", "Клемма проходная доп. (по настройке Каталога)", st=st_txt); sp.formula("M63", "SUM(O51:O60)", st=st_num)
sp.formula("P63", f'IFERROR(INDEX({C}$G$4:$G$500,MATCH({C}$P$8,{C}$C$4:$C$500,0)),"")', st=st_txt)
sp.formula("Q63", f'IFERROR(INDEX({C}$H$4:$H$500,MATCH({C}$P$8,{C}$C$4:$C$500,0)),"")', st=st_txt)
sp.text("L64", "Трёхуровневая = L/N/PE одной линии; у 5-жильных ещё 2 проходные на доп. жилы.", st=st_txt)
sp.save()

open(stp, "w", encoding="utf-8").write(st)
# --- совместимость с Excel: фильтр Каталога + app.xml без устаревших списков листов
wbp = WD + "/xl/workbook.xml"; wb = open(wbp, encoding="utf-8").read()
names = re.findall(r'<sheet name="([^"]+)"', wb)
ci = names.index("Каталог")
if 'localSheetId="%d"' % ci not in wb:
    wb = wb.replace("<definedNames>", f'<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="{ci}" hidden="1">Каталог!$A$3:$M$500</definedName>', 1)
open(wbp, "w", encoding="utf-8").write(wb)
ap = WD + "/docProps/app.xml"
if os.path.exists(ap):
    a = open(ap, encoding="utf-8").read()
    a = re.sub(r"<HeadingPairs>.*?</HeadingPairs>", "", a, flags=re.S)
    a = re.sub(r"<TitlesOfParts>.*?</TitlesOfParts>", "", a, flags=re.S)
    open(ap, "w", encoding="utf-8").write(a)

if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n in order:
        p = os.path.join(WD, n)
        if os.path.exists(p): z.write(p, n)
print("ok", xl_hash(PASSWORD))
