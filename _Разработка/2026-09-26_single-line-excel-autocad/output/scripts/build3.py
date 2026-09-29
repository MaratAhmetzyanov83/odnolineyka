"""v2: separate sheet «Однолинейка» + «Справочник_схемы», auto breaker selection.
Original sheets are not modified (only styles, workbook, rels, content types)."""
import json, re, os, shutil, zipfile, html

SP = "/tmp/claude-0/-home-claude/5e645175-0e52-5350-b586-c0cb5927e030/scratchpad/"
SRC = "/root/.claude/uploads/5e645175-0e52-5350-b586-c0cb5927e030/aa5313f2-__________________.xlsx"
OUT = SP + "out3/Однолинейка ЩР ЭОМ_v3.2.xlsx"
os.makedirs(SP + "out3", exist_ok=True)
FIRST, LAST = 3, 1000
RZ = "'Разбивка_по_щитам'!"
SHN = "Однолинейка"
REFN = "Справочник_схемы"

# ------------------------------------------------------------ example inputs (from DWG, see build_xlsx.py)
slots = json.load(open(SP + "slots.json"))
MANUAL = set(range(0, 18)) | set(range(71, 76)) | set(range(238, 241)) | set(range(256, 268)) | set(range(291, 295)) | set(range(313, 371))
MOD_START = {295: "уточнить-Z", 307: "уточнить-R", 76: "ZIOMB24V2", 103: "ZIO-MB24", 120: "ZIO-MB24", 145: "ZIO-MB24", 171: "ZIOMB8V3", 179: "ZDINDX4", 183: "ZIOMB24V2"}
MOD_END = {211, 313}
DEVICE = {88: "T 12V 54W", 103: "T 24V 100W"}
BUS_KM = set(range(268, 282))
NO_XT = {1, 239, 240, 291, 293} | set(range(324, 371))
inputs = {}; dwg_qf = {}
prev_ch = None; active = False
for d in slots:
    i = d["i"]; r = d["row"]; inp = {}
    line = d["line"] if d["line"] not in (None, 0, "0") else None
    if i in MANUAL:
        inp["K"] = "вручную"
    if "qf" in d and i != 1:
        q = d["qf"]
        txt = f"{q['P']} {q['A']}" + (" 30мА" if q["mA"] else "")
        dwg_qf[r] = txt
        inp["M"] = txt if i in MANUAL else "авто"
        if q["top"] and 1545 < q["top"] < 1556:
            inp["L"] = "ИБП"
        if i in BUS_KM:
            inp["L"] = "КМ-1"
        if "MK-5-1" in d["dev"] and i not in MANUAL:
            inp["O"] = "МК-5-1"
    if i in DEVICE:
        inp["O"] = DEVICE[i]
    if i in NO_XT:
        inp["Q"] = "б/к"
    if i in MANUAL:
        if i in MOD_END:
            inp["P"] = "-"
        inputs[r] = inp; continue
    if i in MOD_START:
        inp["P"] = MOD_START[i]; active = True; expected = 1
    elif i in MOD_END:
        inp["P"] = "-"; active = False; expected = None
    else:
        expected = (prev_ch + 1) if (active and prev_ch) else (1 if active else None)
    ch = int(d["ch"][0]) if d["ch"] else None
    has_xt = bool(d["xt"])
    if ch is not None:
        if ch != expected:
            inp["Q"] = ch
        elif not line and has_xt:
            inp["Q"] = "+"
        elif not line and not has_xt:
            inp["Q"] = "р"
        prev_ch = ch
    elif has_xt and i in range(212, 220):
        inp["Q"] = "="
    inputs[r] = inp
json.dump(dwg_qf, open(SP + "dwg_qf.json", "w"), ensure_ascii=False)

# ------------------------------------------------------------ formulas for sheet «Однолинейка»
def F(r):
    p = r - 1
    rng = lambda c: f"${c}$3:${c}$1000"
    grp = f"{rng('AD')},AD{r},{rng('A')},A{r}"
    nomS = 'SUBSTITUTE(SUBSTITUTE(UPPER(S{r}),"С","C"),"А","A")'.format(r=r)
    return {
        "A": f'IF({RZ}F{r}=0,"",{RZ}F{r})',
        "B": f'IF(A{r}="","",{RZ}C{r})',
        "C": f'IF(B{r}="","",INT((B{r}-1)/Столбцов_на_лист)+1)',
        "D": f'IF(OR({RZ}H{r}="",{RZ}H{r}=0),"",{RZ}H{r})',
        "E": f'IF(D{r}="","",IFERROR({RZ}J{r}&"",""))',
        "F": f'IF(D{r}="","",IFERROR(--{RZ}L{r},""))',
        "G": f'IF(D{r}="","",IFERROR(--{RZ}K{r},""))',
        "H": f'IF(D{r}="","",IFERROR(IF({RZ}M{r}=0,"",{RZ}M{r}),""))',
        "I": f'IF(D{r}="","",IFERROR(SUBSTITUTE({RZ}P{r},CHAR(10)," "),""))',
        "J": f'IF(D{r}="","",IFERROR(--VLOOKUP(D{r},Исходные_данные,21,FALSE),""))',
        # results
        "R": (f'IF(M{r}="","",IF(AO{r}=1,"QFD","QF")&AD{r})'),
        "S": (f'IF(M{r}="","",IF(AND(M{r}<>"авто",M{r}<>"+"),M{r},AP{r}&" С"&AM{r}&"А"&IF(AO{r}=1," 30мА","")))'),
        "T": f'IF(R{r}<>"",R{r},IF(OR(A{r}<>A{p},K{r}="вручную"),"",T{p}))',
        "U": f'IF(M{r}="","",SUMIFS({rng("G")},{grp},{rng("AH")},0))',
        "V": f'IF(M{r}="","",IFERROR(VLOOKUP(U{r},Автомат_номиналы,2,TRUE),""))',
        "W": (f'IF(M{r}="","",IF(NOT(ISNUMBER(AQ{r})),"",IF(U{r}>AQ{r},"ПЕРЕГРУЗ",'
              f'IF(AND(AW{r}=0,AL{r}>0,AQ{r}>AL{r}),"КАБЕЛЬ НЕ ЗАЩИЩЁН","ok"))))'),
        "X": f'IF(P{r}="-","",IF(AS{r}<>"",AS{r},IF(A{r}<>A{p},"",X{p})))',
        "Y": f'IF(X{r}="","",IF(AND(P{r}<>"",P{r}<>"-"),P{r},Y{p}))',
        "Z": (f'IF(Q{r}="=",AT{p},IF(ISNUMBER(Q{r}),Q{r},IF(X{r}="","",'
              f'IF(OR(D{r}<>"",Q{r}="+",Q{r}="р"),IF(OR(AS{r}<>"",AT{p}=""),1,AT{p}+1),""))))'),
        "AA": f'IF(AU{r}=1,"XT"&COUNTIFS($A$3:A{r},A{r},$AU$3:AU{r},1),"")',
        "AB": (f'IF(AND(D{r}="",Z{r}="",AA{r}=""),"",_xlfn.TEXTJOIN(" → ",TRUE,T{r},AV{r},'
               f'IF(X{r}<>"",X{r}&IF(Z{r}<>""," к."&Z{r},""),IF(Z{r}<>"","к."&Z{r},"")),AA{r},IF(AH{r}=1,"(в кабель выше)","")))'),
        "AC": (f'IF(OR(Z{r}="",X{r}=""),"",IF(Z{r}>IFERROR(VLOOKUP(Y{r},Модули_спр,3,FALSE),99),"ЛИШНИЙ КАНАЛ",'
               f'IF(AND(Q{r}<>"=",COUNTIFS({rng("X")},X{r},{rng("Z")},Z{r},{rng("Q")},"<>=")>1),"ДУБЛЬ КАНАЛА","")))'),
        # hidden helpers
        "AD": f'IF(M{r}<>"",COUNTIFS($A$3:A{r},A{r},$M$3:M{r},"?*"),IF(OR(A{r}<>A{p},K{r}="вручную"),"",AD{p}))',
        "AE": f'IF(D{r}="","",IFERROR(--VLOOKUP(D{r},Исходные_данные,8,FALSE),1))',
        "AF": f'IF(D{r}="","",IFERROR(LEFT(D{r},FIND(".",D{r})-1),D{r}))',
        "AG": (f'IF(AF{r}="",0,IFERROR(IF(VLOOKUP(D{r},УГО_спр,3,FALSE)="да",1,0),IFERROR(IF(VLOOKUP(AF{r},УГО_спр,3,FALSE)="да",1,0),'
               f'IFERROR(IF(VLOOKUP(LEFT(AF{r},2),УГО_спр,3,FALSE)="да",1,0),0))))'),
        "AH": f'IF(OR(Q{r}="+",AND(D{r}<>"",D{r}=D{p})),1,0)',
        "AK": f'V{r}',
        "AL": f'IF(M{r}="","",_xlfn.MINIFS({rng("J")},{grp}))',
        "AM": f'IF(M{r}="","",IF(AW{r}=1,MAX(AK{r},Мин_автомат_БП),IF(AL{r}>0,MAX(AK{r},AL{r}),AK{r})))',
        "AW": (f'IF(AV{r}="",0,IF(OR(LEFT(AV{r},2)="T ",LEFT(AV{r},2)="Т ",ISNUMBER(SEARCH("→ T ",AV{r})),'
               f'ISNUMBER(SEARCH("→ Т ",AV{r})),ISNUMBER(SEARCH("БП",AV{r}))),1,0))'),
        "BC": f'IF(OR(D{r}<>"",Z{r}<>"",AU{r}=1),T{r},"")',
        "BD": f'IF(OR(Z{r}<>"",AND(P{r}<>"",P{r}<>"-")),X{r},"")',
        "BE": f'IF(BD{r}="","",Y{r})',
        "AN": f'IF(M{r}="","",MAX(1,_xlfn.MAXIFS({rng("AE")},{grp})))',
        "AO": (f'IF(M{r}="","",IF(AND(M{r}<>"авто",M{r}<>"+"),IF(ISNUMBER(SEARCH("мА",M{r})),1,0),'
               f'IF(N{r}="да",1,IF(N{r}="нет",0,IF(_xlfn.MAXIFS({rng("AG")},{grp})>0,1,0)))))'),
        "AP": f'IF(M{r}="","",IF(AN{r}>=3,IF(AO{r}=1,"4P","3P"),IF(AO{r}=1,"2P","1P")))',
        "AQ": (f'IF(S{r}="","",IFERROR(--MID({nomS},FIND("C",{nomS})+1,'
               f'FIND("A",{nomS}&"A",FIND("C",{nomS}))-FIND("C",{nomS})-1),""))'),
        "AR": f'IF(OR(P{r}="",P{r}="-"),"",IFERROR(VLOOKUP(P{r},Модули_спр,2,FALSE),"R"))',
        "AS": f'IF(AR{r}="","",AR{r}&"."&COUNTIFS($A$3:A{r},A{r},$AR$3:AR{r},AR{r}))',
        "AT": f'IF(Z{r}<>"",Z{r},IF(OR(A{r}<>A{p},P{r}="-",AS{r}<>""),"",AT{p}))',
        "AU": f'IF(Q{r}="б/к",0,IF(OR(ISNUMBER(SEARCH("L=",{RZ}P{r})),Q{r}="+",Q{r}="="),1,0))',
        "AV": f'IF(R{r}<>"",_xlfn.TEXTJOIN(" → ",TRUE,O{r},BA{r},BB{r}),IF(T{r}="","",AV{p}))',
    }

MODULES = [("ZIO-MB24", "R", 24, "реле 16А"), ("ZIOMB24V2", "R", 24, "реле 16А"), ("ZIOMB16V3", "R", 16, "реле"),
           ("ZIOMB8V3", "R", 8, "реле"), ("ZCL-8HT230", "R", 8, "клапаны 230В"), ("ZDINDX4", "Di", 4, "диммер"),
           ("ZDID64X2", "DA", 2, "шлюз DALI"),
           ("уточнить-R", "R", 24, "реле — МОДЕЛЬ УТОЧНИТЬ (R6 в DWG)"),
           ("уточнить-Z", "Z", 8, "датчики протечки — МОДЕЛЬ УТОЧНИТЬ (Z.1 в DWG)")]
UGO = [("R.ШТК", "SOCKET", "нет"), ("R.ЩОПС", "SOCKET", "нет"), ("R.GSM", "SOCKET", "нет"), ("R", "SOCKET", "да"), ("L", "LAMP", "нет"), ("LD", "LAMP", "нет"), ("LED", "LAMP", "нет"), ("DALI", "LAMP", "нет"),
       ("SH", "MOTOR", "нет"), ("В", "MOTOR", "нет"), ("K", "MOTOR", "нет"), ("К", "MOTOR", "нет"), ("ПУ", "MOTOR", "нет"),
       ("УВ", "MOTOR", "нет"), ("ТП", "HEAT", "да"), ("ПН", "HEAT", "да"), ("КМН", "HEAT", "нет"), ("ЭK", "VALVE", "нет"),
       ("ЭК", "VALVE", "нет"), ("TE", "SERVO", "нет"), ("ТЕ", "SERVO", "нет"), ("АСУ", "BOX", "нет")]

def col2n(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n
def n2col(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26); s = chr(65 + rem) + s
    return s
esc = lambda s: html.escape(str(s), quote=False)
def cstr(ref, s, st): return f'<c r="{ref}" s="{st}" t="inlineStr"><is><t xml:space="preserve">{esc(s)}</t></is></c>'
def cnum(ref, v, st): return f'<c r="{ref}" s="{st}"><v>{v}</v></c>'
def cf(ref, f, st): return f'<c r="{ref}" s="{st}"><f>{esc(f)}</f></c>'
def cempty(ref, st): return f'<c r="{ref}" s="{st}"/>'

# ------------------------------------------------------------ unzip
wd = SP + "x4"
shutil.rmtree(wd, ignore_errors=True)
with zipfile.ZipFile(SRC) as z:
    z.extractall(wd); order = z.namelist()

# ------------------------------------------------------------ styles
sp = wd + "/xl/styles.xml"; st = open(sp, encoding="utf-8").read()
def add(tag, xml):
    global st
    m = re.search(r'<%s count="(\d+)"' % tag, st); n = int(m.group(1))
    st = st.replace(m.group(0), '<%s count="%d"' % (tag, n + 1), 1)
    k = st.find("</%s>" % tag); st = st[:k] + xml + st[k:]
    return n
font = lambda extra="", sz=10, color='<color theme="1"/>': add("fonts", f'<font>{extra}<sz val="{sz}"/>{color}<name val="Calibri"/><family val="2"/><charset val="204"/><scheme val="minor"/></font>')
f_n = font(); f_b = font("<b/>"); f_title = font("<b/>", 11); f_grey = font("", 9, '<color rgb="FF7F7F7F"/>')
fill = lambda rgb: add("fills", f'<fill><patternFill patternType="solid"><fgColor rgb="{rgb}"/><bgColor indexed="64"/></patternFill></fill>')
fl_src, fl_in, fl_res, fl_head, fl_hid = fill("FFF2F2F2"), fill("FFFFF2CC"), fill("FFE2EFDA"), fill("FFDDEBF7"), fill("FFFAFAFA")
b_thin = add("borders", '<border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border>')
def xf(fnt, fl, al='<alignment vertical="center"/>', num=0):
    return add("cellXfs", f'<xf numFmtId="{num}" fontId="{fnt}" fillId="{fl}" borderId="{b_thin}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">{al}</xf>')
CEN = '<alignment horizontal="center" vertical="center"/>'
S_SRC = xf(f_n, fl_src); S_SRC_C = xf(f_n, fl_src, CEN); S_SRC_N = xf(f_n, fl_src, CEN, 2)
S_IN = xf(f_n, fl_in, CEN); S_RES = xf(f_n, fl_res, CEN); S_RES_L = xf(f_n, fl_res); S_RES_B = xf(f_b, fl_res, CEN)
S_HID = xf(f_grey, fl_hid, CEN)
S_HEAD = xf(f_b, fl_head, '<alignment horizontal="center" vertical="center" wrapText="1"/>')
S_T_SRC = xf(f_title, fl_src, '<alignment horizontal="left" vertical="center"/>')
S_T_IN = xf(f_title, fl_in, '<alignment horizontal="left" vertical="center"/>')
S_T_RES = xf(f_title, fl_res, '<alignment horizontal="left" vertical="center"/>')
S_T = add("cellXfs", f'<xf numFmtId="0" fontId="{f_title}" fillId="0" borderId="0" xfId="0" applyFont="1"/>')
S_REF = xf(f_n, 0); S_REF_IN = xf(f_n, fl_in); S_REF_RES = xf(f_n, fl_res, CEN)
dxf = lambda x: add("dxfs", f"<dxf>{x}</dxf>")
D_RED = dxf('<font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill>')
D_OK = dxf('<font><color rgb="FF006100"/></font>')
D_BAND = dxf('<fill><patternFill><bgColor rgb="FFEAF2FB"/></patternFill></fill>')
D_START = dxf('<font><b/></font><border><top style="medium"><color rgb="FF2F5597"/></top></border>')
D_MAN = dxf('<font><i/><color rgb="FF808080"/></font>')
open(sp, "w", encoding="utf-8").write(st)

# ------------------------------------------------------------ sheet «Однолинейка» (v3: columns follow the diagram)
# old letter (v2 formulas; BA/BB = аппарат 2/3) -> new letter, in the order of the diagram
ORDER_VIS = ["A", "B", "C",            # место
             "L",                      # шина
             "M", "N", "BC", "S",      # автомат: ввод, 30мА, QF (только в занятых строках), итог
             "O", "BA", "BB",          # аппарат 1..3
             "P", "BD", "Q", "Z",      # модуль ввод, модуль, канал ввод, канал
             "AA",                     # клемма
             "I",                      # кабель
             "D", "E", "F", "G", "H",  # отходящая линия
             "W", "AC",                # проверки
             "K", "U", "V", "J", "BE", "AB"]   # служебное видимое
ORDER_HID = ["AD", "T", "X", "Y", "R", "AE", "AF", "AG", "AH", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW"]
REMAP = {}
for k, old in enumerate(ORDER_VIS):
    REMAP[old] = n2col(k + 1)
GAP = n2col(len(ORDER_VIS) + 1)
for k, old in enumerate(ORDER_HID):
    REMAP[old] = n2col(len(ORDER_VIS) + 2 + k)
MARK = n2col(len(ORDER_VIS) + 2 + len(ORDER_HID) + 1)
N_ = lambda old: REMAP[old]
_ref = re.compile(r'(?<![A-Za-z0-9_!\.\$А-Яа-яЁё])(\$?)([A-Z]{1,2})(\$?)(\d+)(?![A-Za-z0-9_(])')
def remap_formula(f):
    out = []; parts = re.split(r'("[^"]*")', f)
    for k, part in enumerate(parts):
        if k % 2 == 1:
            out.append(part); continue
        # protect cross-sheet refs  'Sheet'!X1
        segs = re.split(r"('[^']*'!\$?[A-Z]{1,2}\$?\d+)", part)
        for j, sg in enumerate(segs):
            if j % 2 == 1:
                out.append(sg)
            else:
                out.append(_ref.sub(lambda m: f"{m.group(1)}{REMAP[m.group(2)]}{m.group(3)}{m.group(4)}", sg))
    return "".join(out)
def F3(r):
    return {REMAP[k]: remap_formula(v) for k, v in F(r).items()}
# self-check: every old column used in formulas must be in REMAP
for k, v in F(10).items():
    for m in _ref.finditer(re.sub(r'"[^"]*"', '', re.sub(r"'[^']*'!\$?[A-Z]{1,2}\$?\d+", "", v))):
        assert m.group(2) in REMAP, (k, m.group(0), v)

SECTIONS_OLD = [
    ("МЕСТО", ["A", "B", "C"], "FFF2F2F2"),
    ("ШИНА", ["L"], "FFFCE4D6"),
    ("АВТОМАТ", ["M", "N", "BC", "S"], "FFDDEBF7"),
    ("АППАРАТЫ (по порядку после автомата)", ["O", "BA", "BB"], "FFEDE7F6"),
    ("МОДУЛЬ / КАНАЛ", ["P", "BD", "Q", "Z"], "FFE2EFDA"),
    ("КЛЕММА", ["AA"], "FFFFF2CC"),
    ("КАБЕЛЬ", ["I"], "FFF2F2F2"),
    ("ОТХОДЯЩАЯ ЛИНИЯ", ["D", "E", "F", "G", "H"], "FFF2F2F2"),
    ("ПРОВЕРКИ", ["W", "AC"], "FFFFFFFF"),
    ("СЛУЖЕБНОЕ — сюда можно не смотреть", ["K", "U", "V", "J", "BE", "AB"], "FFF7F7F7"),
]
SECTIONS = [(t, [N_(c) for c in cs], rgb) for t, cs, rgb in SECTIONS_OLD]
INPUTS = {N_(c) for c in ["L", "M", "N", "O", "BA", "BB", "P", "Q", "K"]}
HID = [N_(c) for c in ORDER_HID]
_H = {
    "A": "Щит", "B": "№\nместа", "C": "Лист\nсхемы",
    "L": "Шина\n(пусто =\nосновная)",
    "M": "Автомат\n«авто» или текст\n(1-я строка группы)", "N": "30 мА\nда / нет\n(пусто = авто)",
    "BC": "QF", "S": "Автомат", "T": "QF гр.", "X": "модуль гр.", "Y": "модель гр.",
    "O": "Аппарат 1\n(МК-5-1, КМ,\nT 24V 100W, БП…)", "BA": "Аппарат 2", "BB": "Аппарат 3",
    "P": "Модуль\n(модель в начале,\n«-» = конец)", "BD": "Модуль", "Q": "Канал\n(пусто = авто,\n№, + = р б/к)", "Z": "Канал",
    "AA": "Клемма", "I": "Кабель",
    "D": "Линия", "E": "Наименование", "F": "Р, кВт", "G": "Ток, А", "H": "Фаза",
    "W": "Автомат", "AC": "Канал",
    "K": "Режим\n(«вручную» —\nне рисовать)", "U": "Ток\nгруппы, А", "V": "Номинал\nпо току, А", "J": "Автомат\nпо кабелю, А",
    "BE": "Модель\nмодуля", "AB": "Цепь одной строкой",
    "AD": "гр.", "R": "QF нач.", "AE": "фаз", "AF": "код", "AG": "30мА лин.", "AH": "в кабель выше", "AK": "по току",
    "AL": "мин. по кабелям", "AM": "номинал", "AN": "фаз гр.", "AO": "30мА гр.", "AP": "полюса", "AQ": "номинал итог",
    "AR": "префикс", "AS": "модуль нач.", "AT": "посл. канал", "AU": "клемма", "AV": "аппараты гр.", "AW": "БП/тр-р",
}
HEAD = {N_(k): v for k, v in _H.items()}
_W = {"A": 6, "B": 6, "C": 6, "L": 9, "M": 15, "N": 8, "BC": 8, "S": 15, "O": 13, "BA": 11, "BB": 11,
      "P": 12, "BD": 7, "Q": 9, "Z": 6, "AA": 7, "I": 26, "D": 16, "E": 40, "F": 7, "G": 7, "H": 6,
      "W": 17, "AC": 14, "K": 11, "U": 8, "V": 8, "J": 9, "BE": 11, "AB": 50}
WIDTH = {N_(k): v for k, v in _W.items()}
NUMC = {N_(c) for c in ["F", "G", "U"]}
BOLDC = {N_(c) for c in ["BC", "S", "D"]}
LEFTC = {N_(c) for c in ["I", "E", "AB"]}

# styles per section: header fill, body fill (input = yellow), left border at section start
b_left = add("borders", '<border><left style="medium"><color rgb="FF7F7F7F"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border>')
fills = {}
def getfill(rgb):
    if rgb not in fills:
        fills[rgb] = fill(rgb)
    return fills[rgb]
_xf_cache = {}
def sxf(fnt, fl, left, align, num=0):
    key = (fnt, fl, left, align, num)
    if key not in _xf_cache:
        _xf_cache[key] = add("cellXfs", f'<xf numFmtId="{num}" fontId="{fnt}" fillId="{fl}" borderId="{b_left if left else b_thin}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">{align}</xf>')
    return _xf_cache[key]
LEFT = '<alignment vertical="center"/>'
WRAPC = '<alignment horizontal="center" vertical="center" wrapText="1"/>'
TITLEAL = '<alignment horizontal="center" vertical="center"/>'
colstyle = {}; headstyle = {}; titlestyle = {}; section_first = {}
for title_, cols_, rgb in SECTIONS:
    for k, c in enumerate(cols_):
        first = (k == 0)
        body_fill = fl_in if c in INPUTS else (getfill(rgb) if rgb != "FFFFFFFF" else 0)
        num = 2 if c in NUMC else 0
        fnt = f_b if c in BOLDC else f_n
        al = LEFT if c in LEFTC else CEN
        colstyle[c] = sxf(fnt, body_fill, first, al, num)
        headstyle[c] = sxf(f_b, fl_in if c in INPUTS else getfill(rgb if rgb != "FFFFFFFF" else "FFF2F2F2"), first, WRAPC)
        titlestyle[c] = sxf(f_title, getfill(rgb if rgb != "FFFFFFFF" else "FFF2F2F2"), first, TITLEAL)
    section_first[cols_[0]] = (title_, cols_[-1])
open(sp, "w", encoding="utf-8").write(st)

VIS = [c for _, cs, _ in SECTIONS for c in cs]
# ---- служебные столбцы для LISP (латинские значения)
_m = col2n(MARK)
LHC = [n2col(_m + 1 + k) for k in range(9)]
def LH(r):
    I, J, K = N_("O"), N_("BA"), N_("BB")
    dev = lambda c: (f'IF({c}{r}="","",IF(OR(LEFT({c}{r},1)="T",LEFT({c}{r},1)="Т"),"T",IF(LEFT(UPPER({c}{r}),2)="БП","PS",'
                     f'IF(OR(LEFT(UPPER({c}{r}),2)="МК",LEFT(UPPER({c}{r}),2)="MK"),"MK",IF(OR(LEFT(UPPER({c}{r}),2)="КМ",LEFT(UPPER({c}{r}),2)="KM"),"KM","X")))))')
    Rl, AL, AJ, Nq = N_("D"), N_("AF"), N_("R"), N_("Q")
    return list(zip(LHC, [
        f'IF({N_("K")}{r}="вручную",0,1)',
        f'IF({N_("L")}{r}="",0,1)',
        f'IF({AJ}{r}="",0,IF(LEFT({AJ}{r},3)="QFD",1,0))',
        f'IF({Nq}{r}="=","=",IF({Nq}{r}="+","+",IF({Nq}{r}="р","R","")))',
        dev(I), dev(J), dev(K),
        (f'IF({Rl}{r}="","",IFERROR(VLOOKUP({Rl}{r},УГО_спр,2,FALSE),IFERROR(VLOOKUP({AL}{r},УГО_спр,2,FALSE),'
         f'IFERROR(VLOOKUP(LEFT({AL}{r},2),УГО_спр,2,FALSE),"BOX"))))'),
        f'IF({Rl}{r}="",0,1)']))
LH_HEAD = ["draw", "bus2", "rcd", "chtype", "dev1", "dev2", "dev3", "sym", "hasline"]
json.dump({"MARK": MARK, "LH": dict(zip(LH_HEAD, LHC)), "REMAP": REMAP}, open(SP + "cols_v32.json", "w"), ensure_ascii=False)
rows = []
title = []
for c in VIS:
    if c in section_first:
        title.append(cstr(f"{c}1", section_first[c][0], titlestyle[c]))
    else:
        title.append(cempty(f"{c}1", titlestyle[c]))
title += [cstr(f"{HID[0]}1", "расчётные (скрыты, «+» над столбцами)", S_T), cstr(f"{MARK}1", "SX_DATA", S_HID)]
rows.append(('<row r="1" ht="22" customHeight="1">', title))
rows.append(('<row r="2" ht="58" customHeight="1">', [cstr(f"{c}2", HEAD[c], headstyle[c]) for c in VIS] +
             [cstr(f"{c}2", HEAD[c], S_HEAD) for c in HID] + [cstr(f"{c}2", h, S_HEAD) for c, h in zip(LHC, LH_HEAD)]))
for r in range(FIRST, LAST + 1):
    fm = F3(r); cells = []
    inp = {REMAP[k]: v for k, v in inputs.get(r, {}).items()}
    for c in VIS:
        if c in INPUTS:
            v = inp.get(c)
            cells.append(cempty(f"{c}{r}", colstyle[c]) if v is None else
                         (cnum(f"{c}{r}", v, colstyle[c]) if isinstance(v, int) else cstr(f"{c}{r}", v, colstyle[c])))
        else:
            cells.append(cf(f"{c}{r}", fm[c], colstyle[c]))
    for c in HID:
        cells.append(cf(f"{c}{r}", fm[c], S_HID))
    for c, f in LH(r):
        cells.append(cf(f"{c}{r}", f, S_HID))
    rows.append((f'<row r="{r}">', cells))
sheetdata = "<sheetData>" + "".join(h + "".join(sorted(c, key=lambda x: col2n(re.match(r'<c r="([A-Z]+)', x).group(1)))) + "</row>" for h, c in rows) + "</sheetData>"
cols = "".join(f'<col min="{col2n(c)}" max="{col2n(c)}" width="{WIDTH[c]}" customWidth="1"/>' for c in VIS)
cols += f'<col min="{col2n(GAP)}" max="{col2n(GAP)}" width="3" customWidth="1"/>'
cols += f'<col min="{col2n(HID[0])}" max="{col2n(HID[-1])}" width="9" hidden="1" outlineLevel="1" customWidth="1"/>'
cols += f'<col min="{col2n(MARK)}" max="{col2n(LHC[-1])}" width="6" hidden="1" customWidth="1"/>'
merges = [f"{c}1:{section_first[c][1]}1" for c in section_first if section_first[c][1] != c]
LASTV = N_("AB"); area = f"A3:{LASTV}1000"; cK, cM, cAD, cW, cAC, cN, cP = N_("K"), N_("M"), N_("AD"), N_("W"), N_("AC"), N_("N"), N_("P")
cfx = (f'<conditionalFormatting sqref="{area}">'
       f'<cfRule type="expression" dxfId="{D_MAN}" priority="1"><formula>${cK}3="вручную"</formula></cfRule>'
       f'<cfRule type="expression" dxfId="{D_START}" priority="2"><formula>${cM}3&lt;&gt;""</formula></cfRule>'
       f'<cfRule type="expression" dxfId="{D_BAND}" priority="9"><formula>AND(ISNUMBER(${cAD}3),ISODD(${cAD}3),LEN(${N_("BC")}3)&gt;0)</formula></cfRule>'
       f'</conditionalFormatting>'
       f'<conditionalFormatting sqref="{cW}3:{cW}1000">'
       f'<cfRule type="cellIs" dxfId="{D_OK}" priority="4" operator="equal"><formula>"ok"</formula></cfRule>'
       f'<cfRule type="expression" dxfId="{D_RED}" priority="3"><formula>AND({cW}3&lt;&gt;"",{cW}3&lt;&gt;"ok")</formula></cfRule>'
       f'</conditionalFormatting>'
       f'<conditionalFormatting sqref="{cAC}3:{cAC}1000"><cfRule type="expression" dxfId="{D_RED}" priority="5"><formula>LEN({cAC}3)&gt;0</formula></cfRule></conditionalFormatting>')
dv = ('<dataValidations count="4">'
      f'<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="{cK}3:{cK}1000"><formula1>"вручную"</formula1></dataValidation>'
      f'<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="{cM}3:{cM}1000"><formula1>"авто"</formula1></dataValidation>'
      f'<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="{cN}3:{cN}1000"><formula1>"да,нет"</formula1></dataValidation>'
      f'<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="{cP}3:{cP}1000"><formula1>Модели_список</formula1></dataValidation>'
      '</dataValidations>')
NS = ('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
sheet_xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet {NS}>'
             '<sheetPr><tabColor rgb="FF2F5597"/><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>'
             f'<dimension ref="A1:{LHC[-1]}1000"/>'
             '<sheetViews><sheetView tabSelected="0" zoomScale="90" workbookViewId="0">'
             '<pane xSplit="3" ySplit="2" topLeftCell="D3" activePane="bottomRight" state="frozen"/>'
             f'<selection pane="bottomRight" activeCell="{cM}3" sqref="{cM}3"/></sheetView></sheetViews>'
             '<sheetFormatPr defaultRowHeight="15" outlineLevelCol="1"/>'
             f'<cols>{cols}</cols>{sheetdata}'
             f'<autoFilter ref="A2:{LASTV}1000"/>'
             f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in merges) + '</mergeCells>'
             f'{cfx}{dv}'
             '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>')

# ------------------------------------------------------------ sheet «Справочник_схемы»
rr = {}
def put(r, cell): rr.setdefault(r, []).append(cell)
put(1, cstr("A1", "Модули умного дома (дополняйте)", S_T)); put(1, cstr("F1", "Код потребителя → УГО и 30 мА по умолчанию", S_T))
put(1, cstr("J1", "Настройки", S_T))
for c, h in zip("ABCD", ["Модель", "Префикс", "Каналов", "Тип"]): put(3, cstr(f"{c}3", h, S_HEAD))
for c, h in zip("FGH", ["Код", "УГО", "30 мА"]): put(3, cstr(f"{c}3", h, S_HEAD))
put(3, cstr("E3", "В проекте", S_HEAD))
for k in range(37):  # rows 4..40
    r = 4 + k
    if k < len(MODULES):
        m = MODULES[k]
        put(r, cstr(f"A{r}", m[0], S_REF_IN)); put(r, cstr(f"B{r}", m[1], S_REF_IN)); put(r, cnum(f"C{r}", m[2], S_REF_IN)); put(r, cstr(f"D{r}", m[3], S_REF_IN))
    else:
        for c in "ABCD": put(r, cempty(f"{c}{r}", S_REF_IN))
    put(r, cf(f"E{r}", f'IF(A{r}="","",COUNTIF({SHN}!${N_("P")}$3:${N_("P")}$1000,A{r}))', S_REF_RES))
    if k < len(UGO):
        u = UGO[k]; put(r, cstr(f"F{r}", u[0], S_REF_IN)); put(r, cstr(f"G{r}", u[1], S_REF_IN)); put(r, cstr(f"H{r}", u[2], S_REF_IN))
    else:
        for c in "FGH": put(r, cempty(f"{c}{r}", S_REF_IN))
put(3, cstr("J3", "Столбцов схемы на лист A3", S_REF)); put(3, cnum("K3", 17, S_REF_IN))
put(4, cstr("J4", "Мин. автомат перед БП / трансформатором, А", S_REF)); put(4, cnum("K4", 10, S_REF_IN))
tot = [("Итоги по однолинейке", None),
       ("Автоматов всего", f'COUNTIF({SHN}!${N_("M")}$3:${N_("M")}$1000,"?*")'),
       ("  из них с 30 мА (QFD)", f'COUNTIF({SHN}!${N_("R")}$3:${N_("R")}$1000,"QFD*")'),
       ("Клемм XT (3-уровн.)", f'COUNTIF({SHN}!${N_("AU")}$3:${N_("AU")}$1000,1)'),
       ("Перемычек «в один кабель»", f'COUNTIF({SHN}!${N_("AH")}$3:${N_("AH")}$1000,1)'),
       ("Перемычек «тот же канал»", f'COUNTIF({SHN}!${N_("Q")}$3:${N_("Q")}$1000,"==")'),
       ("Модулей всего", f'COUNTIF({SHN}!${N_("AS")}$3:${N_("AS")}$1000,"?*")'),
       ("Ошибок по автоматам", f'COUNTIFS({SHN}!${N_("W")}$3:${N_("W")}$1000,"?*",{SHN}!${N_("W")}$3:${N_("W")}$1000,"<>ok")'),
       ("Ошибок по каналам", f'COUNTIF({SHN}!${N_("AC")}$3:${N_("AC")}$1000,"?*")')]
for k, (name, f) in enumerate(tot):
    r = 6 + k
    put(r, cstr(f"J{r}", name, S_T if f is None else S_REF))
    if f: put(r, cf(f"K{r}", f, S_REF_RES))
put(17, cstr("J17", "Как подбирается автомат («авто»)", S_T))
notes = ["1. Ток группы = сумма токов линий группы (линия в нескольких строках считается один раз).",
         "2. Номинал по току — по вашей таблице «Автомат_номиналы» (Справочная F:G).",
         "3. Итоговый номинал = наибольший из: по току и минимального «автомата по кабелю» линий группы —",
         "    так для одиночной линии получается ровно её автомат, а для группы (свет, шторы) — C10/C16 по кабелю.",
         "4. Если по току нужен больший автомат, чем выдерживает кабель линии — «КАБЕЛЬ НЕ ЗАЩИЩЁН».",
         "   Если в цепочке аппаратов есть БП/трансформатор — кабели вторички им не защищаются: берётся ток и мин. номинал (K4).",
         "5. Полюса: 3 фазы → 3P (4P с 30 мА), 1 фаза → 1P (2P с 30 мА).",
         "6. 30 мА: столбец «30 мА» (да/нет) или по коду потребителя из таблицы слева (R, ТП, ПН = да).",
         "7. Свой текст в «Автомат» (например «2P С20А 30мА») — используется как есть и проверяется."]
for k, t in enumerate(notes):
    put(18 + k, cstr(f"J{18 + k}", t, S_REF))
ref_rows = "".join(f'<row r="{r}">' + "".join(sorted(c, key=lambda x: col2n(re.match(r'<c r="([A-Z]+)', x).group(1)))) + "</row>" for r, c in sorted(rr.items()))
ref_xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet {NS}>'
           '<sheetPr><tabColor rgb="FF70AD47"/></sheetPr><dimension ref="A1:K40"/>'
           '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
           '<cols><col min="1" max="1" width="16" customWidth="1"/><col min="2" max="3" width="9" customWidth="1"/>'
           '<col min="4" max="4" width="14" customWidth="1"/><col min="5" max="5" width="10" customWidth="1"/>'
           '<col min="6" max="8" width="10" customWidth="1"/><col min="9" max="9" width="3" customWidth="1"/>'
           '<col min="10" max="10" width="30" customWidth="1"/><col min="11" max="11" width="8" customWidth="1"/></cols>'
           f'<sheetData>{ref_rows}</sheetData>'
           '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="H4:H40"><formula1>"да,нет"</formula1></dataValidation></dataValidations>'
           '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')

open(wd + "/xl/worksheets/sheet12.xml", "w", encoding="utf-8").write(sheet_xml)
open(wd + "/xl/worksheets/sheet13.xml", "w", encoding="utf-8").write(ref_xml)

# ------------------------------------------------------------ workbook wiring
rel = wd + "/xl/_rels/workbook.xml.rels"; t = open(rel, encoding="utf-8").read()
WS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
t = t.replace("</Relationships>", f'<Relationship Id="rIdSX1" Type="{WS}" Target="worksheets/sheet12.xml"/>'
                                  f'<Relationship Id="rIdSX2" Type="{WS}" Target="worksheets/sheet13.xml"/></Relationships>')
t = re.sub(r'<Relationship [^>]*Target="calcChain.xml"/>', "", t)
open(rel, "w", encoding="utf-8").write(t)
ct = wd + "/[Content_Types].xml"; t = open(ct, encoding="utf-8").read()
WSCT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
t = t.replace("</Types>", f'<Override PartName="/xl/worksheets/sheet12.xml" ContentType="{WSCT}"/>'
                          f'<Override PartName="/xl/worksheets/sheet13.xml" ContentType="{WSCT}"/></Types>')
t = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", t)
open(ct, "w", encoding="utf-8").write(t)
if os.path.exists(wd + "/xl/calcChain.xml"):
    os.remove(wd + "/xl/calcChain.xml")
wbp = wd + "/xl/workbook.xml"; wb = open(wbp, encoding="utf-8").read()
# insert new sheets right after «Разбивка_по_щитам»
m = re.search(r'<sheet name="Разбивка_по_щитам"[^>]*/>', wb)
wb = wb[:m.end()] + f'<sheet name="{SHN}" sheetId="40" r:id="rIdSX1"/><sheet name="{REFN}" sheetId="41" r:id="rIdSX2"/>' + wb[m.end():]
# local sheet ids of existing defined names shift? (localSheetId refers to sheet index) -> fix
def shift_local(mm):
    k = int(mm.group(1))
    return f'localSheetId="{k + 2 if k > 3 else k}"'
wb = re.sub(r'localSheetId="(\d+)"', shift_local, wb)
names = (f'<definedName name="Модули_спр">{REFN}!$A$4:$D$40</definedName>'
         f'<definedName name="Модели_список">{REFN}!$A$4:$A$40</definedName>'
         f'<definedName name="УГО_спр">{REFN}!$F$4:$H$40</definedName>'
         f'<definedName name="Столбцов_на_лист">{REFN}!$K$3</definedName>'
         f'<definedName name="Мин_автомат_БП">{REFN}!$K$4</definedName>'
         f'<definedName name="_xlnm._FilterDatabase" localSheetId="4" hidden="1">{SHN}!$A$2:${LASTV}$1000</definedName>')
wb = wb.replace("</definedNames>", names + "</definedNames>", 1)
wb = re.sub(r"<calcPr([^>]*)/>", lambda mm: "<calcPr" + mm.group(1) + ' fullCalcOnLoad="1"/>', wb, 1)
# open on the new sheet
wb = re.sub(r'activeTab="\d+"', 'activeTab="4"', wb) if "activeTab" in wb else wb.replace("<workbookView ", '<workbookView activeTab="4" ', 1)
open(wbp, "w", encoding="utf-8").write(wb)
# deselect other tabs
for f in os.listdir(wd + "/xl/worksheets"):
    if f.endswith(".xml") and f not in ("sheet12.xml",):
        p = wd + "/xl/worksheets/" + f; s = open(p, encoding="utf-8").read()
        s2 = s.replace('tabSelected="1"', 'tabSelected="0"')
        if s2 != s: open(p, "w", encoding="utf-8").write(s2)
sh = open(wd + "/xl/worksheets/sheet12.xml", encoding="utf-8").read().replace('tabSelected="0"', 'tabSelected="1"')
open(wd + "/xl/worksheets/sheet12.xml", "w", encoding="utf-8").write(sh)

if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for name in order + ["xl/worksheets/sheet12.xml", "xl/worksheets/sheet13.xml"]:
        p = os.path.join(wd, name)
        if os.path.exists(p) and name not in z.namelist():
            z.write(p, name)
print("ok", OUT)
