"""v3.4: подсчёт для спецификации (автоматы, аппараты, модули, клеммы) + свободные каналы модулей."""
import sys, os, re, shutil, zipfile
SP = "/tmp/claude-0/-home-claude/5e645175-0e52-5350-b586-c0cb5927e030/scratchpad/"
sys.path.insert(0, SP + "ost")
import xlsxlib as X

SRC = SP + "ost/u33.xlsx"
OUT = SP + "ost/Однолинейка ЩР ЭОМ_v3.4.xlsx"
WD = SP + "ost/w2"
shutil.rmtree(WD, ignore_errors=True)
with zipfile.ZipFile(SRC) as z:
    z.extractall(WD); order = z.namelist()
X.load_sst(WD)

O = "Однолинейка!"
# ---------------------------------------------------------------- помощники на листе «Однолинейка» (BM..BR, скрыты)
sh = X.Sheet(X.sheet_path(WD, "Однолинейка"))
st_h = sh.style("BA21"); st_hd = sh.style("BA2")
HELP = {
    "BM": ('тип автомата №', 'IF(H{r}="","",IF(COUNTIF($H$3:H{r},H{r})=1,MAX($BM$2:BM{p})+1,""))'),
    "BN": ('аппарат 1 №', 'IF(I{r}="","",IF(COUNTIF($I$3:I{r},I{r})+COUNTIF($J$2:J{p},I{r})+COUNTIF($K$2:K{p},I{r})=1,MAX($BN$2:BP{p})+1,""))'),
    "BO": ('аппарат 2 №', 'IF(J{r}="","",IF(COUNTIF($I$3:I{r},J{r})+COUNTIF($J$3:J{r},J{r})+COUNTIF($K$2:K{p},J{r})=1,MAX($BN$2:BP{p},BN{r})+1,""))'),
    "BP": ('аппарат 3 №', 'IF(K{r}="","",IF(COUNTIF($I$3:I{r},K{r})+COUNTIF($J$3:J{r},K{r})+COUNTIF($K$3:K{r},K{r})=1,MAX($BN$2:BP{p},BN{r},BO{r})+1,""))'),
    "BQ": ('кабель клеммы', 'IF(AY{r}=1,IFERROR(MID(Q{r},FIND(" (",Q{r})+2,FIND(")",Q{r},FIND(" (",Q{r}))-FIND(" (",Q{r})-2),"?"),"")'),
    "BR": ('вид клеммы №', 'IF(BQ{r}="","",IF(COUNTIF($BQ$3:BQ{r},BQ{r})=1,MAX($BR$2:BR{p})+1,""))'),
    "BS": ('модуль №', 'IF(AW{r}="","",COUNTIF($AW$3:AW{r},"?*"))'),
    "BT": ('модель №', 'IF(OR(AW{r}="",AI{r}=""),"",IF(COUNTIFS($AW$3:AW{r},"?*",$AI$3:AI{r},AI{r})=1,MAX($BT$2:BT{p})+1,""))'),
}
for col, (head, f) in HELP.items():
    sh.text(f"{col}2", head, st=st_hd)
for r in range(3, 1001):
    for col, (head, f) in HELP.items():
        sh.formula(f"{col}{r}", f.format(r=r, p=r - 1), st=st_h)
# скрыть BM:BR
cols = re.search(r"<cols>(.*?)</cols>", sh.s, flags=re.S)
a, b = X.col2n("BM"), X.col2n("BT")
if not any(int(x) <= a <= int(y) for x, y in re.findall(r'<col min="(\d+)" max="(\d+)"', cols.group(1))):
    sh.s = sh.s.replace("</cols>", f'<col min="{a}" max="{b}" width="8" hidden="1" outlineLevel="1" customWidth="1"/></cols>', 1)
sh.save()

# ---------------------------------------------------------------- лист «Спецификация»: блок подсчёта справа (L:O)
sp = X.Sheet(X.sheet_path(WD, "Спецификация"))
st_head = sp.style("C2"); st_txt = sp.style("C4"); st_num = sp.style("H4")
sp.text("L2", "ПОДСЧЁТ ПО ЛИСТУ «ОДНОЛИНЕЙКА» (считается сам)", st=st_head)
sp.text("M2", "Кол-во", st=st_head); sp.text("N2", "Трёхуровн.", st=st_head); sp.text("O2", "Доп. проходных", st=st_head)
row = 3
def section(title):
    global row
    sp.text(f"L{row}", title, st=st_head); row += 1

section("Автоматы и дифавтоматы (полюса, номинал, 30 мА)")
for k in range(1, 21):
    sp.formula(f"L{row}", f'IFERROR(INDEX({O}$H$3:$H$1000,MATCH({k},{O}$BM$3:$BM$1000,0)),"")', st=st_txt)
    sp.formula(f"M{row}", f'IF(L{row}="","",COUNTIF({O}$H$3:$H$1000,L{row}))', st=st_num)
    row += 1
sp.text(f"L{row}", "  всего автоматов", st=st_txt)
sp.formula(f"M{row}", f'COUNTIF({O}$H$3:$H$1000,"?*")', st=st_num); row += 2

section("Аппараты после автоматов (трансформаторы, БП, МК, КМ)")
for k in range(1, 11):
    sp.formula(f"L{row}", (f'IFERROR(INDEX({O}$I$3:$I$1000,MATCH({k},{O}$BN$3:$BN$1000,0)),'
                           f'IFERROR(INDEX({O}$J$3:$J$1000,MATCH({k},{O}$BO$3:$BO$1000,0)),'
                           f'IFERROR(INDEX({O}$K$3:$K$1000,MATCH({k},{O}$BP$3:$BP$1000,0)),"")))'), st=st_txt)
    sp.formula(f"M{row}", f'IF(L{row}="","",COUNTIF({O}$I$3:$I$1000,L{row})+COUNTIF({O}$J$3:$J$1000,L{row})+COUNTIF({O}$K$3:$K$1000,L{row}))', st=st_num)
    row += 1
row += 1

section("Модули умного дома (по модели)")
for k in range(1, 11):
    sp.formula(f"L{row}", f'IFERROR(INDEX({O}$AI$3:$AI$1000,MATCH({k},{O}$BT$3:$BT$1000,0)),"")', st=st_txt)
    sp.formula(f"M{row}", f'IF(L{row}="","",COUNTIFS({O}$AW$3:$AW$1000,"?*",{O}$AI$3:$AI$1000,L{row}))', st=st_num)
    row += 1
row += 1

section("Клеммы XT по кабелю линии (жилы × сечение)")
for k in range(1, 11):
    sp.formula(f"L{row}", f'IFERROR(INDEX({O}$BQ$3:$BQ$1000,MATCH({k},{O}$BR$3:$BR$1000,0)),"")', st=st_txt)
    sp.formula(f"M{row}", f'IF(L{row}="","",COUNTIF({O}$BQ$3:$BQ$1000,L{row}))', st=st_num)
    cores = f'IFERROR(--LEFT(SUBSTITUTE(L{row},"x","х"),FIND("х",SUBSTITUTE(L{row},"x","х"))-1),3)'
    sp.formula(f"N{row}", f'IF(L{row}="","",M{row})', st=st_num)
    sp.formula(f"O{row}", f'IF(L{row}="","",MAX(0,{cores}-3)*M{row})', st=st_num)
    row += 1
sp.text(f"L{row}", "  всего клемм XT", st=st_txt)
sp.formula(f"M{row}", f'COUNTIF({O}$AY$3:$AY$1000,1)', st=st_num); row += 2
sp.text(f"L{row}", "Трёхуровневая = L/N/PE одной линии; у 5-жильных ещё 2 проходные на доп. жилы (L2, L3 или вторые фазы света).", st=st_txt)
spec_end = row
sp.save()

# ---------------------------------------------------------------- «Справочник_схемы»: модули в проекте и свободные каналы (M:S)
ref = X.Sheet(X.sheet_path(WD, "Справочник_схемы"))
sh_hd = ref.style("A3"); sh_v = ref.style("A4"); sh_t = ref.style("A2") or sh_hd
ref.text("M2", "Модули в проекте — свободные каналы", st=sh_t)
for c, t in zip("MNOPQRS", ["Щит", "Модуль", "Модель", "Каналов", "Занято (посл. №)", "в т.ч. резерв «р»", "Свободно"]):
    ref.text(f"{c}3", t, st=sh_hd)
AG = lambda k: f'(MATCH({k},{O}$BS$1:$BS$1000,0))'
for k in range(1, 36):
    r = 3 + k
    ref.formula(f"M{r}", f'IFERROR(INDEX({O}$A$1:$A$1000,{AG(k)}),"")', st=sh_v)
    ref.formula(f"N{r}", f'IFERROR(INDEX({O}$AW$1:$AW$1000,{AG(k)}),"")', st=sh_v)
    ref.formula(f"O{r}", f'IF(N{r}="","",INDEX({O}$AI$1:$AI$1000,{AG(k)}))', st=sh_v)
    ref.formula(f"P{r}", f'IF(N{r}="","",IFERROR(VLOOKUP(O{r},Модули_спр,3,FALSE),""))', st=sh_v)
    ref.formula(f"Q{r}", f'IF(N{r}="","",_xlfn.MAXIFS({O}$O$3:$O$1000,{O}$AH$3:$AH$1000,N{r},{O}$A$3:$A$1000,M{r}))', st=sh_v)
    ref.formula(f"R{r}", f'IF(N{r}="","",COUNTIFS({O}$AH$3:$AH$1000,N{r},{O}$A$3:$A$1000,M{r},{O}$N$3:$N$1000,"р"))', st=sh_v)
    ref.formula(f"S{r}", f'IF(OR(N{r}="",P{r}=""),"",P{r}-Q{r})', st=sh_v)
ref.text("M40", "Свободных каналов всего", st=sh_hd)
ref.formula("S40", 'SUM(S4:S38)', st=sh_hd)
# ширины столбцов M:S
cols = re.search(r"<cols>(.*?)</cols>", ref.s, flags=re.S)
widths = {"M": 8, "N": 9, "O": 14, "P": 9, "Q": 11, "R": 11, "S": 10}
add = "".join(f'<col min="{X.col2n(c)}" max="{X.col2n(c)}" width="{w}" customWidth="1"/>' for c, w in widths.items())
if cols:
    used = [(int(x), int(y)) for x, y in re.findall(r'<col min="(\d+)" max="(\d+)"', cols.group(1))]
    if not any(x <= X.col2n("M") <= y for x, y in used):
        ref.s = ref.s.replace("</cols>", add + "</cols>", 1)
else:
    ref.s = ref.s.replace("<sheetData", "<cols>" + add + "</cols><sheetData", 1)
ref.save()

# ---------------------------------------------------------------- пересчёт при открытии, без calcChain
wbp = WD + "/xl/workbook.xml"; wb = open(wbp, encoding="utf-8").read()
if "fullCalcOnLoad" not in wb:
    wb = re.sub(r"<calcPr([^>]*)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb)
open(wbp, "w", encoding="utf-8").write(wb)
cc = WD + "/xl/calcChain.xml"
if os.path.exists(cc):
    os.remove(cc)
    rp = WD + "/xl/_rels/workbook.xml.rels"; rels = open(rp, encoding="utf-8").read()
    open(rp, "w", encoding="utf-8").write(re.sub(r'<Relationship [^>]*Target="calcChain.xml"[^>]*/>', "", rels))
    ctp = WD + "/[Content_Types].xml"; ct = open(ctp, encoding="utf-8").read()
    open(ctp, "w", encoding="utf-8").write(re.sub(r'<Override [^>]*PartName="/xl/calcChain.xml"[^>]*/>', "", ct))

if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n in order:
        p = os.path.join(WD, n)
        if os.path.exists(p): z.write(p, n)
print("ok spec rows to", spec_end)
