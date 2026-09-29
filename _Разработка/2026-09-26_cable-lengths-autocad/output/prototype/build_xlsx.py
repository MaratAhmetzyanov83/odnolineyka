import re, json, os, shutil, zipfile, copy
from lxml import etree
S="/tmp/claude-0/-home-claude/2c92c9a1-ecf4-5f1c-89ce-60bce7f86197/scratchpad"
SRC="/mnt/user-data/uploads/01_ПИР_ЭОМ/Однолинейка ЩР ЭОМ.xlsx"
OUT="/home/claude/Однолинейка ЩР ЭОМ (Остров).xlsx"
W=S+"/x2"
os.makedirs(W,exist_ok=True)
with zipfile.ZipFile(SRC) as z: z.extractall(W); names=z.namelist()
M="http://schemas.openxmlformats.org/spreadsheetml/2006/main"; ns={"m":M}
def q(t): return "{%s}%s"%(M,t)
def colnum(c):
    n=0
    for ch in c: n=n*26+ord(ch)-64
    return n
def split(r):
    m=re.match(r"([A-Z]+)(\d+)$",r); return m.group(1),int(m.group(2))
def load(sh): return etree.parse(f"{W}/xl/worksheets/{sh}.xml")
def save(t,sh): t.write(f"{W}/xl/worksheets/{sh}.xml",xml_declaration=True,encoding="UTF-8",standalone=True)
def clear_cell(c, formulas=False):
    f=c.find("m:f",ns)
    if f is not None and not formulas: return False
    for ch in list(c):
        if ch.tag in (q("f"),q("v"),q("is")): c.remove(ch)
    if "t" in c.attrib: del c.attrib["t"]
    return True
def clear(sh, cols, r0, r1, formulas_in=()):
    t=load(sh); n=0
    for c in t.iterfind(".//m:sheetData/m:row/m:c",ns):
        col,row=split(c.get("r"))
        if col in cols and r0<=row<=r1:
            if clear_cell(c, formulas=col in formulas_in): n+=1
    save(t,sh); return n
# --- 1. очистка чужих исходных данных
log={}
log["Исходные данные"]=clear("sheet3",set("C D E F G H I J K L N O Q R S".split()),3,1000,formulas_in={"E","K","L"})
log["Разбивка_по_щитам H"]=clear("sheet4",{"H"},3,1000)
log["Нагрузка_щитов E"]=clear("sheet5",{"E"},7,1000)
log["Спецификация"]=clear("sheet7",set("B C D E F G H I J".split()),3,1000)
# --- 2. заполнение строк из DWG
data=json.load(open(S+"/groups_final.json"))
order=["Ввод","KV","На ИБП","От ИБП","R","RС","В","ТП","ЭK","EK","TE","SH","L","DALI","LD","LED","АПС","К","K","ПУ","ПВ","КМН","УВ","ПН","WD","КУП","PE"]
def key(d):
    g=d["grp"]; pre=g.split(".")[0]
    nums=[int(x) for x in re.findall(r"\d+",g[len(pre):])]
    return (order.index(pre) if pre in order else 99, pre, nums)
data.sort(key=key)
t=load("sheet3"); sd=t.find("m:sheetData",ns)
rows={int(r.get("r")):r for r in sd.iterfind("m:row",ns)}
def row_el(n):
    if n in rows: return rows[n]
    r=etree.SubElement(sd,q("row")); r.set("r",str(n)); rows[n]=r
    # keep order
    items=sorted(sd.iterfind("m:row",ns), key=lambda e:int(e.get("r")))
    for e in items: sd.remove(e); sd.append(e)
    return r
def cell(n, col, style_from=None):
    r=row_el(n); ref=f"{col}{n}"
    for c in r.iterfind("m:c",ns):
        if c.get("r")==ref: return c
    c=etree.Element(q("c")); c.set("r",ref)
    if style_from is not None and style_from.get("s"): c.set("s",style_from.get("s"))
    # insert in column order
    kids=list(r.iterfind("m:c",ns)); pos=len(kids)
    for i,k in enumerate(kids):
        if colnum(split(k.get("r"))[0])>colnum(col): pos=i; break
    if pos<len(kids): kids[pos].addprevious(c)
    else: r.append(c)
    return c
def set_num(c,v):
    clear_cell(c,True); e=etree.SubElement(c,q("v")); e.text=repr(float(v)) if not float(v).is_integer() else str(int(v))
def set_str(c,s):
    clear_cell(c,True); c.set("t","inlineStr"); i=etree.SubElement(c,q("is")); tt=etree.SubElement(i,q("t")); tt.text=s
    tt.set("{http://www.w3.org/XML/1998/namespace}space","preserve")
def set_f(c,f):
    clear_cell(c,True); e=etree.SubElement(c,q("f")); e.text=f
# заголовки новых колонок (стиль как у AT2)
hdr_style=None
for c in rows[2].iterfind("m:c",ns):
    if c.get("r")=="AT2": hdr_style=c
for col,txt in (("AU","DWG: гофра ПВХ по плану, м"),("AV","DWG: гофра ПНД по плану, м"),("AW","DWG: хэндлы полилиний"),("AX","DWG: примечание")):
    set_str(cell(2,col,hdr_style),txt)
r=3
for d in data:
    has=bool(d["handles"])
    set_str(cell(r,"D"),d["grp"])
    set_num(cell(r,"G"),1)
    if has:
        if d["pvh"]: set_num(cell(r,"AU"),round(d["pvh"],2))
        if d["pnd"]: set_num(cell(r,"AV"),round(d["pnd"],2))
        set_f(cell(r,"I"),f"IF(COUNT(AU{r}:AV{r})=0,\"\",ROUNDUP(SUM(AU{r}:AV{r}),0))")
        set_num(cell(r,"S"), 1 if d["pvh"]>=d["pnd"] else 2)
        set_f(cell(r,"K"),f"IF(OR(I{r}=\"\",S{r}=\"\"),\"\",IF(S{r}=1,MAX(0,M{r}-ROUNDUP(N(AV{r}),0)-4),ROUNDUP(N(AU{r}),0)))")
        set_f(cell(r,"L"),f"IF(OR(I{r}=\"\",S{r}=\"\"),\"\",IF(S{r}=2,MAX(0,M{r}-ROUNDUP(N(AU{r}),0)-4),ROUNDUP(N(AV{r}),0)))")
        set_str(cell(r,"AW"),", ".join(d["handles"]))
    if d["notes"]: set_str(cell(r,"AX"),"; ".join(d["notes"]))
    r+=1
last=r-1
save(t,"sheet3")
# --- 3. calcChain долой + пересчёт при открытии
os.remove(f"{W}/xl/calcChain.xml")
p=f"{W}/xl/_rels/workbook.xml.rels"; s=open(p,encoding="utf8").read()
s=re.sub(r'<Relationship [^>]*Target="calcChain.xml"[^>]*/>','',s); open(p,"w",encoding="utf8").write(s)
p=f"{W}/[Content_Types].xml"; s=open(p,encoding="utf8").read()
s=re.sub(r'<Override [^>]*PartName="/xl/calcChain.xml"[^>]*/>','',s); open(p,"w",encoding="utf8").write(s)
p=f"{W}/xl/workbook.xml"; s=open(p,encoding="utf8").read()
s=s.replace('<calcPr calcId="191029"/>','<calcPr calcId="191029" fullCalcOnLoad="1"/>'); open(p,"w",encoding="utf8").write(s)
# --- 4. упаковка в исходном порядке частей
if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT,"w",zipfile.ZIP_DEFLATED) as z:
    for n in names:
        if n=="xl/calcChain.xml": continue
        z.write(f"{W}/{n}",n)
print(log, "строк заполнено:", last-2, "до строки", last)
