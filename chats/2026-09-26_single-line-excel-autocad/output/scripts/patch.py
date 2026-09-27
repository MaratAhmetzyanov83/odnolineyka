"""Patch the user's ЖК остров workbook: nominal by consumer type (sockets 16A, light 10A),
cable-based max breaker check, restore cleared formulas, clear stale «вручную». Pure XML edits."""
import re, html, os, shutil, zipfile

SP = "/tmp/claude-0/-home-claude/5e645175-0e52-5350-b586-c0cb5927e030/scratchpad/"
SRC = SP + "ost/u.xlsx"
OUT = SP + "ost/Однолинейка ЩР ЭОМ_v3.3.xlsx"
TPL = SP + "out3/Однолинейка ЩР ЭОМ_v3.2.xlsx"
WD = SP + "ost/w"
shutil.rmtree(WD, ignore_errors=True)
with zipfile.ZipFile(SRC) as z:
    z.extractall(WD); order = z.namelist()

esc = lambda s: html.escape(s, quote=False)


def col2n(c):
    n = 0
    for ch in c: n = n * 26 + ord(ch) - 64
    return n


def sheet_path(root, name):
    wb = open(root + "/xl/workbook.xml", encoding="utf-8").read()
    rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(name), wb).group(1)
    rels = open(root + "/xl/_rels/workbook.xml.rels", encoding="utf-8").read()
    tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    return root + "/xl/" + tgt.group(1).lstrip("/").replace("xl/", "")


class Sheet:
    def __init__(self, path):
        self.path = path; self.s = open(path, encoding="utf-8").read()

    def _row_span(self, r):
        m = re.search(r'<row r="%d"[^>]*?(?:/>|>.*?</row>)' % r, self.s, flags=re.S)
        return m

    def get(self, ref):
        m = re.search(r'<c r="%s"(?:\s[^>]*?)?(?:/>|>.*?</c>)' % ref, self.s, flags=re.S)
        return m

    def style(self, ref):
        m = self.get(ref)
        if m:
            st = re.search(r'\ss="(\d+)"', m.group(0))
            return st.group(1) if st else None
        return None

    def put(self, ref, cellxml):
        m = self.get(ref)
        if m:
            self.s = self.s[:m.start()] + cellxml + self.s[m.end():]; return
        col = re.match(r"[A-Z]+", ref).group(0); r = int(ref[len(col):])
        rm = self._row_span(r)
        if rm is None:
            raise RuntimeError("row missing " + ref)
        row = rm.group(0)
        if row.endswith("/>"):
            new = row[:-2] + ">" + cellxml + "</row>"
        else:
            cells = list(re.finditer(r'<c r="([A-Z]+)\d+"', row))
            pos = None
            for c in cells:
                if col2n(c.group(1)) > col2n(col):
                    pos = c.start(); break
            if pos is None:
                pos = row.rfind("</row>")
            new = row[:pos] + cellxml + row[pos:]
        self.s = self.s[:rm.start()] + new + self.s[rm.end():]

    def formula(self, ref, f, st=None):
        st = st or self.style(ref)
        sa = f' s="{st}"' if st else ""
        self.put(ref, f'<c r="{ref}"{sa}><f>{esc(f)}</f></c>')

    def text(self, ref, t, st=None):
        st = st or self.style(ref)
        sa = f' s="{st}"' if st else ""
        self.put(ref, f'<c r="{ref}"{sa} t="inlineStr"><is><t xml:space="preserve">{esc(t)}</t></is></c>')

    def num(self, ref, v, st=None):
        st = st or self.style(ref)
        sa = f' s="{st}"' if st else ""
        self.put(ref, f'<c r="{ref}"{sa}><v>{v}</v></c>')

    def clear(self, ref):
        st = self.style(ref)
        sa = f' s="{st}"' if st else ""
        self.put(ref, f'<c r="{ref}"{sa}/>')

    def is_empty(self, ref):
        m = self.get(ref)
        return m is None or ("<f" not in m.group(0) and "<v" not in m.group(0) and "<is>" not in m.group(0))

    def save(self):
        open(self.path, "w", encoding="utf-8").write(self.s)


# shared strings (to read values)
ssx = open(WD + "/xl/sharedStrings.xml", encoding="utf-8").read()
SST = [re.sub(r"<[^>]+>", "", m) for m in re.findall(r"<si>(.*?)</si>", ssx, flags=re.S)]
SST = [html.unescape(x) for x in SST]


def value(sh, ref):
    m = sh.get(ref)
    if not m: return None
    c = m.group(0)
    v = re.search(r"<v>(.*?)</v>", c)
    if 't="s"' in c and v: return SST[int(v.group(1))]
    t = re.search(r"<t[^>]*>(.*?)</t>", c, flags=re.S)
    if t: return html.unescape(t.group(1))
    return v.group(1) if v else None


log = []
# ---------------------------------------------------------------- Справочник_схемы: мин. автомат по коду
ref = Sheet(sheet_path(WD, "Справочник_схемы"))
MIN = {"R": 16, "R.ШТК": 16, "R.ЩОПС": 16, "R.GSM": 16, "L": 10, "LD": 10, "LED": 10, "DALI": 10}
ref.text("I3", "Мин. автомат, А", st=ref.style("H3"))
for r in range(4, 41):
    code = value(ref, f"F{r}")
    st = ref.style(f"H{r}")
    if code in MIN:
        ref.num(f"I{r}", MIN[code], st=st)
    elif ref.is_empty(f"I{r}"):
        ref.put(f"I{r}", f'<c r="I{r}" s="{st}"/>' if st else f'<c r="I{r}"/>')
# notes
NOTES = {
    "3. Итоговый номинал": "3. Итоговый номинал = наибольший из: по току и «Мин. автомат» по коду потребителя (столбец I слева: розетки R — 16 А, свет L/LD/LED/DALI — 10 А) и ручного автомата из «Исходных данных».",
    "    так для одиночной": "    Для группы берётся наибольший минимум среди её линий. Пустая ячейка I = только по току.",
    "4. Если по току": "4. «Макс. автомат по кабелю» = наибольший номинал, который выдерживает сечение линии (по таблице Справочная I:J). Если итоговый номинал больше — «КАБЕЛЬ НЕ ЗАЩИЩЁН».",
}
for r in range(15, 41):
    v = value(ref, f"J{r}")
    if v:
        for k, t in NOTES.items():
            if v.startswith(k):
                ref.text(f"J{r}", t); log.append(f"note J{r}")
ref.save()

# defined name УГО_спр -> F:I
wbp = WD + "/xl/workbook.xml"; wb = open(wbp, encoding="utf-8").read()
wb2 = re.sub(r"(<definedName name=\"УГО_спр\"[^>]*>)[^<]*(</definedName>)", r"\1Справочник_схемы!$F$4:$I$40\2", wb)
assert wb2 != wb
wb2 = re.sub(r"<calcPr([^>]*)/>", lambda m: "<calcPr" + re.sub(r'\sfullCalcOnLoad="[^"]*"', "", m.group(1)) + ' fullCalcOnLoad="1"/>', wb2)
open(wbp, "w", encoding="utf-8").write(wb2)

# ---------------------------------------------------------------- Однолинейка
sh = Sheet(sheet_path(WD, "Однолинейка"))
FIRST, LAST = 3, 1000
AB = ('IF(R{r}="","",IFERROR(LOOKUP(INDEX(Справочная!$I$4:$I$19,MATCH(VLOOKUP(R{r},Исходные_данные,22,FALSE),'
      'Справочная!$J$4:$J$19,0)+1),{{6,10,16,20,25,32,40,50,63,80,100,125}}),""))')
BB = ('IF(R{r}="",0,MAX(IFERROR(--VLOOKUP(R{r},Исходные_данные,18,FALSE),0),'
      'IFERROR(--VLOOKUP(R{r},УГО_спр,4,FALSE),IFERROR(--VLOOKUP(AL{r},УГО_спр,4,FALSE),'
      'IFERROR(--VLOOKUP(LEFT(AL{r},2),УГО_спр,4,FALSE),0)))))')
AO = 'IF(E{r}="","",MAX(N(AA{r}),_xlfn.MAXIFS($BB$3:$BB$1000,$AF$3:$AF$1000,AF{r},$A$3:$A$1000,A{r})))'
AQ = 'IF(E{r}="","",IF(BA{r}=1,MAX(AO{r},Мин_автомат_БП),AO{r}))'
st_hid = sh.style("BA21")
for r in range(FIRST, LAST + 1):
    sh.formula(f"AB{r}", AB.format(r=r))
    sh.formula(f"AO{r}", AO.format(r=r))
    sh.formula(f"AQ{r}", AQ.format(r=r))
    sh.formula(f"BB{r}", BB.format(r=r), st=st_hid)
sh.text("AB2", "Макс. автомат\nпо кабелю, А")
sh.text("BB2", "мин. по типу", st=sh.style("BA2"))

# restore cleared formulas from the template (same layout)
tpl_root = SP + "ost/tpl"
shutil.rmtree(tpl_root, ignore_errors=True)
with zipfile.ZipFile(TPL) as z: z.extractall(tpl_root)
tsh = Sheet(sheet_path(tpl_root, "Однолинейка"))
restored = 0
for col in ("G", "H", "M", "O", "P"):
    for r in range(FIRST, LAST + 1):
        ref_ = f"{col}{r}"
        if sh.is_empty(ref_):
            m = tsh.get(ref_)
            f = re.search(r"<f>(.*?)</f>", m.group(0), flags=re.S) if m else None
            if f:
                sh.formula(ref_, html.unescape(f.group(1)).replace("'Разбивка_по_щитам'!", "Разбивка_по_щитам!"))
                restored += 1
log.append(f"restored {restored}")

# stale «вручную» from the previous object (places 239+ have no lines now)
cleared = 0
for r in range(241, LAST + 1):
    if value(sh, f"Y{r}") == "вручную" and not value(sh, f"R{r}"):
        sh.clear(f"Y{r}"); cleared += 1
log.append(f"cleared manual {cleared}")

# hide BB like neighbours: extend a <col> covering BA if it stops at BA
cols = re.search(r"<cols>(.*?)</cols>", sh.s, flags=re.S)
if cols:
    nBB = col2n("BB")
    covered = any(int(a) <= nBB <= int(b) for a, b in re.findall(r'<col min="(\d+)" max="(\d+)"', cols.group(1)))
    if not covered:
        sh.s = sh.s.replace("</cols>", f'<col min="{nBB}" max="{nBB}" width="6" hidden="1" outlineLevel="1" customWidth="1"/></cols>', 1)
sh.save()

# remove calcChain
ccp = WD + "/xl/calcChain.xml"
if os.path.exists(ccp): os.remove(ccp)
rp = WD + "/xl/_rels/workbook.xml.rels"; rels = open(rp, encoding="utf-8").read()
rels = re.sub(r'<Relationship [^>]*Target="calcChain.xml"[^>]*/>', "", rels); open(rp, "w", encoding="utf-8").write(rels)
ctp = WD + "/[Content_Types].xml"; ct = open(ctp, encoding="utf-8").read()
ct = re.sub(r'<Override [^>]*PartName="/xl/calcChain.xml"[^>]*/>', "", ct); open(ctp, "w", encoding="utf-8").write(ct)

if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n in order:
        p = os.path.join(WD, n)
        if os.path.exists(p): z.write(p, n)
print("ok", log)
