"""v3.6: ревизия. Спецификация сверяется с однолинейкой по артикулу, Каталог считает «в проекте»,
ТКЗ без летучего имени, коды без УГО, свёрнутые служебные столбцы, обновлённое руководство.
«В Акад» и «Диапазоны» не трогаем (коллеги работают по-старому через Data Link)."""
import sys, os, re, shutil, zipfile
SP = "/tmp/claude-0/-home-claude/5e645175-0e52-5350-b586-c0cb5927e030/scratchpad/"
sys.path.insert(0, SP + "ost")
import xlsxlib as X

SRC = SP + "ost/Однолинейка ЩР ЭОМ_v3.5.1.xlsx"
OUT = SP + "ost/Однолинейка ЩР ЭОМ_v3.6.xlsx"
WD = SP + "ost/w5"
shutil.rmtree(WD, ignore_errors=True)
with zipfile.ZipFile(SRC) as z:
    z.extractall(WD); order = z.namelist()
X.load_sst(WD)
O = "Однолинейка!"; C = "Каталог!"

# ---------------------------------------------------------------- стили: красный dxf
stp = WD + "/xl/styles.xml"; st = open(stp, encoding="utf-8").read()
m = re.search(r'<dxfs count="(\d+)"', st); D_RED = int(m.group(1))
st = st.replace(m.group(0), f'<dxfs count="{D_RED + 2}"', 1)
k = st.find("</dxfs>")
st = st[:k] + ('<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'
               '<dxf><font><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>') + st[k:]
D_GRN = D_RED + 1
open(stp, "w", encoding="utf-8").write(st)


def add_cf(sh, xml):
    """вставить conditionalFormatting в правильное место (перед dataValidations/hyperlinks/.../pageMargins)"""
    for tag in ("<dataValidations", "<hyperlinks", "<printOptions", "<pageMargins"):
        i = sh.s.find(tag)
        if i >= 0:
            sh.s = sh.s[:i] + xml + sh.s[i:]; return
    raise RuntimeError("no place for CF")


def ensure_row(sh, r):
    if re.search(r'<row r="%d"[ >/]' % r, sh.s): return
    rows = [(int(mm.group(1)), mm.start()) for mm in re.finditer(r'<row r="(\d+)"', sh.s)]
    pos = next((st_ for n_, st_ in rows if n_ > r), None)
    if pos is None:
        pos = sh.s.find("</sheetData>")
        if pos < 0:
            sh.s = sh.s.replace("<sheetData/>", "<sheetData></sheetData>", 1); pos = sh.s.find("</sheetData>")
    sh.s = sh.s[:pos] + f'<row r="{r}"/>' + sh.s[pos:]


def max_prio(sh):
    p = [int(x) for x in re.findall(r'priority="(\d+)"', sh.s)]
    return max(p) if p else 0


def set_cols(sh, a, b, extra=None, width=None):
    """задать атрибуты для столбцов a..b, разрезая существующие <col>"""
    a, b = X.col2n(a), X.col2n(b)
    m = re.search(r"<cols>(.*?)</cols>", sh.s, flags=re.S)
    cols = re.findall(r"<col [^>]*/>", m.group(1)) if m else []
    out = []
    covered = set()
    for c in cols:
        lo, hi = int(re.search(r'min="(\d+)"', c).group(1)), int(re.search(r'max="(\d+)"', c).group(1))
        def mk(x, y, base=c):
            return re.sub(r'max="\d+"', f'max="{y}"', re.sub(r'min="\d+"', f'min="{x}"', base))
        if hi < a or lo > b:
            out.append((lo, c)); continue
        if lo < a: out.append((lo, mk(lo, a - 1)))
        mid = mk(max(lo, a), min(hi, b))
        for kk, v in (extra or {}).items():
            mid = re.sub(r'\s%s="[^"]*"' % kk, "", mid).replace("<col ", f'<col {kk}="{v}" ', 1)
        if width:
            mid = re.sub(r'width="[^"]*"', f'width="{width}"', mid)
        out.append((max(lo, a), mid))
        covered.update(range(max(lo, a), min(hi, b) + 1))
        if hi > b: out.append((b + 1, mk(b + 1, hi)))
    # непокрытые
    x = a
    while x <= b:
        if x in covered: x += 1; continue
        y = x
        while y + 1 <= b and (y + 1) not in covered: y += 1
        attrs = " ".join(f'{kk}="{v}"' for kk, v in (extra or {}).items())
        out.append((x, f'<col min="{x}" max="{y}" width="{width or 9}" customWidth="1" {attrs}/>'))
        x = y + 1
    out.sort(key=lambda t: t[0])
    new = "<cols>" + "".join(c for _, c in out) + "</cols>"
    if m: sh.s = sh.s[:m.start()] + new + sh.s[m.end():]
    else: sh.s = sh.s.replace("<sheetData", new + "<sheetData", 1)


# ================================================================= 1. Каталог: «В проекте, шт.» (столбец N)
cat = X.Sheet(X.sheet_path(WD, "Каталог"))
s_head = cat.style("M3"); s_calc = cat.style("R4")
cat.text("N3", "В проекте, шт. (по однолинейке)", st=s_head)
for r in range(4, 501):
    f = (f'IF(OR(A{r}="Автомат",A{r}="Дифавтомат"),IF(T{r}="","",COUNTIF({O}$H$3:$H$1000,R{r})),'
         f'IF(A{r}="Аппарат",IF(C{r}="","",COUNTIF({O}$I$3:$I$1000,C{r})+COUNTIF({O}$J$3:$J$1000,C{r})+COUNTIF({O}$K$3:$K$1000,C{r})),'
         f'IF(A{r}="Модуль",IF(C{r}="","",COUNTIFS({O}$AW$3:$AW$1000,"?*",{O}$AI$3:$AI$1000,C{r})),'
         f'IF(A{r}="Клемма",IF(C{r}=$P$7,Спецификация!$M$62,IF(C{r}=$P$8,Спецификация!$M$63,"")),""))))')
    cat.formula(f"N{r}", f, st=s_calc)
set_cols(cat, "N", "N", width=11)
cat.save()

# ================================================================= 2. Спецификация: количества из однолинейки
sp = X.Sheet(X.sheet_path(WD, "Спецификация"))
s_h2 = sp.style("H2"); s_h = sp.style("H4")
# артикулы, которые Каталог умеет считать (категории с автоподсчётом); клеммы — только 3L и QUATTRO
catx = X.Sheet(X.sheet_path(WD, "Каталог"))
AUTO = set()
for r in range(4, 200):
    a = X.value(catx, f"A{r}"); art = X.value(catx, f"H{r}")
    if not art: continue
    if a in ("Автомат", "Дифавтомат", "Аппарат", "Модуль") or art in ("VPR-2.5-3L-GY", "VPR-2.5-QUATTRO-GY"):
        AUTO.add(str(art).strip())
sp.text("K2", "Было вручную (до v3.6)", st=s_h2)
LASTSP = 120
auto_rows = []
for r in range(4, LASTSP + 1):
    ensure_row(sp, r)
    art = X.value(sp, f"E{r}")
    if art is not None and str(art).strip() in AUTO:
        old = X.value(sp, f"H{r}")
        if old not in (None, ""):
            try: sp.num(f"K{r}", float(old) if "." in str(old) else int(old), st=s_h)
            except ValueError: sp.text(f"K{r}", str(old), st=s_h)
        sp.formula(f"H{r}", f'IFERROR(INDEX({C}$N$4:$N$500,MATCH(E{r}&"",{C}$H$4:$H$500,0))+0,0)', st=s_h)
        auto_rows.append(r)
last_k = LASTSP
# L:R подсчёт → S «есть в спецификации?»
sp.text("S2", "Есть в спецификации?", st=sp.style("R2"))
for r in list(range(4, 24)) + list(range(27, 37)) + list(range(39, 49)) + [62, 63]:
    sp.formula(f"S{r}", f'IF(OR(L{r}="",Q{r}=""),IF(P{r}="нет в каталоге","добавьте в Каталог",""),IF(COUNTIF($E$4:$E${LASTSP},Q{r})>0,"есть","НЕТ — добавьте строку"))', st=sp.style("R4"))
p = max_prio(sp)
add_cf(sp, f'<conditionalFormatting sqref="K4:K{LASTSP}">'
           f'<cfRule type="expression" dxfId="{D_RED}" priority="{p+1}"><formula>AND($K4&lt;&gt;"",IFERROR(--$K4,-1)&lt;&gt;$H4)</formula></cfRule>'
           f'</conditionalFormatting>'
           f'<conditionalFormatting sqref="S4:S63">'
           f'<cfRule type="expression" dxfId="{D_RED}" priority="{p+2}"><formula>OR(LEFT($S4,3)="НЕТ",LEFT($S4,3)="доб")</formula></cfRule>'
           f'</conditionalFormatting>')
ensure_row(sp, 65); ensure_row(sp, 66)
sp.text("L65", f"Автоматы, аппараты, модули и клеммы 3L/QUATTRO в столбце H теперь считаются по однолинейке (по артикулу из E через Каталог) — {len(auto_rows)} строк. "
               "Корпус, шины, провод, изоляторы и прочее — вручную, как раньше.", st=sp.style("L64"))
sp.text("L66", "K — сколько было вписано вручную до v3.6 (красный = отличается от однолинейки). Столбец S: «НЕТ» — позиция есть в однолинейке, но строки в спецификации нет.", st=sp.style("L64"))
set_cols(sp, "K", "K", width=12)
set_cols(sp, "L", "L", width=34)
set_cols(sp, "P", "P", width=40)
set_cols(sp, "Q", "R", width=14)
set_cols(sp, "S", "S", width=20)
sp.save()
print("spec auto rows", auto_rows)

# ================================================================= 3. ТКЗ: имя без INDIRECT (не пересчитывается на каждую правку)
wbp = WD + "/xl/workbook.xml"; wb = open(wbp, encoding="utf-8").read()
wb2 = re.sub(r'(<definedName name="Кабельный_журнал">)[^<]*(</definedName>)', r"\1Кабельный_журнал!$P$5:$AB$1000\2", wb)
assert wb2 != wb
open(wbp, "w", encoding="utf-8").write(wb2)

# ================================================================= 4. Коды потребителей → УГО: одна таблица в «Справочной» (AU:AX)
spr = X.Sheet(X.sheet_path(WD, "Справочная"))
ref = X.Sheet(X.sheet_path(WD, "Справочник_схемы"))
UGO = {}
for r in range(4, 41):
    code = X.value(ref, f"F{r}")
    if code: UGO[code] = (X.value(ref, f"G{r}"), X.value(ref, f"H{r}"), X.value(ref, f"I{r}"))
sh_hd = spr.style("P3"); sh_v = spr.style("Q4"); s_in = spr.style("R4") or sh_v
spr.text("AU2", "УГО на однолинейке (по коду)", st=spr.style("AQ2") or sh_hd)
for c, t in zip(["AU", "AV", "AW", "AX"], ["Код (из P)", "УГО", "30 мА", "Мин. автомат, А"]):
    spr.text(f"{c}3", t, st=sh_hd)
codesP = {}
for r in range(4, 52):
    codesP[r] = X.value(spr, f"P{r}")
    spr.formula(f"AU{r}", f'IF(P{r}="","",P{r})', st=sh_v)
    u = UGO.get(codesP[r])
    if u:
        if u[0]: spr.text(f"AV{r}", str(u[0]), st=s_in)
        if u[1]: spr.text(f"AW{r}", str(u[1]), st=s_in)
        if u[2] not in (None, ""): spr.num(f"AX{r}", u[2], st=s_in)
spr.text("AU53", "Доп. коды (нет в списке P — латиница/кириллица и т.п.)", st=sh_hd)
extra = [c for c in UGO if c not in codesP.values()]
for i, c in enumerate(extra):
    r = 54 + i
    spr.text(f"AU{r}", c, st=s_in); u = UGO[c]
    if u[0]: spr.text(f"AV{r}", str(u[0]), st=s_in)
    if u[1]: spr.text(f"AW{r}", str(u[1]), st=s_in)
    if u[2] not in (None, ""): spr.num(f"AX{r}", u[2], st=s_in)
spr.text("AU75", "УГО: SOCKET, LAMP, MOTOR, HEAT, VALVE, SERVO, BOX. Пустое УГО (жёлтое) — линия без значка.", st=sh_v)
# список УГО и да/нет
dv = ('<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="AV4:AV72"><formula1>"SOCKET,LAMP,MOTOR,HEAT,VALVE,SERVO,BOX"</formula1></dataValidation>'
      '<dataValidation type="list" allowBlank="1" showErrorMessage="0" sqref="AW4:AW72"><formula1>"да,нет"</formula1></dataValidation>')
mm = re.search(r'<dataValidations count="(\d+)">', spr.s)
if mm:
    spr.s = spr.s.replace(mm.group(0), f'<dataValidations count="{int(mm.group(1)) + 2}">', 1)
    spr.s = spr.s.replace("</dataValidations>", dv + "</dataValidations>", 1)
else:
    spr.s = spr.s.replace("<pageMargins", f'<dataValidations count="2">{dv}</dataValidations><pageMargins', 1)
p = max_prio(spr)
add_cf(spr, f'<conditionalFormatting sqref="AV4:AV51"><cfRule type="expression" dxfId="{D_RED}" priority="{p+1}">'
            f'<formula>AND($AU4&lt;&gt;"",$AV4="")</formula></cfRule></conditionalFormatting>')
set_cols(spr, "AT", "AT", width=3); set_cols(spr, "AU", "AU", width=12); set_cols(spr, "AV", "AV", width=10)
set_cols(spr, "AW", "AW", width=7); set_cols(spr, "AX", "AX", width=10)
spr.save()
# имя УГО_спр → новая таблица
wb = open(wbp, encoding="utf-8").read()
wb = re.sub(r'(<definedName name="УГО_спр">)[^<]*(</definedName>)', r"\1Справочная!$AU$4:$AX$72\2", wb)
open(wbp, "w", encoding="utf-8").write(wb)
# Справочник_схемы: старую таблицу F:I очистить, оставить указатель
for r in range(3, 41):
    for c in "FGHI":
        ref.clear(f"{c}{r}")
ref.text("F1", "Коды потребителей → УГО, 30 мА, мин. автомат — теперь на листе «Справочная», столбцы AU:AX", st=ref.style("F1"))
ref.s = re.sub(r'<dataValidations[^>]*>.*?</dataValidations>', "", ref.s, flags=re.S)
ref.save()
print("UGO moved", len(UGO), "extra", extra)

# ================================================================= 5. Однолинейка: свернуть W:X и Z:AD
sh = X.Sheet(X.sheet_path(WD, "Однолинейка"))
set_cols(sh, "W", "X", extra={"hidden": "1", "outlineLevel": "1"})
set_cols(sh, "Z", "AD", extra={"hidden": "1", "outlineLevel": "1"})
if "<outlinePr" not in sh.s:
    if "<sheetPr/>" in sh.s:
        sh.s = sh.s.replace("<sheetPr/>", '<sheetPr><outlinePr summaryRight="1"/></sheetPr>', 1)
    elif re.search(r"<sheetPr[^>]*/>", sh.s):
        sh.s = re.sub(r"<sheetPr([^>]*)/>", r'<sheetPr\1><outlinePr summaryRight="1"/></sheetPr>', sh.s, 1)
sh.save()

# ================================================================= 6. Краткое руководство
g = X.Sheet(X.sheet_path(WD, "Краткое руководство"))
sB = g.style("B11"); sC = g.style("C11"); sE = g.style("E9"); sF = g.style("F9")
g.text("E10", "v3.5 (схема)", st=sE)
g.text("F10", "Лист «Однолинейка» + программа AutoCAD SX_Schema (SXDRAW/SXUPDATE). Каталог оборудования с выпадающими списками, "
              "светофор (AE), защита паролем sx, лист «Бирки», подсчёт и артикулы в «Спецификации».", st=sF)
g.text("E11", "v3.6 (схема)", st=sE)
g.text("F11", "В «Спецификации» автоматы, аппараты, модули и клеммы 3L считаются по однолинейке (по артикулу), остальное вручную. Каталог показывает «в проекте, шт.». "
              "ТКЗ больше не пересчитывается при каждой правке. Коды → УГО перенесены в «Справочную» AU:AX. Служебные столбцы однолинейки свёрнуты.", st=sF)
RULES = [
    ("11", "ОДНОЛИНЕЙКА: вводить только в белые столбцы D, E, F, I–L, N, P, Y. Остальное считается само (лист защищён, пароль sx)."),
    ("12", "Столбец AE «Статус»: красный — ошибка (дубль клеммы, нет фазы, автомат/канал), жёлтый — проверить (нет мощности, 5 жил на 1 фазе, розетка без 30 мА, потери > 4%)."),
    ("13", "Канал (N): пусто — следующий по порядку, № — вручную, «=» — параллельно предыдущей, «+» — новый канал, «р» — резерв, «б»/«к» — без канала модуля."),
    ("14", "Новое оборудование — строкой вниз на листе «Каталог» (модель, артикул, полюса, номинал, каналы). Сразу появится в списках и в спецификации."),
    ("15", "AutoCAD: APPLOAD → SX_Schema.lsp. SXDRAW — построить (окно), SXUPDATE — перестроить без окна, SXCLEAR — удалить схему."),
    ("16", "Новый код потребителя: строка в «Справочной» (P…Y) + в той же строке УГО/30 мА/мин. автомат (AU:AX). Жёлтое УГО = значок не задан."),
    ("17", "Лист «В Акад» — старый способ (Data Link), оставлен для коллег. Новый способ — SXDRAW по листу «Однолинейка»."),
]
for i, (n, t) in enumerate(RULES):
    r = 12 + i
    g.text(f"B{r}", n, st=sB); g.text(f"C{r}", t, st=sC)
g.save()

# ---------------------------------------------------------------- сборка
if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n in order:
        p_ = os.path.join(WD, n)
        if os.path.exists(p_): z.write(p_, n)
print("ok; spec K to row", last_k)
