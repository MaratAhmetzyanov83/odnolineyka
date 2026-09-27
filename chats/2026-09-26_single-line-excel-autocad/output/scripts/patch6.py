"""v3.8: «Справочник_схемы» → «Справочная»; спецификация по варианту В (черновик + ваша значениями);
порядок/цвета/скрытие листов; сборка двух файлов: пример (с данными ЖК «Остров») и чистый шаблон."""
import sys, os, re, shutil, zipfile, html, copy
SP = "/tmp/claude-0/-home-claude/5e645175-0e52-5350-b586-c0cb5927e030/scratchpad/"
sys.path.insert(0, SP + "ost")
import xlsxlib as X
import openpyxl

SRC = SP + "ost/Однолинейка ЩР ЭОМ_v3.7.xlsx"
RC7 = SP + "ost/rc7/Однолинейка ЩР ЭОМ_v3.7.xlsx"      # пересчитанная копия (значения авто-спецификации)
OUT_EX = SP + "ost/v38/Однолинейка_пример_ЖК-Остров.xlsx"
OUT_TPL = SP + "ost/v38/Однолинейка_шаблон.xlsx"
WD = SP + "ost/w8"
os.makedirs(SP + "ost/v38", exist_ok=True)
shutil.rmtree(WD, ignore_errors=True)
with zipfile.ZipFile(SRC) as z:
    z.extractall(WD); order = z.namelist()
X.load_sst(WD)
esc = lambda s: html.escape(str(s), quote=False)
O = "Однолинейка!"


def rd(p): return open(p, encoding="utf-8").read()
def wr(p, s): open(p, "w", encoding="utf-8").write(s)


def bulk(sh_, cells, create=True):
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
    if byrow and create:   # недостающие строки
        for r in sorted(byrow):
            cells_ = byrow[r]
            row = f'<row r="{r}">' + "".join(cells_[c] for c in sorted(cells_, key=X.col2n)) + "</row>"
            rows = [(int(m.group(1)), m.start()) for m in re.finditer(r'<row r="(\d+)"', sh_.s)]
            pos = next((p for n, p in rows if n > r), None)
            if pos is None: pos = sh_.s.find("</sheetData>")
            sh_.s = sh_.s[:pos] + row + sh_.s[pos:]
    elif byrow:
        raise RuntimeError(("rows missing", sorted(byrow)[:5]))

def _s(s): return f' s="{s}"' if s not in (None, "None", "") else ""
def ct(ref, v, s): return f'<c r="{ref}"{_s(s)} t="inlineStr"><is><t xml:space="preserve">{esc(v)}</t></is></c>'
def cn(ref, v, s): return f'<c r="{ref}"{_s(s)}><v>{v}</v></c>'
def cf(ref, f, s): return f'<c r="{ref}"{_s(s)}><f>{esc(f)}</f></c>'
def ce(ref, s): return f'<c r="{ref}"{_s(s)}/>'


def add_dxf(xml):
    stp = WD + "/xl/styles.xml"; st = rd(stp)
    m = re.search(r'<dxfs count="(\d+)"', st); n = int(m.group(1))
    st = st.replace(m.group(0), f'<dxfs count="{n + 1}"', 1)
    k = st.find("</dxfs>"); st = st[:k] + xml + st[k:]
    wr(stp, st); return n
D_RED = add_dxf('<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>')

wbp = WD + "/xl/workbook.xml"
rp = WD + "/xl/_rels/workbook.xml.rels"
ctp = WD + "/[Content_Types].xml"


def sheet_file(name):
    return X.sheet_path(WD, name)


# ================================================================= 1. «Справочник_схемы» → «Справочная» (BE:BN)
ref = X.Sheet(sheet_file("Справочник_схемы"))
R_T = ref.style("A1"); R_H = ref.style("A3"); R_L = ref.style("A4"); R_V = ref.style("B8"); R_IN = ref.style("B4")
spr = X.Sheet(sheet_file("Справочная"))
c = {}
c["BE2"] = ct("BE2", "НАСТРОЙКИ ОДНОЛИНЕЙКИ И СВОДКА (перенесено с листа «Справочник_схемы»)", R_T)
c["BE3"] = ct("BE3", "НАСТРОЙКИ (жёлтые — можно менять)", R_H); c["BF3"] = ce("BF3", R_H)
c["BE4"] = ct("BE4", "Мест (столбцов схемы) на лист A3", R_L); c["BF4"] = cn("BF4", 17, R_IN)
c["BE5"] = ct("BE5", "Мин. автомат перед БП / трансформатором, А", R_L); c["BF5"] = cn("BF5", 10, R_IN)
c["BE7"] = ct("BE7", "ИТОГИ ПО ОДНОЛИНЕЙКЕ", R_H); c["BF7"] = ce("BF7", R_H)
ITOG = [("Автоматов всего", f'COUNTIF({O}$E$3:$E$1000,"?*")'),
        ("  из них с 30 мА (QFD)", f'COUNTIF({O}$AJ$3:$AJ$1000,"QFD*")'),
        ("Клемм XT (номеров)", f'COUNTIF({O}$AY$3:$AY$1000,1)'),
        ("  трёхуровневых блоков", f'SUM({O}$BW$3:$BW$1000)'),
        ("  доп. проходных", f'SUM({O}$BX$3:$BX$1000)'),
        ("Модулей всего", f'COUNTIF({O}$AW$3:$AW$1000,"?*")'),
        ("Строк с ошибками (красные в AE)", f'COUNTIF({O}$BU$3:$BU$1000,2)'),
        ("Строк с предупреждениями (жёлтые в AE)", f'COUNTIF({O}$BU$3:$BU$1000,1)')]
for i, (a, f) in enumerate(ITOG):
    c[f"BE{8+i}"] = ct(f"BE{8+i}", a, R_L); c[f"BF{8+i}"] = cf(f"BF{8+i}", f, R_V)
c["BE18"] = ct("BE18", "КАК ПОДБИРАЕТСЯ АВТОМАТ («авто»)", R_H); c["BF18"] = ce("BF18", R_H)
NOTES = [
    "1. Ток группы = сумма токов линий группы (линия в нескольких строках считается один раз).",
    "2. Номинал по току — по таблице «Автомат_номиналы» (эта страница, F:G).",
    "3. Итог = наибольший из: по току, «Мин. автомат» по коду (AX: розетки 16 А, свет 10 А) и ручного автомата из «Исходных данных».",
    "4. «Макс. автомат по кабелю» — по сечению (I:J). Если итог больше — «КАБЕЛЬ НЕ ЗАЩИЩЁН».",
    "5. Если после автомата БП/трансформатор — берётся ток и мин. номинал BF5.",
    "6. Полюса: 3 фазы → 3P (4P с 30 мА), 1 фаза → 1P (2P с 30 мА).",
    "7. 30 мА: столбец F «Однолинейки» или по коду (AW).",
    "8. Свой текст в «Автомат» (например «2P С20А 30мА») — используется как есть и проверяется."]
for i, t in enumerate(NOTES): c[f"BE{19+i}"] = ct(f"BE{19+i}", t, R_L)
c["BH3"] = ct("BH3", "МОДУЛИ В ПРОЕКТЕ — СВОБОДНЫЕ КАНАЛЫ", R_H)
for col in ("BI", "BJ", "BK", "BL", "BM", "BN"): c[f"{col}3"] = ce(f"{col}3", R_H)
for col, t in zip(["BH", "BI", "BJ", "BK", "BL", "BM", "BN"], ["Щит", "Модуль", "Модель", "Каналов", "Занято (посл. №)", "Резерв «р»", "Свободно"]):
    c[f"{col}4"] = ct(f"{col}4", t, R_H)
AGm = lambda k: f'(MATCH({k},{O}$BS$1:$BS$1000,0))'
for k in range(1, 36):
    r = 4 + k
    c[f"BH{r}"] = cf(f"BH{r}", f'IFERROR(INDEX({O}$A$1:$A$1000,{AGm(k)}),"")', R_V)
    c[f"BI{r}"] = cf(f"BI{r}", f'IFERROR(INDEX({O}$AW$1:$AW$1000,{AGm(k)}),"")', R_V)
    c[f"BJ{r}"] = cf(f"BJ{r}", f'IF(BI{r}="","",INDEX({O}$AI$1:$AI$1000,{AGm(k)}))', R_V)
    c[f"BK{r}"] = cf(f"BK{r}", f'IF(BI{r}="","",IFERROR(VLOOKUP(BJ{r},Модули_спр,3,FALSE),""))', R_V)
    c[f"BL{r}"] = cf(f"BL{r}", f'IF(BI{r}="","",_xlfn.MAXIFS({O}$O$3:$O$1000,{O}$AH$3:$AH$1000,BI{r},{O}$A$3:$A$1000,BH{r}))', R_V)
    c[f"BM{r}"] = cf(f"BM{r}", f'IF(BI{r}="","",COUNTIFS({O}$AH$3:$AH$1000,BI{r},{O}$A$3:$A$1000,BH{r},{O}$N$3:$N$1000,"р"))', R_V)
    c[f"BN{r}"] = cf(f"BN{r}", f'IF(OR(BI{r}="",BK{r}=""),"",BK{r}-BL{r})', R_V)
c["BH41"] = ct("BH41", "Свободных каналов всего", R_H)
for col in ("BI", "BJ", "BK", "BL", "BM"): c[f"{col}41"] = ce(f"{col}41", R_H)
c["BN41"] = cf("BN41", "SUM(BN5:BN39)", R_H)
bulk(spr, c)
# ширины
m = re.search(r"<cols>(.*?)</cols>", spr.s, flags=re.S)
cols = re.findall(r"<col [^>]*/>", m.group(1))
W = {"BD": 3, "BE": 48, "BF": 10, "BG": 3, "BH": 8, "BI": 9, "BJ": 14, "BK": 9, "BL": 11, "BM": 10, "BN": 10}
cols = [x for x in cols if not any(X.col2n(a) <= int(re.search(r'max="(\d+)"', x).group(1)) and int(re.search(r'min="(\d+)"', x).group(1)) <= X.col2n(a) for a in W)]
cols += [f'<col min="{X.col2n(a)}" max="{X.col2n(a)}" width="{w}" customWidth="1"/>' for a, w in W.items()]
cols.sort(key=lambda x: int(re.search(r'min="(\d+)"', x).group(1)))
spr.s = spr.s[:m.start()] + "<cols>" + "".join(cols) + "</cols>" + spr.s[m.end():]
spr.save()

# шапка «Однолинейки»: сводка светофора в AE1
sh = X.Sheet(sheet_file("Однолинейка"))
bulk(sh, {"AE1": cf("AE1", '"Ошибок: "&COUNTIF($BU$3:$BU$1000,2)&"   Предупреждений: "&COUNTIF($BU$3:$BU$1000,1)', sh.style("Y1") or sh.style("W1"))})
sh.save()

# удалить лист «Справочник_схемы»
wb = rd(wbp)
rid = re.search(r'<sheet name="Справочник_схемы"[^>]*r:id="(rId\d+)"', wb).group(1)
rels = rd(rp); tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels).group(1)
os.remove(WD + "/xl/" + tgt)
order.remove("xl/" + tgt)
wr(rp, re.sub(r'<Relationship Id="%s"[^>]*/>' % rid, "", rels))
wr(ctp, rd(ctp).replace(f'<Override PartName="/xl/{tgt}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>', ""))
names_before = re.findall(r'<sheet name="([^"]+)"', wb)
del_idx = names_before.index("Справочник_схемы")
wb = re.sub(r'<sheet name="Справочник_схемы"[^>]*/>', "", wb)
def _dec(mm):
    i = int(mm.group(1)); assert i != del_idx
    return f'localSheetId="{i - 1 if i > del_idx else i}"'
wb = re.sub(r'localSheetId="(\d+)"', _dec, wb)
wb = re.sub(r'activeTab="(\d+)"', lambda mm: f'activeTab="{int(mm.group(1)) - (1 if int(mm.group(1)) > del_idx else 0)}"', wb)
wr(wbp, wb)

# ================================================================= 2. Спецификация: черновик (авто) + ваша (значения)
wb = rd(wbp)
wb = wb.replace('<sheet name="Спецификация" ', '<sheet name="Спецификация_авто" ', 1)
# ссылки других листов на подсчёт спецификации
for nm in ("Каталог",):
    p = sheet_file(nm); s = rd(p)
    s = s.replace("Спецификация!", "Спецификация_авто!")
    wr(p, s)
wr(wbp, wb)
auto = X.Sheet(sheet_file("Спецификация_авто"))
bulk(auto, {"C1": cf("C1", ('IF(Каталог!$R$28="","ЧЕРНОВИК — считается сам, не править. Перенос в «Спецификацию»: выделите B4:J160 → Копировать → '
                            'лист «Спецификация», ячейка B4 → Вставить → Значения (или кнопка/макрос «ОбновитьСпецификацию»).",'
                            '"ВНИМАНИЕ: в однолинейке есть позиции, которых нет в Каталоге: "&Каталог!$R$28&IF(Каталог!$R$29="","",", "&Каталог!$R$29)'
                            '&IF(Каталог!$R$30="","",", "&Каталог!$R$30)&" — добавьте их в Каталог.")'), auto.style("C1"))})
auto.save()

# ваша «Спецификация»: из скрытого листа «Спецификация_до_v3.7» (старый вид, только значения B:J)
wb = rd(wbp)
rid_old = re.search(r'<sheet name="Спецификация_до_v3.7"[^>]*r:id="(rId\d+)"', wb).group(1)
man_path = WD + "/xl/" + re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid_old, rd(rp)).group(1)
wb = re.sub(r'<sheet name="Спецификация_до_v3.7"([^>]*) state="hidden"', r'<sheet name="Спецификация"\1', wb)
wr(wbp, wb)
man = X.Sheet(man_path)
st_row = {col: (man.style(f"{col}20") or "0") for col in "BCDEFGHIJ"}
# значения авто-спецификации (пересчитаны LibreOffice)
vwb = openpyxl.load_workbook(RC7, data_only=True)
va = vwb["Спецификация"]
EX_ROWS = []
for r in range(4, 161):
    vals = [va[f"{col}{r}"].value for col in "BCDEFGHIJ"]
    if vals[0] in (None, ""): break
    EX_ROWS.append(vals)
print("строк в спецификации (пример):", len(EX_ROWS))
cells = {}
for r in range(4, 161):
    rec = EX_ROWS[r - 4] if r - 4 < len(EX_ROWS) else None
    for j, col in enumerate("BCDEFGHIJ"):
        ref_ = f"{col}{r}"; s_ = st_row[col]
        v = rec[j] if rec else None
        if v in (None, ""): cells[ref_] = ce(ref_, s_)
        elif isinstance(v, (int, float)): cells[ref_] = cn(ref_, v, s_)
        else: cells[ref_] = ct(ref_, v, s_)
bulk(man, cells)
man.s = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:J160"/>', man.s)
ps = re.search(r"<pageSetup [^>]*/>", auto.s)
if ps and "<pageSetup" not in man.s:
    tag = re.sub(r'\sr:id="[^"]*"', "", ps.group(0))
    man.s = re.sub(r"(<pageMargins [^>]*/>)", r"\1" + tag, man.s, count=1)
man.save()
MAN_PATH = man_path

# ================================================================= 3. Порядок листов, цвета, скрытие
ORDER = ["Краткое руководство", "Справочная", "Каталог", "Исходные данные", "Разбивка_по_щитам", "Однолинейка",
         "Нагрузка_щитов", "ТКЗ", "Кабельный_журнал", "Спецификация", "Спецификация_авто", "Бирки",
         "В Акад", "Диапазоны", "Справочная_потребители"]
HIDE = {"В Акад", "Диапазоны", "Справочная_потребители"}
COLOR = {"Справочная": "FF4472C4", "Каталог": "FF4472C4",
         "Исходные данные": "FF70AD47", "Разбивка_по_щитам": "FF70AD47", "Однолинейка": "FF70AD47",
         "Нагрузка_щитов": "FFED7D31", "ТКЗ": "FFED7D31", "Кабельный_журнал": "FFED7D31", "Спецификация": "FFED7D31", "Бирки": "FFED7D31",
         "Спецификация_авто": "FFA6A6A6", "В Акад": "FFA6A6A6", "Диапазоны": "FFA6A6A6", "Справочная_потребители": "FFA6A6A6"}
wb = rd(wbp)
tags = {re.search(r'name="([^"]+)"', t).group(1): t for t in re.findall(r"<sheet [^>]*/>", wb)}
old_names = [re.search(r'name="([^"]+)"', t).group(1) for t in re.findall(r"<sheet [^>]*/>", wb)]
assert set(old_names) == set(ORDER), (set(old_names) ^ set(ORDER))
newtags = []
for n in ORDER:
    t = re.sub(r'\sstate="[^"]*"', "", tags[n])
    if n in HIDE: t = t.replace("/>", ' state="hidden"/>')
    newtags.append(t)
wb = re.sub(r"<sheets>.*?</sheets>", "<sheets>" + "".join(newtags) + "</sheets>", wb, flags=re.S)
# localSheetId по имени
def remap(mm):
    old = int(mm.group(1)); return f'localSheetId="{ORDER.index(old_names[old])}"'
wb = re.sub(r'localSheetId="(\d+)"', remap, wb)
wb = re.sub(r'activeTab="\d+"', f'activeTab="{ORDER.index("Однолинейка")}"', wb)
wb = re.sub(r'\sfirstSheet="\d+"', "", wb)
# имена настроек → Справочная
def setname(name, val):
    global wb
    pat = r'(<definedName name="%s"[^>]*>)[^<]*(</definedName>)' % re.escape(name)
    assert re.search(pat, wb), name
    wb = re.sub(pat, lambda mm: mm.group(1) + esc(val) + mm.group(2), wb)
setname("Столбцов_на_лист", "Справочная!$BF$4")
setname("Мин_автомат_БП", "Справочная!$BF$5")
wr(wbp, wb)
for n in ORDER:
    p = sheet_file(n); s = rd(p)
    s = re.sub(r'\stabSelected="1"', "", s)
    if n == "Однолинейка":
        s = re.sub(r"<sheetView ", '<sheetView tabSelected="1" ', s, count=1)
    col = COLOR.get(n)
    s = re.sub(r"<tabColor [^>]*/>", "", s)
    if col:
        tc = f'<tabColor rgb="{col}"/>'
        if re.search(r"<sheetPr\b[^>]*/>", s):
            s = re.sub(r"<sheetPr\b([^>]*)/>", r"<sheetPr\1>" + tc + "</sheetPr>", s, count=1)
        elif re.search(r"<sheetPr\b[^>]*>", s):
            s = re.sub(r"(<sheetPr\b[^>]*>)", r"\1" + tc, s, count=1)
        else:
            s = re.sub(r"(<worksheet\b[^>]*>)", r"\1<sheetPr>" + tc + "</sheetPr>", s, count=1)
    wr(p, s)

# ================================================================= 4. Краткое руководство
gd = X.Sheet(sheet_file("Краткое руководство"))
sE = gd.style("E9") or gd.style("E12"); sF = gd.style("F9") or gd.style("F12"); sB = gd.style("B12"); sC = gd.style("C12")
cells = {
    "E13": ct("E13", "v3.8 (схема)", sE),
    "F13": ct("F13", "Лист «Справочник_схемы» убран: настройки, итоги и свободные каналы — «Справочная» BE:BN; сводка ошибок — «Однолинейка» AE1. "
                     "Спецификация: «Спецификация» — ваша (значения, правьте как хотите), «Спецификация_авто» — черновик (формулы). "
                     "Листы по цветам: синий — справочники, зелёный — ввод, оранжевый — результаты, серый — служебные (часть скрыта).", sF),
}
RULES = {
    6: "Листы: синие — справочники (Справочная, Каталог), зелёные — ввод (Исходные данные > Разбивка по щитам > Однолинейка), оранжевые — результаты, серые — служебные. Скрытые (В Акад, Диапазоны, Справочная_потребители): правая кнопка по ярлыку → Показать.",
    15: "СПЕЦИФИКАЦИЯ: «Спецификация_авто» считается сама (Каталог: N — по однолинейке, O — вручную/запас). Когда схема готова — выделить «Спецификация_авто» B4:J160 → Копировать → «Спецификация» B4 → Вставить → Значения. Дальше «Спецификацию» правьте как обычно.",
    17: "Новый код потребителя: строка в «Справочной» (P…Y) + в той же строке УГО (список), 30 мА, мин. автомат, клемм 3L на линию (AU:AY). Настройки однолинейки (мест на лист, мин. автомат перед БП) — «Справочная» BF4:BF5.",
}
for r, t in RULES.items():
    cells[f"B{r}"] = ct(f"B{r}", str(r - 1), sB); cells[f"C{r}"] = ct(f"C{r}", t, sC)
bulk(gd, cells)
gd.save()

# проверка: нигде не осталось ссылок на удалённый лист
for n in ORDER:
    s = rd(sheet_file(n))
    assert "Справочник_схемы!" not in s, n
assert "Справочник_схемы!" not in rd(wbp)


def pack(out):
    if os.path.exists(out): os.remove(out)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n_ in order:
            p_ = os.path.join(WD, n_)
            if os.path.exists(p_): z.write(p_, n_)

pack(OUT_EX)
print("пример готов")

# ================================================================= 5. Шаблон: очистить ввод объекта
def clear_consts(name, cols, r1, r2):
    sh_ = X.Sheet(sheet_file(name)); n = [0]
    colset = set(cols)
    def fixrow(mm):
        r = int(mm.group(1))
        if not (r1 <= r <= r2): return mm.group(0)
        def fixc(cm):
            if cm.group(1) not in colset: return cm.group(0)
            cell = cm.group(0)
            if "<f" in cell: return cell
            if "<v>" not in cell and "<is>" not in cell: return cell
            st_ = re.search(r'\ss="(\d+)"', cell)
            n[0] += 1
            return f'<c r="{cm.group(1)}{r}"' + (f' s="{st_.group(1)}"' if st_ else "") + "/>"
        return re.sub(r'<c r="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)', fixc, mm.group(0), flags=re.S)
    sh_.s = re.sub(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', fixrow, sh_.s, flags=re.S)
    sh_.save(); print(f"  очищено {name}: {n[0]}")

clear_consts("Исходные данные", ["C", "D", "E", "F", "G", "H", "I", "O", "Q", "R", "S", "AU", "AV", "AW", "AX", "AY"], 3, 1000)
clear_consts("Разбивка_по_щитам", ["F", "H"], 3, 1000)
clear_consts("Однолинейка", ["D", "E", "F", "I", "J", "K", "L", "N", "P", "Y"], 3, 1000)
clear_consts("Нагрузка_щитов", ["E"], 3, 1000)
clear_consts("Каталог", ["O"], 4, 500)
clear_consts("Спецификация", list("BCDEFGHIJ"), 4, 160)
gd = X.Sheet(sheet_file("Краткое руководство"))
bulk(gd, {"C1": ct("C1", "ШАБЛОН однолинейки v3.8 — копируйте на каждый объект/щит (имя файла без номера версии). Пример заполнения — «Однолинейка_пример_ЖК-Остров.xlsx».", gd.style("C1") or gd.style("C2"))})
gd.save()
# убрать сохранённые результаты формул — шаблон пересчитается сам при открытии
for n in ORDER:
    p_ = sheet_file(n); s_ = rd(p_)
    def strip(cm):
        cell = cm.group(0)
        if "<f" not in cell: return cell
        cell = re.sub(r"<v>.*?</v>", "", cell, flags=re.S)
        cell = re.sub(r'\st="(?:str|e|b|n)"', "", cell, count=1)
        return cell
    s_ = re.sub(r'<c r="[A-Z]+\d+"[^>]*?>.*?</c>', strip, s_, flags=re.S)
    wr(p_, s_)
pack(OUT_TPL)
print("шаблон готов")
