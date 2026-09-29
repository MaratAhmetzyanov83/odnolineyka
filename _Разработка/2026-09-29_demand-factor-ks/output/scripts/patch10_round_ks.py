"""v3.11: округление выводных величин таблицы нагрузок щита (лист «В Акад», блоки P:U строки 2-6, 9-13, 16-20).

AutoCAD берёт таблицу (Руст, Рр, Iр, L1р-L3р, Cos(ф), Кс) через Data Link с «В Акад», а «В Акад» ссылается на «Нагрузка_щитов»
V5:V9 / V14:V18 / V23:V27, Y5:Y7 и т.д.  Кс (V9/V18/V27) и cos φ (V8/V17/V26) в «В Акад» шли в формате «Общий»
(0,902854638).  Теперь выводные ячейки = ROUND(<источник>;2) (текст/ошибка — как есть: IF(ISNUMBER(x);ROUND(x;2);x)) + числовой формат 0,00.  Расчётные ячейки «Нагрузка_щитов»
не трогаем.  Правка на уровне XML (openpyxl такие книги ломает); общие формулы не затрагиваются, проверка check_shared.

    py patch10_round_ks.py <книга.xlsx> [<книга.xlsx> ...]     # правит на месте
"""
import sys, re, zipfile, html, os, shutil

F_RX = re.compile(r'<f(\s[^>]*?)?(?:/>|>([^<]*)</f>)')
def fattrs(t): return dict(re.findall(r'(\w+)="([^"]*)"', t or ""))

def check_shared(name, s):
    masters, deps = set(), []
    for fm in F_RX.finditer(s):
        at = fattrs(fm.group(1))
        if at.get("t") != "shared": continue
        if "ref" in at: masters.add(at["si"])
        else: deps.append(at.get("si"))
    orphan = [x for x in deps if x not in masters]
    assert not orphan, (name, "общие формулы без мастера", orphan[:5])

def sheet_path(z, title):
    wb = z.read("xl/workbook.xml").decode("utf-8"); rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rm = {re.search(r'Id="([^"]*)"', t).group(1): re.search(r'Target="([^"]*)"', t).group(1)
          for t in re.findall(r"<Relationship [^>]*>", rels)}
    for n, r in re.findall(r'<sheet name="([^"]*)" sheetId="\d+"[^>]*r:id="(rId\d+)"', wb):
        if html.unescape(n) == title: return "xl/" + rm[r].lstrip("/").replace("xl/", "")
    raise KeyError(title)

BLOCKS = (2, 9, 16)          # первая строка блока (Руст/L1р); +3 = cos φ, +4 = Кс
NUM = [(0, "Q"), (0, "T"), (1, "Q"), (1, "T"), (2, "Q"), (2, "T")]   # Руст, L1р / Рр, L2р / Iр, L3р

def patch(path):
    with zipfile.ZipFile(path) as z:
        infos = z.infolist(); data = {i.filename: z.read(i.filename) for i in infos}
        sp = sheet_path(z, "В Акад")
    s = data[sp].decode("utf-8"); st = data["xl/styles.xml"].decode("utf-8")
    m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', st, re.S)
    xfs = re.findall(r"<xf [^>]*?(?:/>|>.*?</xf>)", m.group(2), re.S)
    assert len(xfs) == int(m.group(1))
    added = []
    def with_fmt2(sid):
        """стиль sid с numFmtId=2 (0,00): существующий или новый xf в конце cellXfs"""
        x = xfs[sid]
        if re.search(r'numFmtId="2"', x): return sid
        y = re.sub(r'numFmtId="\d+"', 'numFmtId="2"', x, count=1)
        if "applyNumberFormat" not in y: y = y.replace("<xf ", '<xf applyNumberFormat="1" ', 1)
        key = lambda e: (sorted(re.findall(r'(\w+)="([^"]*)"', e.split(">")[0])), e[e.find(">"):])
        for i, e in enumerate(xfs + added):
            if key(e) == key(y): return i
        added.append(y); return len(xfs) + len(added) - 1
    changed = []
    def cell(ref, kind):
        nonlocal s
        mm = re.search(r'<c r="%s"([^>]*?)>(<f>([^<]*)</f>)(<v>([^<]*)</v>)?</c>' % ref, s)
        assert mm, ("нет ячейки/формулы", ref)
        attrs, f, v = mm.group(1), html.unescape(mm.group(3)), mm.group(5)
        if "ROUND(" in f: return                         # уже сделано
        assert re.fullmatch(r"Нагрузка_щитов![A-Z]+\d+", f), (ref, f)
        sid = int(re.search(r' s="(\d+)"', attrs).group(1))
        attrs2 = re.sub(r' s="\d+"', f' s="{with_fmt2(sid)}"', attrs)
        nv = ""
        if v is not None:
            if 't="' in attrs2:                                 # ошибка/текст — кэш оставляем
                nv = f"<v>{v}</v>"
            else:
                r_ = repr(round(float(v), 2))
                nv = "<v>%s</v>" % (r_[:-2] if r_.endswith(".0") else r_)
        new = f'<c r="{ref}"{attrs2}><f>IF(ISNUMBER({f}),ROUND({f},2),{f})</f>{nv}</c>'
        s = s.replace(mm.group(0), new, 1); changed.append((ref, f, v))
    for b in BLOCKS:
        for k in range(3):
            for col in "QT": cell(f"{col}{b + k}", "num")
        cell(f"Q{b + 3}", "cos"); cell(f"Q{b + 4}", "ks")
    if added:
        xml = "".join(added)
        st = st.replace(m.group(0), f'<cellXfs count="{len(xfs) + len(added)}">{m.group(2)}{xml}</cellXfs>', 1)
    check_shared(sp, s)
    data[sp] = s.encode("utf-8"); data["xl/styles.xml"] = st.encode("utf-8")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w") as zo:
        for i in infos:
            zi = zipfile.ZipInfo(i.filename, i.date_time); zi.compress_type = i.compress_type
            zi.external_attr = i.external_attr; zi.create_system = i.create_system
            zo.writestr(zi, data[i.filename])
    with zipfile.ZipFile(tmp) as z:
        for n in z.namelist():
            if n.startswith("xl/worksheets/sheet"): check_shared(n, z.read(n).decode("utf-8"))
    os.replace(tmp, path)
    return changed, len(added)

if __name__ == "__main__":
    for p in sys.argv[1:]:
        ch, na = patch(p)
        print(os.path.basename(p), "ячеек:", len(ch), "новых стилей:", na)
