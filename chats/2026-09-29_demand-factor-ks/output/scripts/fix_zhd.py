"""Точечные правки готовых файлов: импорт от вышестоящего — пусто/0 считается «нет данных»; вводы в однолинейке — «вручную»."""
import sys, os, re, glob, html
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import patch7 as P
from patch7 import X, ct, ce, bulk
import openpyxl
def fix(path, vvod=True):
    WD = path + "_w"; order = P.unpack(path, WD); X.load_sst(WD)
    ps = X.Sheet(X.sheet_path(WD, "Питание_щита")); n = 0
    def rep(m):
        nonlocal n
        cell = m.group(0); new = cell
        for k in range(3):
            ir = 6 + k
            new = (new.replace(f"ISNUMBER($C${ir})", f"N($C${ir})&lt;&gt;0").replace(f"ISNUMBER($D${ir})", f"N($D${ir})&gt;0")
                      .replace(f"ISNUMBER($E${ir})", f"N($E${ir})&gt;0").replace(f"ISNUMBER($F${ir})", f"N($F${ir})&gt;0"))
        if new != cell: n += 1
        return new
    ps.s = re.sub(r'<c r="[CDE](?:33|34|36|59)"[^>]*?(?:/>|>.*?</c>)', rep, ps.s, flags=re.S); ps.save()
    nv = 0
    if vvod:
        rz = openpyxl.load_workbook(path, read_only=True)["Разбивка_по_щитам"]
        rows = [i for i, row in enumerate(rz.iter_rows(min_row=1, max_row=1000, min_col=8, max_col=8, values_only=True), 1)
                if row[0] and str(row[0]).startswith("Ввод")]
        ol = X.Sheet(X.sheet_path(WD, "Однолинейка")); c = {}
        for r in rows:
            c[f"E{r}"] = ce(f"E{r}", ol.style(f"E{r}")); c[f"Y{r}"] = ct(f"Y{r}", "вручную", ol.style(f"Y{r}")); nv += 1
        bulk(ol, c); ol.save()
    P.pack(WD, order, path)
    import shutil; shutil.rmtree(WD)
    print(os.path.basename(path), "формул", n, "вводов", nv)
for f in sys.argv[1:]:
    fix(f, vvod="_" in os.path.basename(f))
