# -*- coding: utf-8 -*-
"""snapshot.py - «снимок» ключевых результатов книги «Однолинейка» (после полного пересчёта).

    py snapshot.py <книга.xlsx> [-o снимок.json] [--backend auto|excel|libreoffice|cache]

Книга НЕ изменяется: пересчитывается временная копия (Excel COM, CalculateFull, без обновления
внешних связей; Excel закрывается всегда). Результат - нормализованный JSON (числа округлены до
6 знаков, ключи по месту/линии), пригодный для сравнения в check.py.
"""
import argparse
import json
import os
import re
import sys
from collections import Counter, OrderedDict

from openpyxl.utils import column_index_from_string as CI, get_column_letter as CL

import xlgrid

FORMAT = 1
ROUND = 6

# ---- листы и адреса (определены по книге v3.11 и документации) --------------------------------
S_GUIDE = "Краткое руководство"
S_REF = "Справочная"            # итоги: BE8:BF15, мест на лист: BF4
S_SRC = "Исходные данные"       # линии: D - номер линии, строки с данными
S_ODN = "Однолинейка"           # рисуемые строки: BD=1
S_LOAD = "Нагрузка_щитов"       # блоки «РАСЧЁТ Рр» в AE (строки 4/13/22)
S_PZ = "Расчёт_нагрузок"
S_POW = "Питание_щита"
S_SUB = "Нижестоящие_щиты"
S_SPEC = "Спецификация_авто"
S_ACAD = "В Акад"              # таблица нагрузок для AutoCAD (Data Link): P:U, 3 блока

# Однолинейка: столбец -> короткое имя поля (порядок вывода)
ODN_FIELDS = OrderedDict([
    ("C", "Лист"), ("G", "QF"), ("H", "Автомат итог"), ("M", "Модуль"), ("O", "Канал"),
    ("P", "Клемма"), ("Q", "Кабель"), ("R", "Линия"), ("T", "Р, кВт"), ("U", "Ток, А"),
    ("V", "Фаза"), ("W", "Проверка автомата"), ("X", "Проверка канала"),
    ("Z", "Ток группы, А"), ("AA", "Номинал по току"), ("AB", "Макс. автомат по кабелю"),
    ("AC", "Модель модуля"), ("AD", "Цепь"), ("AE", "Статус"),
    ("BD", "draw"), ("BE", "bus2"), ("BF", "rcd"), ("BG", "chtype"), ("BH", "dev1"),
    ("BI", "dev2"), ("BJ", "dev3"), ("BK", "sym"), ("BL", "hasline"), ("BM", "тип автомата №"),
])
# «Исходные данные»: расчётные столбцы по линиям
SRC_FIELDS = OrderedDict([
    ("E", "Мощность, кВт"), ("M", "Длина итог, м"), ("T", "Ток, А"), ("U", "Автомат"),
    ("V", "Сечение"), ("AA", "Потери, В"), ("AB", "Потери, %"), ("AD", "Кабель"),
    ("BC", "Группа Кс"), ("BF", "Допуст. ΔU, %"), ("BG", "ΔU от ВРУ, %"), ("BI", "Iкз min, А"),
    ("BJ", "Iкз/In"), ("BK", "Отключение ≤0,4 с"), ("BL", "Zпетли, Ом"), ("BM", "Iкз max, кА"),
])
# заголовки, за изменением которых следим (структура книги)
HDR_WATCH = {S_ODN: (2, ["A", "B", "C", "G", "H", "M", "O", "P", "T", "U", "V", "Z", "AA", "AB", "AC",
                          "AD", "AE", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM"]),
             S_SRC: (2, ["D", "E", "M", "T", "U", "V", "AB", "BC", "BG", "BI", "BJ", "BK", "BL", "BM"])}


# ---------------------------------------------------------------------------------------------
def is_err(v):
    return isinstance(v, str) and v.startswith("#") and (v.endswith("!") or v.endswith("?") or v in ("#Н/Д", "#N/A"))


def nv(v):
    """Нормализация значения для JSON."""
    if v is None:
        return None
    if isinstance(v, float):
        if v != v:
            return None
        r = round(v, ROUND)
        if r == 0:
            r = 0.0
        return int(r) if r == int(r) and abs(r) < 1e12 else r
    if isinstance(v, str):
        s = " ".join(v.replace("\n", " ").split())
        return s or None
    return v


class Book:
    def __init__(self, grids):
        self.g = grids

    def has(self, sheet):
        return sheet in self.g

    def val(self, sheet, addr):
        m = re.match(r"([A-Z]+)(\d+)$", addr)
        gr = self.g.get(sheet)
        return gr.get(int(m[2]), CI(m[1])) if gr else None

    def cell(self, sheet, r, col):
        gr = self.g.get(sheet)
        return gr.get(r, CI(col) if isinstance(col, str) else col) if gr else None


def uniq_key(d, key):
    """Ключ без коллизий: при повторе добавляется #2, #3..."""
    if key not in d:
        return key
    n = 2
    while "%s #%d" % (key, n) in d:
        n += 1
    return "%s #%d" % (key, n)


def put(d, key, fields):
    fields = OrderedDict((k, nv(v)) for k, v in fields.items() if nv(v) is not None)
    d[uniq_key(d, key)] = fields


# ---- разделы снимка ---------------------------------------------------------------------------
def sec_meta(b, backend, src):
    ver = None
    gr = b.g.get(S_GUIDE)
    if gr:
        for r in gr.rows():
            v = gr.get(r, CI("E"))
            if isinstance(v, str) and re.match(r"v?\d+\.\d+", v.strip()):
                ver = v.strip()
    hdr = {}
    for sheet, (row, cols) in HDR_WATCH.items():
        if b.has(sheet):
            hdr[sheet] = OrderedDict((c, nv(b.cell(sheet, row, c))) for c in cols)
    return OrderedDict([
        ("книга", os.path.basename(src)),
        ("версия шаблона", ver),
        ("бэкенд пересчёта", backend),
        ("листы", list(b.g.keys())),
        ("заголовки", hdr),
    ])


def sec_errors(b):
    """Ячейки с ошибками: лист -> столбец -> тип -> число."""
    out = OrderedDict()
    for name, gr in b.g.items():
        cols = {}
        for i, row in enumerate(gr.data):
            for j, v in enumerate(row):
                if is_err(v):
                    col = CL(gr.c0 + j)
                    cols.setdefault(col, Counter())[v] += 1
        tot = OrderedDict()
        total = 0
        for col in sorted(cols, key=CI):
            tot[col] = OrderedDict(sorted(cols[col].items()))
            total += sum(cols[col].values())
        out[name] = OrderedDict([("всего", total), ("по столбцам", tot)])
    return out


def odn_rows(b):
    """Строки «Однолинейки», которые рисует LISP: BD=1 и есть содержимое (линия/автомат/канал/клемма)."""
    gr = b.g.get(S_ODN)
    res = []
    if not gr:
        return res
    for r in range(3, gr.r0 + gr.nr):
        pos = gr.get(r, 2)
        if not isinstance(pos, (int, float)) or isinstance(pos, bool):
            continue
        res.append(r)
    return res


def lisp_nonnil(v):
    return v is not None and not is_err(v)


def sec_oneline(b):
    gr = b.g.get(S_ODN)
    out = OrderedDict()
    if not gr:
        return out
    for r in odn_rows(b):
        draw = gr.get(r, CI("BD"))
        line = gr.get(r, CI("R"))
        content = lisp_nonnil(gr.get(r, CI("AJ"))) or lisp_nonnil(gr.get(r, CI("O"))) or gr.get(r, CI("AY")) == 1
        if not (draw == 1 and (nv(line) is not None or content)):
            continue
        key = "%s | место %d" % (gr.get(r, 1), int(gr.get(r, 2)))
        put(out, key, OrderedDict((n, gr.get(r, CI(c))) for c, n in ODN_FIELDS.items()))
    return out


def sec_counters(b):
    gr = b.g.get(S_ODN)
    ref = OrderedDict()
    if b.has(S_REF):
        rg = b.g[S_REF]
        for r in range(7, 16):
            lab, v = rg.get(r, CI("BE")), rg.get(r, CI("BF"))
            if isinstance(lab, str) and v is not None and not lab.startswith("ИТОГИ"):
                ref[nv(lab)] = nv(v)
    own = OrderedDict()
    lisp = OrderedDict()
    if gr:
        per = 17
        try:
            per = int(b.val(S_REF, "BF4") or 17)
        except Exception:
            pass
        rows = odn_rows(b)
        panels = OrderedDict()
        qf, xt, mod, sheets, lines, drawn = set(), set(), set(), set(), 0, 0
        a34 = a35 = 0
        for r in rows:
            p = gr.get(r, 1)
            p = p if isinstance(p, str) else str(p)
            pn = panels.setdefault(p, {"recs": 0, "draw": 0, "sheets": set()})
            pn["recs"] += 1
            if gr.get(r, CI("BD")) != 1:
                continue
            pn["draw"] += 1
            line = nv(gr.get(r, CI("R")))
            content = lisp_nonnil(gr.get(r, CI("AJ"))) or lisp_nonnil(gr.get(r, CI("O"))) or gr.get(r, CI("AY")) == 1
            if content:
                pn["sheets"].add(int((gr.get(r, 2) - 1) // per))       # индекс листа как в LISP (0-based)
                sheets.add(gr.get(r, CI("C")))
            if line is not None:
                lines += 1
            drawn += 1
            for col, s in (("G", qf), ("P", xt), ("M", mod)):
                v = nv(gr.get(r, CI(col)))
                if v is not None and not is_err(v) and v != 0:
                    s.add(v)
        own = OrderedDict([
            ("линий (рисуемые строки с линией)", lines),
            ("строк draw=1", drawn),
            ("автоматов (различных QF)", len(qf)),
            ("номеров XT (различных)", len(xt)),
            ("модулей (различных)", len(mod)),
            ("листов схемы (с содержимым)", len(sheets)),
            ("мест на лист", per),
        ])
        for p, d in panels.items():
            lisp[p] = OrderedDict([("recs", d["recs"]), ("draw", d["draw"]),
                                   ("sheets", sorted(d["sheets"]))])
    return OrderedDict([("итоги «Справочная» BE8:BF15", ref), ("пересчёт по «Однолинейке»", own)]), lisp


def sec_status(b):
    gr = b.g.get(S_ODN)
    out = OrderedDict()
    if not gr:
        return out
    out["сводка AE1"] = nv(gr.get(1, CI("AE")))
    cnt = Counter()
    for r in odn_rows(b):
        if gr.get(r, CI("BD")) != 1:
            continue
        st = gr.get(r, CI("AE"))
        if isinstance(st, str):
            for part in st.split(";"):
                part = part.strip()
                if part:
                    part = re.sub(r"XT\d+", "XT#", part)     # «дубль клеммы XT16» -> одна категория
                    cnt[part] += 1
    out["строк по статусам (светофор, AE)"] = OrderedDict(sorted(cnt.items()))
    return out


def sec_loads(b):
    gr = b.g.get(S_LOAD)
    out = OrderedDict()
    if not gr:
        return out
    blocks = [r for r in gr.rows() if isinstance(gr.get(r, CI("AE")), str) and gr.get(r, CI("AE")).startswith("РАСЧЁТ")]
    for k, r in enumerate(blocks, 1):
        name = gr.get(r, CI("R"))
        f = OrderedDict()
        f["щит"] = "(нет)" if (name is None or is_err(name)) else name
        if f["щит"] != "(нет)":
            for i in range(1, 9):
                lab = gr.get(r + i, CI("AE"))
                if not isinstance(lab, str):
                    continue
                for col, sec in (("AF", "весь щит"), ("AG", "секция 1"), ("AH", "секция 2")):
                    f["%s | %s" % (nv(lab), sec)] = gr.get(r + i, CI(col))
            for i in range(1, 6):        # V5:V9 в блоке 1 (Руст, Рр, Iр, cos, Кс) - то, что идёт «в Акад»
                lab = gr.get(r + i, CI("U"))
                if isinstance(lab, str):
                    f["итог %s" % nv(lab).rstrip("=").strip()] = gr.get(r + i, CI("V"))
            for j, c in enumerate(("O", "P", "Q")):
                f["Р по фазе L%d, кВт" % (j + 1)] = gr.get(r + 3, CI(c))
            f["Руст итог (O)"] = gr.get(r + 1, CI("O"))
            f["КРМ"] = gr.get(r + 1, CI("AB"))
        put(out, "блок %d" % k, f)
    return out


def sec_lines_src(b):
    gr = b.g.get(S_SRC)
    out = OrderedDict()
    if not gr:
        return out
    for r in range(3, gr.r0 + gr.nr):
        line = gr.get(r, CI("D"))
        if line is None or is_err(line) or (isinstance(line, float) and line == 0):
            continue
        put(out, str(nv(line)), OrderedDict((n, gr.get(r, CI(c))) for c, n in SRC_FIELDS.items()))
    return out


def sec_pz(b):
    gr = b.g.get(S_PZ)
    out = OrderedDict()
    if not gr:
        return out
    hdr = ["", "", "Руст", "Кс", "cos φ", "tg φ", "Рр, кВт", "Qр, квар", "Sр, кВА", "Iр, А"]
    for r in range(4, 65):
        lab = gr.get(r, CI("B"))
        if not isinstance(lab, str):
            continue
        if 11 <= r <= 40 or r in (41, 42):
            f = OrderedDict((hdr[i], gr.get(r, i + 1)) for i in range(2, 10))
            put(out, "категория: " + nv(lab), f)
        elif 53 <= r <= 64 or r in (4, 5):
            put(out, "показатель: " + nv(lab), OrderedDict([("значение", gr.get(r, CI("C")))]))
    for a in ("C43", "C44", "H4", "H5"):
        v = b.val(S_PZ, a)
        if v is not None:
            put(out, "ячейка " + a, OrderedDict([("значение", v)]))
    return out


def sec_power(b):
    gr = b.g.get(S_POW)
    out = OrderedDict()
    if not gr:
        return out
    for r in range(29, 42):
        lab = gr.get(r, CI("B"))
        if isinstance(lab, str) and gr.get(r, CI("C")) is not None:
            put(out, "результат: " + nv(lab), OrderedDict([("значение", gr.get(r, CI("C")))]))
    hdr = [nv(gr.get(44, CI(c))) for c in "CDEFGH"]
    for r in range(45, 49):
        n = gr.get(r, CI("B"))
        if n is not None:
            put(out, "для вышестоящего: " + str(nv(n)),
                OrderedDict((h or c, gr.get(r, CI(c))) for h, c in zip(hdr, "CDEFGH")))
    return out


def sec_sub(b):
    gr = b.g.get(S_SUB)
    out = OrderedDict()
    if not gr:
        return out
    hdr = [(c, nv(gr.get(4, CI(c)))) for c in "DEFGHIJKLMNO"]
    for r in range(5, gr.r0 + gr.nr):
        lab = gr.get(r, CI("B"))
        if lab is None or is_err(lab):
            continue
        put(out, str(nv(lab)), OrderedDict((h or c, gr.get(r, CI(c))) for c, h in hdr))
    return out


def sec_acad(b):
    """Таблица нагрузок, которую AutoCAD берёт через Data Link: «В Акад» P:U, строки 1-25."""
    gr = b.g.get(S_ACAD)
    out = OrderedDict()
    if not gr:
        return out
    for r in range(1, 26):
        row = OrderedDict((c, gr.get(r, CI(c))) for c in "PQRSTU")
        if any(v is not None for v in row.values()):
            out["строка %d" % r] = row
    return out


def sec_spec(b):
    """Спецификация_авто: позиция (артикул | наименование) -> количество; блок подсчёта L:O."""
    gr = b.g.get(S_SPEC)
    items, counts = OrderedDict(), OrderedDict()
    warn = None
    if not gr:
        return items, counts, warn
    warn = nv(gr.get(1, CI("C")))
    for r in range(3, gr.r0 + gr.nr):
        name, code, qty = gr.get(r, CI("C")), gr.get(r, CI("E")), gr.get(r, CI("H"))
        if isinstance(name, str) and qty is not None:
            key = "%s | %s" % (nv(code) or "-", (nv(name) or "")[:60])
            items[uniq_key(items, key)] = nv(qty)
        lab = gr.get(r, CI("L"))
        if isinstance(lab, str) and gr.get(r, CI("M")) is not None:
            f = OrderedDict([("кол-во", gr.get(r, CI("M"))), ("трёхуровн.", gr.get(r, CI("N"))),
                             ("доп. проходных", gr.get(r, CI("O")))])
            put(counts, nv(lab), f)
    return items, counts, warn


# ---------------------------------------------------------------------------------------------
def build(path, backend="auto", log=print):
    grids, used, warns = xlgrid.load(path, backend, log=log)
    b = Book(grids)
    counters, lisp = sec_counters(b)
    spec, spec_cnt, spec_warn = sec_spec(b)
    snap = OrderedDict()
    snap["формат"] = FORMAT
    snap["мета"] = sec_meta(b, used, path)
    snap["ошибки в ячейках"] = sec_errors(b)
    snap["счётчики"] = counters
    snap["статус (светофор)"] = sec_status(b)
    snap["однолинейка (строки)"] = sec_oneline(b)
    snap["исходные данные (линии)"] = sec_lines_src(b)
    snap["нагрузки щита"] = sec_loads(b)
    snap["расчёт нагрузок (ПЗ)"] = sec_pz(b)
    snap["питание щита"] = sec_power(b)
    snap["нижестоящие щиты"] = sec_sub(b)
    snap["в акад (таблица нагрузок)"] = sec_acad(b)
    snap["спецификация (позиция → кол-во)"] = spec
    snap["спецификация: подсчёт"] = spec_cnt
    snap["спецификация: предупреждение"] = spec_warn
    snap["ожидания LISP (SX_log)"] = lisp
    return snap, warns


def dump(snap, path):
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        json.dump(snap, f, ensure_ascii=False, indent=1)
        f.write("\n")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("book")
    ap.add_argument("-o", "--out", help="куда записать JSON (по умолчанию - <имя книги>.snapshot.json в текущей папке)")
    ap.add_argument("--backend", default="auto", choices=["auto", "excel", "libreoffice", "cache"])
    a = ap.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    snap, warns = build(a.book, a.backend)
    out = a.out or os.path.splitext(os.path.basename(a.book))[0] + ".snapshot.json"
    dump(snap, out)
    for w in warns:
        print(w)
    print("Снимок записан:", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
