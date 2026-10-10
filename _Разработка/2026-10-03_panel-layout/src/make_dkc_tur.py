# Клеммы DKC с винтовым зажимом (TUR, TUR-R, TUR-M, KRU, KRUH, KRUSLG, TTURB, KRUK) — вид спереди
# (сверху на клемму на DIN-рейке), 1:1, мм.
# Размеры: карточки dkc.ru (длина, ширина/шаг; торцевые изоляторы 2,2 / 2,5 мм).
# Ступени по длине — по боковым фото DKC; сверху: колодец с винтом (крестовой шлиц), уступы ввода провода,
# шахты перемычек и паз маркировки в середине. Цвета — по фото.
# Точка вставки: левый край, Y=0 — середина длины (как у клемм VPR в базе).
#   py make_tur.py lib.dxf     — все блоки + лист-каталог в модели
import math, random, sys
import ezdxf
from ezdxf.enums import TextEntityAlignment as TA

OUT = sys.argv[1]

# ступени: (u0, u1, вид); u — мм вдоль длины для номинальной длины Ln. Виды:
#   end — торец/лапка, ledge — уступ ввода провода, t — башенка с винтом (t2 — с двумя винтами),
#   c — середина (перемычки, маркировка); суффикс 1 — нижний уровень двухуровневых клемм.
SEG = {
    "TUR":         (53.0, [(0, 2.4, "end"), (2.4, 12.0, "ledge"), (12.0, 21.2, "t"), (21.2, 31.6, "c"),
                           (31.6, 40.8, "t"), (40.8, 50.6, "ledge"), (50.6, 53.0, "end")]),
    "TUR-TWIN":    (57.8, [(0, 1.3, "end"), (1.3, 9.0, "ledge"), (9.0, 28.9, "t2"), (28.9, 38.4, "c"),
                           (38.4, 47.8, "t"), (47.8, 56.4, "ledge"), (56.4, 57.8, "end")]),
    "TUR-QUATTRO": (65.4, [(0, 0.9, "end"), (0.9, 9.2, "ledge"), (9.2, 29.4, "t2"), (29.4, 38.6, "c"),
                           (38.6, 55.6, "t2"), (55.6, 64.4, "ledge"), (64.4, 65.4, "end")]),
    "TUR-16":      (55.3, [(0, 1.0, "end"), (1.0, 9.0, "ledge"), (9.0, 21.2, "t"), (21.2, 33.7, "c"),
                           (33.7, 45.6, "t"), (45.6, 54.3, "ledge"), (54.3, 55.3, "end")]),
    "TUR-35R":     (56.4, [(0, 2.0, "end"), (2.0, 8.0, "ledge"), (8.0, 22.5, "t"), (22.5, 34.0, "c"),
                           (34.0, 48.5, "t"), (48.5, 54.5, "ledge"), (54.5, 56.4, "end")]),
    "TUR-M":       (28.0, [(0, 1.0, "end"), (1.0, 3.5, "ledge"), (3.5, 11.0, "t"), (11.0, 17.0, "c"),
                           (17.0, 24.5, "t"), (24.5, 27.0, "ledge"), (27.0, 28.0, "end")]),
    "KRU-35N":     (50.0, [(0, 1.5, "end"), (1.5, 7.0, "ledge"), (7.0, 21.0, "t"), (21.0, 29.0, "c"),
                           (29.0, 43.0, "t"), (43.0, 48.5, "ledge"), (48.5, 50.0, "end")]),
    "KRUSLG":      (100.0, [(0, 2.5, "end"), (2.5, 16.0, "ledge"), (16.0, 84.0, "t2"), (84.0, 97.5, "ledge"),
                            (97.5, 100.0, "end")]),
    "KRUH":        (100.0, [(0, 2.0, "end"), (2.0, 98.0, "t2"), (98.0, 100.0, "end")]),
    "TTURB-4":     (73.5, [(0, 5.4, "end"), (5.4, 14.0, "t1"), (14.0, 23.6, "ledge1"), (23.6, 27.7, "ledge"),
                           (27.7, 36.6, "t"), (36.6, 46.2, "c"), (46.2, 55.2, "t"), (55.2, 59.6, "ledge"),
                           (59.6, 68.0, "t1"), (68.0, 73.5, "end")]),
    "KRUK":        (56.0, [(0, 5.0, "end"), (5.0, 13.0, "t1"), (13.0, 16.5, "ledge"), (16.5, 25.0, "t"),
                           (25.0, 31.0, "c"), (31.0, 39.5, "t"), (39.5, 42.8, "ledge"), (42.8, 51.0, "t1"),
                           (51.0, 56.0, "end")]),
    "KRUKB":       (67.0, [(0, 5.0, "end"), (5.0, 13.0, "t1"), (13.0, 24.0, "ledge1"), (24.0, 27.8, "ledge"),
                           (27.8, 36.0, "t"), (36.0, 42.4, "c"), (42.4, 50.6, "t"), (50.6, 54.3, "ledge"),
                           (54.3, 62.0, "t1"), (62.0, 67.0, "end")]),
}
SHADE = {"t": 1.0, "t2": 1.0, "c": 0.88, "ledge": 0.8, "t1": 0.8, "ledge1": 0.66, "end": 0.68}

# артикулы: (артикул, ступени, длина, ширина, цвет)
ARTS = []
def add(art, seg, L, W, col):
    ARTS.append((art, seg, L, W, col))
for s, W in (("2.5", 5.2), ("4", 6.2), ("6", 8.2), ("10", 10.2)):
    for suf, col in (("", "GY"), ("-BU", "BU"), ("-PE", "PE")):
        add(f"TUR-{s}{suf}", "TUR", 53.0, W, col)
for s, W in (("2.5", 5.2), ("4", 6.2)):
    for suf, col in (("", "GY"), ("-BU", "BU")):
        add(f"TUR-{s}-TWIN{suf}", "TUR-TWIN", 57.8, W, col)
        add(f"TUR-{s}-QUATTRO{suf}", "TUR-QUATTRO", 65.4, W, col)
add("TUR-16", "TUR-16", 55.3, 12.0, "GY"); add("TUR-16-BU", "TUR-16", 55.3, 12.0, "BU")
for suf, col in (("", "GY"), ("-BK", "BK"), ("-BU", "BU"), ("-GN", "GN"), ("-OR", "OR"), ("-RD", "RD"), ("-YE", "YE")):
    add(f"TUR-2.5R{suf}", "TUR", 53.3, 5.2, col)
for suf, col in (("", "GY"), ("-BU", "BU"), ("-OR", "OR"), ("-RD", "RD"), ("-YE", "YE")):
    add(f"TUR-4R{suf}", "TUR", 53.3, 6.2, col)
add("TUR-35R", "TUR-35R", 56.4, 16.0, "GY"); add("TUR-35R-BU", "TUR-35R", 56.4, 16.0, "BU")
for s, W, sufs in (("2.5", 5.2, (("", "GY"), ("-BU", "BU"), ("-PE", "PE"))), ("4", 6.2, (("", "GY"), ("-BU", "BU")))):
    for suf, col in sufs:
        add(f"TUR-M-{s}{suf}", "TUR-M", 28.0, W, col)
add("KRU-35N", "KRU-35N", 50.0, 15.2, "GY"); add("KRU-35N-BU", "KRU-35N", 50.0, 15.2, "BU")
for s, L, W in (("16N", 42.5, 12.2), ("35N", 50.0, 15.2), ("50N", 70.5, 20.0), ("95N", 83.0, 25.0)):
    add(f"KRUSLG-{s}", "KRUSLG", L, W, "PE")
for s, L, W in (("50", 70.5, 20.0), ("95", 83.0, 25.0), ("150", 100.0, 31.0), ("240", 100.0, 36.0)):
    add(f"KRUH-{s}", "KRUH", L, W, "GYL"); add(f"KRUH-{s}-BU", "KRUH", L, W, "BU")
add("TTURB-4", "TTURB-4", 73.5, 6.2, "GY"); add("TTURB-4-BU", "TTURB-4", 73.5, 6.2, "BU")
add("KRUK-3", "KRUK", 56.0, 5.2, "GY"); add("KRUK-3-PV", "KRUK", 56.0, 5.2, "GY")
add("KRUKB-3", "KRUKB", 67.0, 5.2, "GY"); add("KRUKB-3-BU", "KRUKB", 67.0, 5.2, "BU")
add("KRUKB-5", "KRUKB", 67.0, 6.2, "GY"); add("KRUKB-5-BU", "KRUKB", 67.0, 6.2, "BU")
add("KRUKB-5-PE", "KRUKB", 67.0, 6.2, "PE"); add("KRUK-5-PV", "KRUKB", 67.0, 6.2, "GY")

# торцевые изоляторы: (артикул, ступени, длина, толщина, цвет)
ENDS = [("D-TUR-2.5-10", "TUR", 50.0, 2.2, "END"), ("D-TUR-2.5-10R", "TUR", 52.7, 2.2, "END"),
        ("D-TUR-2.5-10R-BU", "TUR", 50.0, 2.2, "BU"), ("D-TUR-2.5-TWIN", "TUR-TWIN", 57.8, 2.2, "END"),
        ("D-TUR-2.5-QUATTRO", "TUR-QUATTRO", 65.4, 2.2, "END"), ("D-TUR-4-TWIN", "TUR-TWIN", 57.8, 2.2, "END"),
        ("D-TUR-4-QUATTRO", "TUR-QUATTRO", 65.4, 2.2, "END"), ("D-TUR-16", "TUR-16", 54.4, 2.2, "END"),
        ("D-TUR-16R-BU", "TUR-16", 54.4, 2.2, "BU"), ("D-TUR-35R", "TUR-35R", 56.4, 2.2, "END"),
        ("D-TUR-35R-BU", "TUR-35R", 55.8, 2.2, "BU"), ("D-TUR-M-2.5-4", "TUR-M", 28.1, 2.2, "END"),
        ("D-TTURB-4", "TTURB-4", 73.2, 2.2, "END"), ("D-KRUK-3-5", "KRUK", 56.5, 2.5, "END"),
        ("D-KRUKB-3-5", "KRUKB", 67.7, 2.5, "END")]

COL = {"GY": (152, 150, 146), "GYL": (168, 167, 163), "BU": (40, 112, 180), "OR": (226, 112, 34),
       "PE": (78, 162, 112), "BK": (46, 46, 48), "GN": (40, 140, 70), "RD": (196, 42, 42),
       "YE": (226, 196, 40), "END": (171, 167, 161)}
C_YEL = (196, 196, 92)
C_SCREW = (196, 198, 202)       # оцинкованный винт
C_SLOT = (92, 94, 98)
C_HOLE = (52, 52, 55)

doc = ezdxf.new("R2018", setup=False)
doc.units = ezdxf.units.MM
doc.header["$INSUNITS"] = 4
doc.header["$INSBASE"] = (0, 0, 0)
doc.styles.new("TUR_Arial", dxfattribs={"font": "arial.ttf"})
A = {"layer": "0"}


def shade(c, k):
    return tuple(max(0, min(255, int(v * k))) for v in c)


class B:
    def __init__(self, lay, L):
        self.l, self.h = lay, L / 2.0

    def P(self, x, u):
        return (x, u - self.h)

    def fill(self, pts, rgb):
        h = self.l.add_hatch(color=8, dxfattribs=dict(A))
        h.rgb = rgb
        h.paths.add_polyline_path([self.P(*p) for p in pts], is_closed=True)

    def frect(self, x0, u0, x1, u1, rgb):
        self.fill([(x0, u0), (x1, u0), (x1, u1), (x0, u1)], rgb)

    def rr(self, x0, u0, x1, u1, r):
        r = min(r, (x1 - x0) / 2 - 1e-3, (u1 - u0) / 2 - 1e-3)
        b = math.tan(math.pi / 8)
        pts = [(x0 + r, u0, 0), (x1 - r, u0, b), (x1, u0 + r, 0), (x1, u1 - r, b),
               (x1 - r, u1, 0), (x0 + r, u1, b), (x0, u1 - r, 0), (x0, u0 + r, b)]
        return [(*self.P(x, u), bb) for x, u, bb in pts]

    def fround(self, x0, u0, x1, u1, r, rgb):
        h = self.l.add_hatch(color=8, dxfattribs=dict(A))
        h.rgb = rgb
        h.paths.add_polyline_path(self.rr(x0, u0, x1, u1, r), is_closed=True)

    def round_pl(self, x0, u0, x1, u1, r, rgb):
        self.l.add_lwpolyline(self.rr(x0, u0, x1, u1, r), format="xyb", close=True, dxfattribs=dict(A)).rgb = rgb

    def fcircle(self, x, u, r, rgb):
        h = self.l.add_hatch(color=8, dxfattribs=dict(A))
        h.rgb = rgb
        h.paths.add_edge_path().add_arc(self.P(x, u), r, 0, 360)

    def circle(self, x, u, r, rgb):
        self.l.add_circle(self.P(x, u), r, dxfattribs=dict(A)).rgb = rgb

    def line(self, x0, u0, x1, u1, rgb=None):
        e = self.l.add_line(self.P(x0, u0), self.P(x1, u1), dxfattribs=dict(A))
        if rgb:
            e.rgb = rgb

    def pline(self, pts, rgb=None):
        e = self.l.add_lwpolyline([self.P(*p) for p in pts], close=True, dxfattribs=dict(A))
        if rgb:
            e.rgb = rgb


def marble(b, W, L, seed):
    rnd = random.Random(seed)
    u = rnd.uniform(0.2, 0.8)
    while u < L - 0.8:
        t = rnd.uniform(0.15, 0.55) * max(1.0, W / 8)
        skew = rnd.uniform(-1.2, 1.2) * max(1.0, W / 8)
        u0 = max(0.0, min(L - t, u))
        u1 = max(0.0, min(L - t, u + skew))
        b.fill([(0, u0), (W, u1), (W, u1 + t), (0, u0 + t)], C_YEL)
        u += t + rnd.uniform(0.5, 1.8) * max(1.0, W / 8)


def screw(b, W, u, T, body):
    """колодец с винтом: тень колодца, головка, крестовой шлиц (комби PZ/шлиц)"""
    d = min(W - 1.0, T - 1.2, 9.0 if W < 20 else 14.0)
    b.fcircle(W / 2, u, d / 2, shade(body, 0.42))                  # колодец
    rs = d / 2 * 0.82
    b.fcircle(W / 2, u, rs, C_SCREW)                               # головка
    b.circle(W / 2, u, rs, shade(C_SCREW, 0.62))
    b.circle(W / 2, u, rs * 0.72, shade(C_SCREW, 0.82))
    sw, sl = max(0.18, rs * 0.16), rs * 0.86                       # шлиц
    b.frect(W / 2 - sl, u - sw, W / 2 + sl, u + sw, C_SLOT)
    b.frect(W / 2 - sw, u - sl * 0.62, W / 2 + sw, u + sl * 0.62, C_SLOT)
    b.circle(W / 2, u, d / 2, shade(body, 0.55))


def terminal(lay, name, seg, L, W, color):
    Ln, segs = SEG[seg]
    k = L / Ln
    segs = [(a * k, c * k, kind) for a, c, kind in segs]
    b = B(lay, L)
    body = COL[color]
    edge = shade(body, 0.6) if color != "BK" else (95, 95, 98)
    for a, c, kind in segs:
        b.frect(0, a, W, c, shade(body, SHADE[kind]))
    if color == "PE":
        marble(b, W, L, sum(map(ord, name)))
        for a, c, kind in segs:                                     # ступени поверх «мрамора» — полупрозрачно не умеем,
            if kind in ("end", "ledge1"):                            # поэтому только самые глубокие места темним заново
                b.frect(0, a, W, c, shade(body, SHADE[kind]))
    # середина: шахты перемычек и паз маркировки
    for a, c, kind in segs:
        if kind != "c":
            continue
        M = c - a
        sw = W * 0.5
        sl = min(2.2, M * 0.16)
        if M < 8:
            b.frect((W - sw) / 2, a + M * 0.3, (W + sw) / 2, c - M * 0.3, C_HOLE)
            continue
        for s0 in (a + 0.9, c - 0.9 - sl):
            b.frect((W - sw) / 2, s0, (W + sw) / 2, s0 + sl, C_HOLE)
        ml, mw = min(12.0, M * 0.4), W * 0.78
        b.fround((W - mw) / 2, (a + c - ml) / 2, (W + mw) / 2, (a + c + ml) / 2, 0.4,
                 shade(body, 1.08) if color not in ("PE", "BK") else shade(body, 1.25))
        b.round_pl((W - mw) / 2, (a + c - ml) / 2, (W + mw) / 2, (a + c + ml) / 2, 0.4, edge)
    # уступы ввода провода: тёмная щель у стенки башенки (виден край воронки)
    for j, (a, c, kind) in enumerate(segs):
        if kind not in ("ledge", "ledge1"):
            continue
        nxt = segs[j + 1][2] if j + 1 < len(segs) else ""
        prv = segs[j - 1][2] if j else ""
        ew = max(W - 1.4, W * 0.7)
        if nxt.startswith("t"):
            b.fround((W - ew) / 2, c - 1.0, (W + ew) / 2, c - 0.2, 0.3, shade(body, 0.45))
        elif prv.startswith("t"):
            b.fround((W - ew) / 2, a + 0.2, (W + ew) / 2, a + 1.0, 0.3, shade(body, 0.45))
    # винты
    for a, c, kind in segs:
        if kind in ("t", "t1"):
            screw(b, W, (a + c) / 2, c - a, shade(body, SHADE[kind]))
        elif kind == "t2":
            m = (a + c) / 2
            for u0, u1 in ((a, m), (m, c)):
                screw(b, W, (u0 + u1) / 2, u1 - u0, body)
            if seg not in ("KRUH", "KRUSLG"):
                b.line(0, m, W, m, edge)
            else:                                                   # поле маркировки между винтами
                ml, mw = min(14.0, (c - a) * 0.16), W * 0.62
                b.fround((W - mw) / 2, m - ml / 2, (W + mw) / 2, m + ml / 2, 0.6,
                         shade(body, 1.1) if color != "PE" else (236, 232, 200))
                b.round_pl((W - mw) / 2, m - ml / 2, (W + mw) / 2, m + ml / 2, 0.6, edge)
    # грани ступеней, контур
    for a, c, kind in segs[1:]:
        b.line(0, a, W, a, edge)
    b.line(0.25, 0, 0.25, L, shade(body, 1.12) if color != "BK" else (80, 80, 84))
    b.pline([(0, 0), (W, 0), (W, L), (0, L)])
    return W, L


def endplate(lay, name, seg, L, T, color):
    Ln, segs = SEG[seg]
    k = L / Ln
    b = B(lay, L)
    body = COL[color]
    b.frect(0, 0, T, L, body)
    for a, c, kind in segs:
        if kind.startswith("t"):
            b.frect(0, a * k, T, c * k, shade(body, 1.06))
    for a, c, kind in segs[1:]:
        b.line(0, a * k, T, a * k, shade(body, 0.7))
    b.pline([(0, 0), (T, 0), (T, L), (0, L)])
    return T, L


ALL = [(n, (lambda lay, n=n, s=s, L=L, W=W, c=c: terminal(lay, n, s, L, W, c))) for n, s, L, W, c in ARTS] + \
      [(n, (lambda lay, n=n, s=s, L=L, T=T, c=c: endplate(lay, n, s, L, T, c))) for n, s, L, T, c in ENDS]

if __name__ == "__main__":
    msp = doc.modelspace()
    sizes = {}
    for n, fn in ALL:
        sizes[n] = fn(doc.blocks.new(n, base_point=(0, 0)))
    x = 0.0
    for n, _ in ALL:
        msp.add_blockref(n, (x, 0))
        t = msp.add_text(n, height=2.0, dxfattribs={"style": "TUR_Arial", "rotation": 90})
        t.set_placement((x + sizes[n][0] / 2, sizes[n][1] / 2 + 3), align=TA.MIDDLE_LEFT)
        x += sizes[n][0] + 8
    doc.saveas(OUT)
    print("ok", len(ALL), "blocks")
