# Клеммы DKC push-in серии VPR — вид спереди (сверху на клемму на DIN-рейке), 1:1, мм.
# Размеры: карточки dkc.ru (длина, ширина/шаг, торцевые изоляторы 2,2 мм).
# Положение «башенок» (гнездо провода + оранжевая кнопка) — по боковым фото DKC, цвета — по фото.
# Точка вставки блока: левый край, Y=0 — середина длины (как у прежнего VPR-2.5-3L-GY в базе).
#   py make_vpr.py lib.dxf            — все блоки + лист-каталог в модели
#   py make_vpr.py raw.dxf raw NAME   — один блок геометрией прямо в модели (для переопределения в базе)
import math, random, sys
import ezdxf
from ezdxf.enums import TextEntityAlignment as TA

OUT = sys.argv[1]
RAW = len(sys.argv) > 3 and sys.argv[2] == "raw"
RAW_NAME = sys.argv[3] if RAW else None

# ---------------------------------------------------------------- геометрия типов
# башенки: (u0, u1, точек, уровень); u — вдоль длины от нижнего торца. Середина/зазоры — всё, что между.
GEO = {
    "1.5":         (45.5, 3.5,  [(1.4, 13.7, 1, 1), (31.4, 44.0, 1, 1)]),
    "1.5-TWIN":    (54.3, 3.5,  [(1.4, 25.4, 2, 1), (41.6, 52.9, 1, 1)]),
    "1.5-QUATTRO": (63.3, 3.5,  [(1.4, 25.4, 2, 1), (37.9, 61.9, 2, 1)]),
    "2.5":         (48.7, 5.2,  [(1.8, 14.7, 1, 1), (33.9, 47.0, 1, 1)]),
    "2.5-TWIN":    (60.7, 5.2,  [(2.1, 26.1, 2, 1), (45.5, 58.6, 1, 1)]),
    "2.5-QUATTRO": (73.0, 5.2,  [(1.7, 26.8, 2, 1), (46.1, 71.3, 2, 1)]),
    "4":           (55.9, 6.2,  [(1.9, 18.0, 1, 1), (37.0, 54.5, 1, 1)]),
    "4-TWIN":      (67.0, 6.2,  [(1.5, 29.0, 2, 1), (48.2, 65.5, 1, 1)]),
    "4-QUATTRO":   (79.0, 6.2,  [(1.4, 29.1, 2, 1), (48.3, 77.6, 2, 1)]),
    "6":           (57.7, 8.2,  [(1.4, 17.8, 1, 1), (36.8, 56.3, 1, 1)]),
    "6-TWIN":      (74.2, 8.2,  [(1.3, 33.5, 2, 1), (54.1, 72.9, 1, 1)]),
    "10":          (67.7, 10.2, [(1.2, 20.9, 1, 1), (41.8, 66.5, 1, 1)]),
    "10-TWIN":     (88.9, 10.2, [(1.3, 41.5, 2, 1), (62.7, 87.6, 1, 1)]),
    "16":          (75.4, 12.2, [(1.2, 20.4, 1, 1), (54.1, 74.3, 1, 1)]),
    "16-TWIN":     (100.2, 12.2, [(1.5, 45.3, 2, 1), (69.8, 99.1, 1, 1)]),
    "2.5-3L":      (103.0, 5.2, [(0.9, 14.3, 1, 1), (19.4, 33.1, 1, 2), (37.8, 54.0, 1, 3),
                                 (58.3, 71.5, 1, 3), (72.8, 86.7, 1, 2), (87.5, 102.1, 1, 1)]),
}

# цвета по фото DKC
COL = {"GY": (152, 150, 146), "BU": (40, 112, 180), "OR": (226, 112, 34), "PE": (78, 162, 112)}
C_YEL = (196, 196, 92)          # жёлтые прожилки PE
C_BTN = (228, 111, 48)          # кнопки push-in
C_BTN_OR = (178, 70, 18)        # кнопки на оранжевой клемме — темнее, чтобы читались
C_HOLE = (52, 52, 55)           # гнёзда провода, шахты перемычек
C_END = (171, 167, 161)         # торцевой изолятор

# артикулы: (имя блока, геометрия, цвет)
ARTS = []
for g, cols in [("1.5", ["GY", "BU", "OR", "PE"]), ("1.5-TWIN", ["GY", "BU", "OR", "PE"]),
                ("1.5-QUATTRO", ["GY", "OR", "PE"]),
                ("2.5", ["GY", "BU", "PE-YG"]), ("2.5-TWIN", ["GY", "BU", "PE-YG"]),
                ("2.5-QUATTRO", ["GY", "BU", "PE"]),
                ("4", ["GY", "BU", "PE-YG"]), ("4-TWIN", ["GY", "BU", "PE-YG"]), ("4-QUATTRO", ["GY", "BU", "PE"]),
                ("6", ["GY", "BU", "PE-YG"]), ("6-TWIN", ["GY"]),
                ("10", ["GY", "BU", "PE-YG"]), ("10-TWIN", ["GY"]),
                ("16", ["GY", "BU", "PE-YG"]), ("16-TWIN", ["GY"]),
                ("2.5-3L", ["GY", "BU", "PE-YG", "PV-GY", "PV-BU"])]:
    for c in cols:
        ARTS.append((f"VPR-{g}-{c}", g, c.replace("PV-", "").replace("PE-YG", "PE")))
LEN_FIX = {"VPR-6-PE-YG": 55.7}                       # у PE 6 мм² длина по каталогу 55,7
ENDS = [(f"D-VPR-{g}", g) for g in GEO]                # торцевые изоляторы
END_LEN = {"D-VPR-1.5": 45.0, "D-VPR-1.5-TWIN": 54.1, "D-VPR-2.5-QUATTRO": 72.5, "D-VPR-2.5-TWIN": 60.6,
           "D-VPR-4": 56.0, "D-VPR-4-QUATTRO": 78.4, "D-VPR-6-TWIN": 74.0, "D-VPR-10": 67.8,
           "D-VPR-10-TWIN": 88.7, "D-VPR-16-TWIN": 100.0}
T_END = 2.2

# ---------------------------------------------------------------- DXF
doc = ezdxf.new("R2018", setup=False)
doc.units = ezdxf.units.MM
doc.header["$INSUNITS"] = 4
doc.header["$INSBASE"] = (0, 0, 0)
doc.styles.new("VPR_Arial", dxfattribs={"font": "arial.ttf"})
A = {"layer": "0"}


def shade(c, k):
    return tuple(max(0, min(255, int(v * k))) for v in c)


class B:
    """построитель одного блока; u — вдоль длины (Y), x — поперёк (0..W)"""

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

    def fround(self, x0, u0, x1, u1, r, rgb):
        h = self.l.add_hatch(color=8, dxfattribs=dict(A))
        h.rgb = rgb
        h.paths.add_polyline_path(self.rr(x0, u0, x1, u1, r), is_closed=True)

    def rr(self, x0, u0, x1, u1, r):
        r = min(r, (x1 - x0) / 2 - 1e-3, (u1 - u0) / 2 - 1e-3)
        b = math.tan(math.pi / 8)
        pts = [(x0 + r, u0, 0), (x1 - r, u0, b), (x1, u0 + r, 0), (x1, u1 - r, b),
               (x1 - r, u1, 0), (x0 + r, u1, b), (x0, u1 - r, 0), (x0, u0 + r, b)]
        return [(*self.P(x, u), bb) for x, u, bb in pts]

    def round_pl(self, x0, u0, x1, u1, r, rgb):
        e = self.l.add_lwpolyline(self.rr(x0, u0, x1, u1, r), format="xyb", close=True, dxfattribs=dict(A))
        e.rgb = rgb

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

    def pline(self, pts, rgb=None, close=True):
        e = self.l.add_lwpolyline([self.P(*p) for p in pts], close=close, dxfattribs=dict(A))
        if rgb:
            e.rgb = rgb


def marble(b, W, L, seed):
    """жёлто-зелёный «мрамор» PE: частые тонкие косые прожилки во всю ширину"""
    rnd = random.Random(seed)
    u = rnd.uniform(0.2, 0.8)
    while u < L - 0.8:
        t = rnd.uniform(0.15, 0.55)
        skew = rnd.uniform(-1.2, 1.2)
        u0 = max(0.0, min(L - t, u))
        u1 = max(0.0, min(L - t, u + skew))
        b.fill([(0, u0), (W, u1), (W, u1 + t), (0, u0 + t)], C_YEL)
        u += t + rnd.uniform(0.5, 1.8)


def point(b, W, o, i, body, btn):
    """одна точка подключения в пролёте башенки: o — внешний край (к торцу), i — внутренний.
    По фото сверху: круглое гнездо провода у торца (воронка, темнеет вглубь), оранжевая кнопка
    со шлицем под отвёртку, у внутренней стенки — маленькое овальное гнездо под щуп."""
    sg = 1 if i > o else -1
    T = abs(i - o)
    at = lambda f: o + sg * f * T
    # гнездо провода — круглая воронка почти во всю ширину
    dh = min(W - 0.7, 0.36 * T)
    uc = o + sg * (0.07 * T + dh / 2)
    b.fcircle(W / 2, uc, dh / 2, shade(body, 0.74))                          # верх воронки
    b.fcircle(W / 2, uc + sg * dh * 0.07, dh / 2 * 0.74, shade(body, 0.52))  # стенка
    b.fcircle(W / 2, uc + sg * dh * 0.13, dh / 2 * 0.42, shade(body, 0.30))  # глубина
    b.circle(W / 2, uc, dh / 2, shade(body, 0.55))
    # кнопка push-in: оранжевая, косой шлиц посередине
    bw = W * 0.62
    lb = min(0.27 * T, W * 0.8)
    c0, c1 = sorted((at(0.50), at(0.50) + sg * lb))
    x0, x1 = (W - bw) / 2, (W + bw) / 2
    b.fround(x0, c0, x1, c1, min(0.3, bw * 0.12), btn)
    m, sk, hw = (c0 + c1) / 2, lb * 0.13, max(0.12, lb * 0.06)
    b.fill([(x0, m - hw - sk), (x1, m - hw + sk), (x1, m + hw + sk), (x0, m + hw - sk)], shade(btn, 0.5))
    b.round_pl(x0, c0, x1, c1, min(0.3, bw * 0.12), shade(btn, 0.62))
    # гнездо под щуп — маленький овал у внутренней стенки
    tw_, tl = min(1.6, W * 0.36), min(1.1, W * 0.26)
    tu = at(0.89)
    b.fround((W - tw_) / 2, tu - tl / 2, (W + tw_) / 2, tu + tl / 2, tl / 2, shade(body, 0.42))
    return T


def terminal(lay, name, g, color):
    L0, W, towers = GEO[g]
    L = LEN_FIX.get(name, L0)
    k = L / L0
    b = B(lay, L)
    body = COL[color]
    btn = C_BTN_OR if color == "OR" else C_BTN
    lv_k = {3: 1.0, 2: 0.9, 1: 0.8} if g.endswith("3L") else {1: 1.0}
    edge = shade(body, 0.6)
    tw = [(a * k, c * k, n, lv) for a, c, n, lv in towers]
    # 1. корпус: нижние части (середина, зазоры) темнее башенок
    b.frect(0, 0, W, L, shade(body, 0.86 if not g.endswith("3L") else 0.7))
    for a, c, n, lv in tw:
        b.frect(0, a, W, c, shade(body, lv_k[lv]))
    if color == "PE":
        marble(b, W, L, hash(name) & 0xFFFF)
    # торцевые стенки
    b.frect(0, 0, W, tw[0][0], shade(body, 0.75))
    b.frect(0, tw[-1][1], W, L, shade(body, 0.75))
    # 2. середина: шахты перемычек у башенок, поле маркировки по центру
    gaps = [(tw[j][1], tw[j + 1][0]) for j in range(len(tw) - 1)]
    sw = W * 0.5
    for g0, g1 in gaps:
        M = g1 - g0
        if M < 6:                                     # узкий зазор (3L) — одна шахта
            b.frect((W - sw) / 2, g0 + M * 0.25, (W + sw) / 2, g1 - M * 0.25, C_HOLE)
            continue
        sl = min(2.2, M * 0.14)
        for s0 in (g0 + 1.0, g1 - 1.0 - sl):
            b.frect((W - sw) / 2, s0, (W + sw) / 2, s0 + sl, C_HOLE)
        ml = min(12.0, M * 0.42)                      # паз под маркер
        mw = W * 0.78
        b.fround((W - mw) / 2, (g0 + g1 - ml) / 2, (W + mw) / 2, (g0 + g1 + ml) / 2, 0.4,
                 shade(body, 1.08) if color != "PE" else shade(body, 0.95))
        b.round_pl((W - mw) / 2, (g0 + g1 - ml) / 2, (W + mw) / 2, (g0 + g1 + ml) / 2, 0.4, edge)
    # 3. точки подключения
    for j, (a, c, n, lv) in enumerate(tw):
        lower_side = (a + c) / 2 < L / 2              # башенка у нижнего торца — гнёзда смотрят вниз
        if n == 1:
            o, i = (a, c) if lower_side else (c, a)
            point(b, W, o, i, body, btn)
        else:
            m = (a + c) / 2
            for o, i in ((a, m), (m, c)) if lower_side else ((c, m), (m, a)):
                point(b, W, o, i, body, btn)
            b.line(0, m, W, m, edge)
    # 4. линии: грани башенок, контур
    for a, c, n, lv in tw:
        b.line(0, a, W, a, edge)
        b.line(0, c, W, c, edge)
    b.line(0.25, 0, 0.25, L, shade(body, 1.12))       # блик кромки
    b.pline([(0, 0), (W, 0), (W, L), (0, L)])        # контур — ПоСлою
    return W, L


def endplate(lay, name, g):
    L0, W0, towers = GEO[g]
    L = END_LEN.get(name, L0)
    k = L / L0
    b = B(lay, L)
    b.frect(0, 0, T_END, L, C_END)
    for a, c, n, lv in towers:
        b.frect(0, a * k, T_END, c * k, shade(C_END, 1.06))
        b.line(0, a * k, T_END, a * k, shade(C_END, 0.7))
        b.line(0, c * k, T_END, c * k, shade(C_END, 0.7))
    b.pline([(0, 0), (T_END, 0), (T_END, L), (0, L)])
    return T_END, L


ALL = [(n, lambda lay, n=n, g=g, c=c: terminal(lay, n, g, c)) for n, g, c in ARTS] + \
      [(n, lambda lay, n=n, g=g: endplate(lay, n, g)) for n, g in ENDS]

if RAW:
    fn = dict(ALL)[RAW_NAME]
    fn(doc.modelspace())
else:
    msp = doc.modelspace()
    sizes = {}
    for n, fn in ALL:
        sizes[n] = fn(doc.blocks.new(n, base_point=(0, 0)))
    # лист-каталог: ряды по типоразмерам
    groups = {}
    for n, g, c in ARTS:
        groups.setdefault(g, []).append(n)
    y = 0.0
    for g in GEO:
        names = groups[g] + [f"D-VPR-{g}"]
        x = 0.0
        L = GEO[g][0]
        t = msp.add_text(f"VPR-{g}", height=5, dxfattribs={"style": "VPR_Arial"})
        t.set_placement((-10, y), align=TA.MIDDLE_RIGHT)
        for n in names:
            msp.add_blockref(n, (x, y))
            tt = msp.add_text(n, height=1.6, dxfattribs={"style": "VPR_Arial", "rotation": 90})
            tt.set_placement((x + sizes[n][0] / 2, y + L / 2 + 2), align=TA.MIDDLE_LEFT)
            x += sizes[n][0] + 6
        y -= L + 40
doc.saveas(OUT)
print("ok", len(ALL), "blocks")
