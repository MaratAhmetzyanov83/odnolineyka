# Arlight: блоки питания на DIN-рейку и 4-канальные диммеры на DIN-рейку — вид спереди 1:1, мм.
# Габариты — карточки arlight.ru (06.10.2026); вид сверху ARV-DR — чертежи Arlight (63,8 / 45,4 / 90,3);
# расположение клемм, надписей, цвета — по фото с arlight.ru (data/arlight/*.jpg, KNX-104 и DALI-504 — фото анфас, масштаб по габариту).
# Точка вставки всех блоков: левый край, Y=0 — ось DIN-рейки (как у KNX-20E-640). Имя блока = модель Arlight («/» -> «-»).
#   py make_arlight.py lib.dxf     — все блоки + лист-каталог в модели (ряды как в базе)
#   py make_arlight.py lib.dxf list — то же и печать таблицы «имя ширина высота» для LISP
import math, sys
import ezdxf
from ezdxf.enums import TextEntityAlignment as TA

OUT = sys.argv[1]
LIST = len(sys.argv) > 2 and sys.argv[2] == "list"

doc = ezdxf.new("R2018", setup=False)
doc.units = ezdxf.units.MM
doc.header["$INSUNITS"] = 4
doc.header["$INSBASE"] = (0, 0, 0)
for st, f in (("AL_Arial", "arial.ttf"), ("AL_ArialBd", "arialbd.ttf"), ("AL_ArialBdI", "arialbi.ttf"),
              ("AL_ArialN", "arialn.ttf")):
    doc.styles.new(st, dxfattribs={"font": f})
A = {"layer": "0"}            # всё на 0: контур ПоСлою, остальное — истинные цвета

# ---------------------------------------------------------------- цвета (по фото)
WHITE = (238, 238, 236)
PRINT_D = (70, 70, 72)          # печать на светлом
PRINT_L = (226, 226, 226)       # печать на тёмном
HOLE = (40, 40, 42)
SCREW = (196, 198, 200)
SCREW_D = (120, 122, 126)
GREEN = (98, 190, 120)          # клеммники
GREEN_D = (40, 120, 70)
GREEN_H = (28, 84, 50)
ORANGE = (238, 120, 40)
ORANGE_D = (176, 78, 20)
RED_BTN = (226, 66, 82)
LED_G = (90, 200, 110)
SILVER = (184, 186, 188)        # алюминий ARV-DRP
SILVER_D = (146, 148, 152)
SILVER_L = (206, 208, 210)
KNX_GREY = (212, 212, 208)      # корпус KNX/DALI IARL
KNX_GREY_L = (232, 232, 229)
DISP = (24, 24, 26)
SEG_R = (232, 40, 36)
SEG_W = (236, 236, 236)


def sh(c, k):
    return tuple(max(0, min(255, int(v * k))) for v in c)


class B:
    """построитель блока; x, y — от левого нижнего угла изделия, y0 — высота оси рейки"""

    def __init__(self, lay, y0):
        self.l, self.dy = lay, -y0

    def P(self, x, y):
        return (x, y + self.dy)

    def _h(self, rgb):
        h = self.l.add_hatch(color=8, dxfattribs=dict(A))
        h.rgb = rgb
        return h

    @staticmethod
    def rr(x0, y0, x1, y1, r):
        r = max(0.0, min(r, (x1 - x0) / 2 - 1e-3, (y1 - y0) / 2 - 1e-3))
        b = math.tan(math.pi / 8)
        return [(x0 + r, y0, 0), (x1 - r, y0, b), (x1, y0 + r, 0), (x1, y1 - r, b),
                (x1 - r, y1, 0), (x0 + r, y1, b), (x0, y1 - r, 0), (x0, y0 + r, b)]

    def frect(self, x0, y0, x1, y1, rgb, r=0):
        if x1 < x0:
            x0, x1 = x1, x0
        if y1 < y0:
            y0, y1 = y1, y0
        h = self._h(rgb)
        if r:
            h.paths.add_polyline_path([(*self.P(x, y), bb) for x, y, bb in self.rr(x0, y0, x1, y1, r)], is_closed=True)
        else:
            h.paths.add_polyline_path([self.P(x0, y0), self.P(x1, y0), self.P(x1, y1), self.P(x0, y1)], is_closed=True)

    def fpoly(self, pts, rgb):
        self._h(rgb).paths.add_polyline_path([self.P(*p) for p in pts], is_closed=True)

    def fcircle(self, x, y, r, rgb):
        self._h(rgb).paths.add_edge_path().add_arc(self.P(x, y), r, 0, 360)

    def rect(self, x0, y0, x1, y1, rgb=None, r=0):
        if r:
            e = self.l.add_lwpolyline([(*self.P(x, y), bb) for x, y, bb in self.rr(x0, y0, x1, y1, r)],
                                      format="xyb", close=True, dxfattribs=dict(A))
        else:
            e = self.l.add_lwpolyline([self.P(x0, y0), self.P(x1, y0), self.P(x1, y1), self.P(x0, y1)],
                                      close=True, dxfattribs=dict(A))
        if rgb:
            e.rgb = rgb
        return e

    def pline(self, pts, rgb=None, close=False):
        e = self.l.add_lwpolyline([self.P(*p) for p in pts], close=close, dxfattribs=dict(A))
        if rgb:
            e.rgb = rgb
        return e

    def line(self, x0, y0, x1, y1, rgb=None):
        e = self.l.add_line(self.P(x0, y0), self.P(x1, y1), dxfattribs=dict(A))
        if rgb:
            e.rgb = rgb
        return e

    def circle(self, x, y, r, rgb=None):
        e = self.l.add_circle(self.P(x, y), r, dxfattribs=dict(A))
        if rgb:
            e.rgb = rgb
        return e

    def arc(self, x, y, r, a0, a1, rgb=None):
        e = self.l.add_arc(self.P(x, y), r, a0, a1, dxfattribs=dict(A))
        if rgb:
            e.rgb = rgb
        return e

    def text(self, s, x, y, h, rgb=PRINT_D, align=TA.MIDDLE_LEFT, style="AL_Arial", wf=0.85, rot=0):
        t = self.l.add_text(s, height=h, dxfattribs={**A, "style": style, "width": wf, "rotation": rot})
        t.set_placement(self.P(x, y), align=align)
        t.rgb = rgb
        return t


# ---------------------------------------------------------------- общие детали
def screw(b, x, y, r, head=SCREW, ring=HOLE, slot=None):
    """винт клеммы сверху: гнездо, головка, крестовой шлиц"""
    b.fcircle(x, y, r, ring)
    b.fcircle(x, y, r * 0.74, head)
    s = slot or sh(head, 0.45)
    k = r * 0.5
    b.line(x - k, y, x + k, y, s)
    b.line(x, y - k, x, y + k, s)


def vents(b, x0, x1, y0, y1, w=0.9, pitch=2.4, rgb=HOLE, r=0.3):
    """ряд вертикальных щелей между x0 и x1"""
    n = int((x1 - x0 + pitch - w) // pitch)
    if n <= 0:
        return
    s = x0 + ((x1 - x0) - (n * pitch - (pitch - w))) / 2
    for i in range(n):
        b.frect(s + i * pitch, y0, s + i * pitch + w, y1, rgb, r=min(r, w / 2.2))


def arlight_logo(b, x, y, h, rgb, mark=True, align="L"):
    """логотип arlight: «кольцо-семёрка» над словом (x — левый край слова или центр)"""
    wl = h * 3.3                                      # ширина слова при width 1.0
    xc = x + wl / 2 if align == "L" else x
    if mark:
        r = h * 0.85
        cy = y + h * 1.75
        b.arc(xc, cy, r, 150, 400, rgb)
        b.arc(xc, cy, r * 0.82, 160, 390, rgb)
        b.pline([(xc - r * 0.45, cy + r * 0.35), (xc + r * 0.35, cy + r * 0.35), (xc - r * 0.1, cy - r * 0.6)], rgb)
    b.text("arlight", xc, y, h, rgb, TA.MIDDLE_CENTER, "AL_ArialBdI", 1.0)


def intelligent_logo(b, x, y, h, rgb, red=(226, 40, 46)):
    """логотип INTELLIGENT ARLIGHT: «песочные часы» + слово; x,y — левый край, середина слова"""
    s = h * 1.15
    x0, yb, yt = x, y - s * 0.75, y + s * 0.75
    b.fpoly([(x0, yt), (x0 + s, yt), (x0 + s * 0.5, y)], rgb)                  # верхний треугольник
    b.pline([(x0, yb), (x0 + s * 0.5, y), (x0 + s * 0.62, yb)], rgb, close=True)  # нижний — контур
    b.fpoly([(x0 + s * 0.62, yb), (x0 + s, yb), (x0 + s * 0.81, yb + s * 0.32)], red)
    b.text("INTELLIGENT", x0 + s * 1.25, y + h * 0.18, h, rgb, TA.MIDDLE_LEFT, "AL_ArialBd", 0.85)
    b.text("A R L I G H T", x0 + s * 1.25, y - h * 0.78, h * 0.42, rgb, TA.MIDDLE_LEFT, "AL_Arial", 1.0)


def certs(b, x, y, h, rgb, items=("EAC", "CE", "RoHS")):
    cx = x
    for it in items:
        if it == "EAC":
            b.text("EAC", cx, y, h, rgb, TA.MIDDLE_LEFT, "AL_ArialN", 0.8)
            cx += h * 2.2
        elif it == "CE":
            b.text("CE", cx, y, h, rgb, TA.MIDDLE_LEFT, "AL_Arial", 1.0)
            cx += h * 2.4
        elif it == "RoHS":
            b.text("RoHS", cx, y, h * 0.62, rgb, TA.MIDDLE_LEFT, "AL_ArialBd", 0.9)
            cx += h * 2.2
        elif it == "box":
            b.rect(cx, y - h * 0.4, cx + h * 0.8, y + h * 0.4, rgb)
            b.rect(cx + h * 0.15, y - h * 0.25, cx + h * 0.65, y + h * 0.25, rgb)
            cx += h * 1.3


def gnd(b, x, y, s, rgb):
    """знак защитного заземления в кружке (как на корпусах ARV-DRP)"""
    b.circle(x, y, s * 0.55, rgb)
    b.line(x, y + s * 0.32, x, y - s * 0.02, rgb)
    for i, w in enumerate((0.56, 0.36, 0.16)):
        b.line(x - s * w / 2, y - s * 0.02 - i * s * 0.13, x + s * w / 2, y - s * 0.02 - i * s * 0.13, rgb)


def label(b, s, x, y, h, rgb, style="AL_ArialBd", wf=0.8):
    """подпись клеммы; «PE» — знаком заземления"""
    if s == "PE":
        gnd(b, x, y, h * 1.15, rgb)
    else:
        b.text(s, x, y, h, rgb, TA.MIDDLE_CENTER, style, wf)


def seg7(b, x, y, h, n, on=SEG_R, dark=None):
    """индикатор «8.8.8»: n цифр, все сегменты включены (вид как на фото)"""
    w = h * 0.55
    t = h * 0.1
    for i in range(n):
        x0 = x + i * (w + h * 0.22)
        for (a, c) in (((x0, y + h), (x0 + w, y + h)), ((x0, y + h / 2), (x0 + w, y + h / 2)), ((x0, y), (x0 + w, y))):
            b.frect(a[0] + t, a[1] - t / 2, c[0] - t, a[1] + t / 2, on)
        for (xx, ya, yb) in ((x0, y + h / 2, y + h), (x0 + w, y + h / 2, y + h), (x0, y, y + h / 2), (x0 + w, y, y + h / 2)):
            b.frect(xx - t / 2, ya + t * 0.7, xx + t / 2, yb - t * 0.7, on)
        b.fcircle(x0 + w + h * 0.11, y, t * 0.6, on)


def qr(b, x0, y0, s, rgb=(30, 30, 32)):
    """QR-код (условный, фиксированный узор) со стороны s"""
    n = 21
    c = s / n
    for (fx, fy) in ((0, 0), (n - 7, n - 7), (0, n - 7)):
        b.frect(x0 + fx * c, y0 + fy * c, x0 + (fx + 7) * c, y0 + (fy + 7) * c, rgb)
        b.frect(x0 + (fx + 1) * c, y0 + (fy + 1) * c, x0 + (fx + 6) * c, y0 + (fy + 6) * c, (240, 240, 240))
        b.frect(x0 + (fx + 2) * c, y0 + (fy + 2) * c, x0 + (fx + 5) * c, y0 + (fy + 5) * c, rgb)
    seed = 12345
    for j in range(n):
        for i in range(n):
            if (i < 8 and (j < 8 or j >= n - 8)) or (i >= n - 8 and j >= n - 8):
                continue
            seed = (seed * 1103515245 + 12345) & 0x7FFFFFFF
            if (seed >> 16) & 1:
                b.frect(x0 + i * c, y0 + j * c, x0 + (i + 1) * c, y0 + (j + 1) * c, rgb)


def pushin(b, x, y0, y1, w, body=GREEN, lever=ORANGE, up=True):
    """одна клемма push-in: зелёное тело и оранжевый рычаг (рычаг к краю корпуса)"""
    b.frect(x, y0, x + w, y1, body)
    b.line(x, y0, x, y1, GREEN_D)
    L = (y1 - y0)
    ly0, ly1 = (y0 + L * 0.28, y1 + 0.6) if up else (y0 - 0.6, y1 - L * 0.28)
    b.frect(x + w * 0.16, ly0, x + w * 0.84, ly1, lever, r=0.3)
    b.frect(x + w * 0.4, ly0 + 1.0, x + w * 0.6, ly1 - 1.0, ORANGE_D)
    hy = y0 + L * 0.14 if up else y1 - L * 0.14                            # гнездо провода
    b.frect(x + w * 0.25, hy - 0.8, x + w * 0.75, hy + 0.8, GREEN_H, r=0.3)


def round_term(b, xs, y, r, body=GREEN, frame=True):
    """клеммник с круглыми гнёздами (ARV-DRP): колодка и кольца гнёзд"""
    p = (xs[1] - xs[0]) if len(xs) > 1 else 2 * r
    if frame:
        b.frect(xs[0] - p / 2 - 0.6, y - r - 1.0, xs[-1] + p / 2 + 0.6, y + r + 1.0, HOLE, r=1.2)
    for x in xs:
        b.fcircle(x, y, r, body)
        b.circle(x, y, r, GREEN_D)
        b.fcircle(x, y, r * 0.62, sh(body, 0.7))
        b.fcircle(x, y, r * 0.38, GREEN_H)


# ================================================================ 1. ARV-DR (белые) / ARV-DRS (чёрные), 90,3 мм
DR_H = 90.3
Z_T = (90.3 - 63.8) / 2         # 13.25 — уровень клемм
Z_N = (90.3 - 45.4) / 2         # 22.45 — край «носика»


VRANGE = {12: "12-13V", 24: "23-25.5V", 48: "47-49V"}   # регулировка выхода — инструкции ARV-DR


def arv_dr(lay, m):
    """Вид спереди по чертежам Arlight (ступени 63,8 / 45,4, корпус 90,3 + защёлка 2,3) и фото:
    выход сверху, вход снизу (защёлка со стороны входа); надписи горизонтально, как читаются на изделии.
    Клеммы по фото: 17,5 — по центру; 35…70 — вход слева, выход справа; 105 — выход слева, вход справа."""
    name, art, W, V, I, P, black = m
    b = B(lay, DR_H / 2)
    body = (40, 40, 43) if black else WHITE
    pr = PRINT_L if black else PRINT_D
    lo, mid, nose = sh(body, 0.93 if not black else 0.9), sh(body, 0.97 if not black else 1.08), sh(body, 1.0 if not black else 1.2)
    edge = sh(body, 0.78) if not black else (78, 78, 82)
    slot = (150, 150, 150) if not black else (16, 16, 18)
    H = DR_H
    # защёлка рейки внизу (выступает на 2,3)
    lw = min(10.0, W - 5)
    b.frect(W / 2 - lw / 2, -2.3, W / 2 + lw / 2, 0.5, sh(body, 0.88), r=0.6)
    b.rect(W / 2 - lw / 2, -2.3, W / 2 + lw / 2, 0.5, edge, r=0.6)
    # ступени
    b.frect(0, 0, W, H, lo, r=0.6)
    b.frect(0, Z_T, W, H - Z_T, mid)
    b.frect(0, Z_N, W, H - Z_N, nose)
    for yy in (Z_T, H - Z_T, Z_N, H - Z_N):
        b.line(0, yy, W, yy, edge)
    vents(b, 1.6, W - 1.6, Z_T + 1.6, Z_N - 1.6, w=0.9, pitch=2.3, rgb=slot)
    vents(b, 1.6, W - 1.6, H - Z_N + 1.6, H - Z_T - 1.6, w=0.9, pitch=2.3, rgb=slot)
    # клеммы
    ty_out, ty_in = H - 6.9, 6.9
    if W < 20:                                          # 17,5 мм: два винта в «восьмёрке»
        for ty in (ty_out, ty_in):
            b.frect(W / 2 - 6.4, ty - 2.5, W / 2 + 6.4, ty + 2.5, HOLE, r=2.5)
            for dx in (-3.6, 3.6):
                screw(b, W / 2 + dx, ty, 2.3)
        xo, xi = [W / 2 - 3.6, W / 2 + 3.6], [W / 2 - 3.6, W / 2 + 3.6]
        lab_o, lab_i = ["+V", "–V"], ["N", "L"]
        pot = (W / 2, H - Z_N - 9.0)                    # подстроечник на этикетке
    else:
        if W < 40:                                      # 35: выход 2 справа, подстроечник левее
            xo = [W - 11.8, W - 4.8]
            lab_o, lab_i = ["–V", "+V"], ["N", "L"]
            pot = (W - 18.6, ty_out)
            xi = [4.8, 11.8]
        elif W < 100:                                   # 52,5 / 70: выход 4 справа, подстроечник с края
            xo = [W - 10.6 - (3 - i) * 5.8 for i in range(4)]
            lab_o, lab_i = ["–V", "–V", "+V", "+V"], ["L", "N"]
            pot = (W - 4.6, ty_out)
            xi = [4.8, 13.9]
        else:                                           # 105: выход 4 слева + подстроечник, вход справа
            xo = [4.8 + i * 5.8 for i in range(4)]
            lab_o, lab_i = ["–V", "–V", "+V", "+V"], ["L", "N"]
            pot = (xo[-1] + 7.4, ty_out)
            xi = [W - 13.9, W - 4.8]
        for x in xo:
            screw(b, x, ty_out, 2.5 if len(xo) == 2 else 2.4)
        for x in xi:
            screw(b, x, ty_in, 2.5)
    b.fcircle(*pot, 2.0, HOLE)
    b.fcircle(*pot, 1.5, (210, 190, 120) if not black else (150, 150, 150))
    b.line(pot[0] - 1.0, pot[1], pot[0] + 1.0, pot[1], HOLE)
    # светодиод — окошко у края этикетки со стороны выхода
    lx = 2.6 if W < 20 else (W - 6.0 if W < 40 else 3.4)
    b.frect(lx, H - Z_N - 2.6, lx + 2.4, H - Z_N - 1.4, LED_G)
    # этикетка на носике: строка OUTPUT у верхнего края, INPUT — у нижнего, текст горизонтально
    lx0, lx1, ly0, ly1 = 1.2, W - 1.2, Z_N + 1.4, H - Z_N - 1.2
    b.rect(lx0, ly0, lx1, ly1, edge, r=0.5)
    vr = VRANGE[V]
    if W < 20:
        b.text("OUTPUT", W / 2, ly1 - 2.6, 1.15, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
        b.text(f"DC {vr}", W / 2, ly1 - 4.4, 1.0, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.75)
        b.text("+V ADJ", W / 2, pot[1] - 3.0, 1.0, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.75)
        arlight_logo(b, W / 2, (ly0 + ly1) / 2 - 7.0, 1.5, pr, mark=True, align="C")
        b.text(f"Art. {art}", W / 2, ly0 + 7.0, 1.0, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.75)
        b.text("INPUT", W / 2, ly0 + 4.4, 1.15, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
        b.text("AC 100-240V", W / 2, ly0 + 2.6, 1.0, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.75)
    else:
        h = 2.2 if W > 80 else (1.9 if W > 60 else (1.6 if W > 40 else 1.2))
        xr = (xo[0] + xo[-1]) / 2
        b.text(f"DC {vr}", xr, ly1 - 1.6, 1.15, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.75)
        b.text("OUTPUT", xr, ly1 - 3.3, 1.15, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
        xi_c = (xi[0] + xi[1]) / 2
        b.text("INPUT", xi_c, ly0 + 3.3, 1.15, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
        b.text("AC 100-240V", xi_c, ly0 + 1.6, 1.15, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.75)
        lines = [("POWER SUPPLY", "AL_ArialBd", h), (name, "AL_Arial", h),
                 ("Input: AC 100-240V, 50/60Hz", "AL_Arial", h * 0.75),
                 (f"Output: DC {V}V, {I}A, {P}W (Max)", "AL_Arial", h * 0.75),
                 (f"Art. {art}   LPS  IP20", "AL_Arial", h * 0.75)]
        tx = lx0 + 1.6
        arlight_logo(b, tx, ly1 - 6.0 - h * 2.0, h * 1.25, pr)
        y = ly1 - 8.0 - h * 2.6
        for s, st, hh in lines:
            b.text(s, tx, y, hh, pr, TA.MIDDLE_LEFT, st, 0.78)
            y -= hh * 1.5
        if W > 40:
            certs(b, lx1 - (22 if W > 60 else 18), ly0 + 6.5, 2.4 if W > 60 else 1.9, pr)
    # подписи клемм — у винтов
    for x, s in zip(xo, lab_o):
        b.text(s, x, H - 11.4, 1.6, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
    for x, s in zip(xi, lab_i):
        b.text(s, x, 11.4, 1.6, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
    b.rect(0, 0, W, H, r=0.6)                           # контур — ПоСлою
    return W, H


# ================================================================ 2. ARV-DRP (металл, 125 мм)
def arv_drp(lay, m):
    name, art, W, V, I, P, pfc = m
    H = 125.0
    b = B(lay, H / 2)
    if W > 100:
        return arv_drp960(b, m)
    # торцы корпуса (крышка-кожух) и лицевая панель с фасками
    b.frect(0, 0, W, H, SILVER_D, r=0.8)
    b.frect(0.8, 4.0, W - 0.8, H - 4.0, SILVER)
    b.frect(0.8, H - 4.0, W - 0.8, H - 1.2, SILVER_L)               # фаска сверху (блик)
    b.frect(0.8, 1.2, W - 0.8, 4.0, sh(SILVER, 0.9))
    for x in (0.8, W - 0.8):
        b.line(x, 1.2, x, H - 1.2, sh(SILVER, 0.75))
    pr = (60, 60, 62)
    # защёлка рейки внизу (чертёж Arlight: ~9,3 мм шириной, выступ 5,8)
    b.frect(W / 2 - 4.65, -5.8, W / 2 + 4.65, 0.5, (150, 152, 156), r=0.6)
    b.rect(W / 2 - 4.65, -5.8, W / 2 + 4.65, 0.5, (100, 100, 104), r=0.6)
    b.frect(W / 2 - 2.5, -4.6, W / 2 + 2.5, -2.6, (90, 90, 94))
    # клеммы — по чертежам Arlight: центр колодки (42: 22,2; 63: 30,5; 84: 29,6), шаг 6,6 / 6,75 / 7,25
    co, po, ci, pi = {42.0: (22.2, 6.6, 23.0, 7.3), 63.0: (30.5, 6.75, 30.5, 7.3), 84.0: (29.6, 7.25, 28.0, 6.0)}[W]
    xo = [co + (i - 1.5) * po for i in range(4)]
    round_term(b, xo, H - 8.0, po * 0.42)
    xi = [ci + (i - 1) * pi for i in range(3)]
    round_term(b, xi, 8.5, min(pi, 6.6) * 0.42)
    for x, s in zip(xo, ["V-", "V-", "V+", "V+"]):
        b.text(s, x, H - 16.5, 1.9, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
    for x, s in zip(xi, ["PE", "N", "L"]):
        label(b, s, x, 17.0, 1.9, pr)
    # полосы OUTPUT / INPUT
    for yy, s1, s2, xc in ((H - 25.5, f"DC {V}V", "OUTPUT", co), (25.0, "INPUT", "AC 100-240V", ci)):
        b.text(s1, xc, yy + 2.0, 1.6, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)  # надписи — над колодкой (как на изделии)
        b.text(s2, xc, yy - 1.0, 1.6, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
        for x0, x1 in ((2.5, xc - 6.5), (xc + 6.5, W - 2.5)):
            yl = yy - 1.0 if s1.startswith("DC") else yy + 2.0
            b.frect(x0, yl - 0.25, x1, yl + 0.25, (70, 70, 72))
            b.frect(x0, yl - 0.95, x1, yl - 0.55, (70, 70, 72))
    # логотип, светодиод, подстроечник
    arlight_logo(b, 3.0, H - 46.0, 2.3, pr)
    lx = W - 11.8                                       # чертёж: светодиод 39,5 и подстроечник 49 мм от верха
    b.fcircle(lx, H - 39.5, 0.8, (60, 70, 60))
    b.text("DC OK", lx, H - 42.6, 1.4, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
    b.fcircle(lx, H - 49.0, 1.9, HOLE)
    b.arc(lx, H - 49.0, 2.6, 270, 60, pr)
    b.text("+V ADJ", lx - 0.5, H - 53.0, 1.3, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
    # текст этикетки
    y = H - 56.5
    for s, h, st in (("SWITCHING", 1.8, "AL_ArialBd"), ("POWER SUPPLY", 1.8, "AL_ArialBd"), (name, 1.8, "AL_Arial"),
                     (f"Art. {art}", 1.4, "AL_ArialBd"), ("Input: AC 100-240V, 50/60Hz", 1.2, "AL_Arial"),
                     (f"Output: DC {V}V, {I}A, {P}W (Max)", 1.2, "AL_Arial"), ("ta: -30...+70°C   IP20", 1.2, "AL_Arial")):
        b.text(s, 2.8, y, h, pr, TA.MIDDLE_LEFT, st, 0.78)
        y -= h * 1.5
    certs(b, 2.8, 31.5, 2.4, pr)
    b.rect(0, 0, W, H, r=0.8)
    return W, H


def arv_drp960(b, m):
    name, art, W, V, I, P, pfc = m
    H = 125.0
    pr = (60, 60, 62)
    b.frect(0, 0, W, H, SILVER_D, r=0.8)
    b.frect(0.6, 0.8, W - 0.6, H - 0.8, SILVER)
    b.frect(1.5, 8.0, W - 1.5, H - 15.0, (232, 232, 230))           # белая лицевая панель
    b.rect(1.5, 8.0, W - 1.5, H - 15.0, (170, 170, 172))
    # выход: 6 клемм слева сверху, разъём DC OK
    xo = [26.0 + i * 7.0 for i in range(6)]
    round_term(b, xo, H - 7.5, 2.9)
    for x, s in zip(xo, ["V+", "V+", "V+", "V-", "V-", "V-"]):
        b.text(s, x, H - 18.5, 1.9, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
    b.frect(5.0, H - 22.0, 8.0, H - 16.5, (230, 230, 230))
    b.rect(5.0, H - 22.0, 8.0, H - 16.5, pr)
    for i, s in enumerate(["DC", "OK", "P+", "P-"]):
        b.text(s, 9.0, H - 17.2 - i * 1.4, 0.9, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.8)
    xi = [W / 2 - 7.0, W / 2, W / 2 + 7.0]
    round_term(b, xi, 8.5, 2.9)
    for x, s in zip(xi, ["PE", "N", "L"]):
        label(b, s, x, 17.5, 1.9, pr)
    for yy, s1, s2, up in ((H - 26.0, f"DC {V}V", "OUTPUT", True), (22.0, "INPUT", "AC 200-240V", False)):
        b.text(s1, 34.0, yy + 1.6, 1.6, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.8)
        b.text(s2, 34.0, yy - 1.4, 1.6, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.8)
        yl = yy - 1.4 if up else yy + 1.6
        for x0, x1 in ((4.0, 32.0), (48.0, W - 6.0)):
            b.frect(x0, yl - 0.25, x1, yl + 0.25, (90, 60, 40))
            b.frect(x0, yl - 0.95, x1, yl - 0.55, (90, 60, 40))
    b.fcircle(12.0, 75.0, 2.6, HOLE)
    b.arc(12.0, 75.0, 3.4, 60, 200, pr)
    b.text("+V ADJ", 15.5, 73.5, 1.4, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.8)
    b.fcircle(12.0, 68.0, 1.4, HOLE)
    b.text("DC OK", 15.5, 68.0, 1.4, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.8)
    arlight_logo(b, 40.0, 80.0, 3.0, pr)
    y = 72.0
    for s, h, st in (("POWER SUPPLY", 2.0, "AL_ArialBd"), (name, 2.0, "AL_Arial"), (f"Art. {art}", 1.6, "AL_ArialBd"),
                     ("Input: AC 200-240V, 50/60Hz, 6A (230V)", 1.4, "AL_Arial"),
                     (f"Output: DC {V}V, {I}A, {P}W (Max)", 1.4, "AL_Arial"), ("ta: -30...+50°C   IP20", 1.4, "AL_Arial")):
        b.text(s, 40.0, y, h, pr, TA.MIDDLE_LEFT, st, 0.8)
        y -= h * 1.55
    certs(b, 100.0, 33.0, 3.0, pr, ("RoHS", "EAC", "CE"))
    b.rect(0, 0, W, H, r=0.8)
    return W, H


# ================================================================ 3. ARV-DRP-Y (металл, push-in)
def arv_drpy(lay, m):
    name, art, W, H, I, P, n_out = m
    b = B(lay, H / 2)
    pr = (60, 60, 62)
    b.frect(0, 0, W, H, SILVER_D, r=0.8)
    b.frect(0.6, 3.5, W - 0.6, H - 3.5, SILVER)
    b.frect(0.6, H - 3.5, W - 0.6, H - 1.0, SILVER_L)
    b.frect(0.6, 1.0, W - 0.6, 3.5, sh(SILVER, 0.9))
    # выход: n клемм push-in сверху
    pw = 4.2
    x0 = W / 2 - n_out * pw / 2 - (2.0 if n_out == 6 else 0)
    for i in range(n_out):
        pushin(b, x0 + i * pw, H - 15.5, H - 4.0, pw, up=True)
    b.rect(x0, H - 15.5, x0 + n_out * pw, H - 4.0, GREEN_D)
    big = W > 45                                       # 480 Вт: 6 силовых, DC OK — отдельный разъём, вход ⏚ N L
    labs = {4: ["V+", "V-", "DC", "OK"], 6: ["DC", "OK", "V+", "V+", "V-", "V-"]}[n_out] if not big else         ["V-", "V-", "V-", "V+", "V+", "V+"]
    for i, s in enumerate(labs):
        b.text(s, x0 + (i + 0.5) * pw, H - 17.5, 1.1, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.7)
    # вход: 3 клеммы снизу
    xi0 = W / 2 - 1.5 * pw
    for i in range(3):
        pushin(b, xi0 + i * pw, 4.0, 15.5, pw, up=False)
    b.rect(xi0, 4.0, xi0 + 3 * pw, 15.5, GREEN_D)
    for i, s in enumerate(["L", "N", "PE"] if not big else ["PE", "N", "L"]):
        label(b, s, xi0 + (i + 0.5) * pw, 17.5, 1.1, pr, wf=0.7)
    for yy, s1, s2, up in ((H - 21.5, "DC 24V", "OUTPUT", True), (22.0, "INPUT", "AC 100-240V", False)):
        b.text(s1, W / 2, yy + 1.2, 1.2, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
        b.text(s2, W / 2, yy - 1.0, 1.2, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
        yl = yy - 1.0 if up else yy + 1.2
        for xa, xb in ((1.8, W / 2 - 5.0), (W / 2 + 5.0, W - 1.8)):
            b.frect(xa, yl - 0.2, xb, yl + 0.2, (80, 80, 82))
    if n_out == 6 and W > 45:                           # 480 Вт: разъём сигнала DC OK
        b.frect(W - 9.0, H - 33.0, W - 4.5, H - 24.0, GREEN)
        for j in range(4):
            b.frect(W - 7.8, H - 31.8 + j * 2.0, W - 5.7, H - 30.6 + j * 2.0, GREEN_H)
        for j, s in enumerate(["P-", "P+", "DC", "OK"]):
            b.text(s, W - 10.0, H - 25.0 - j * 2.0, 0.9, pr, TA.MIDDLE_RIGHT, "AL_ArialBd", 0.7)
    arlight_logo(b, 2.5, H - 37.5, 1.9, pr)
    b.fcircle(W - 6.0, H - 40.5, 1.8, HOLE)
    b.arc(W - 6.0, H - 40.5, 2.5, 70, 200, pr)
    b.text("+V ADJ", W - 9.0, H - 40.5, 1.0, pr, TA.MIDDLE_RIGHT, "AL_ArialBd", 0.75)
    b.fcircle(W - 6.0, H - 46.0, 1.5, HOLE)
    b.text("DC OK", W - 9.0, H - 46.0, 1.0, pr, TA.MIDDLE_RIGHT, "AL_ArialBd", 0.75)
    y = H - 51.0
    for s, h, st in (("БЛОК ПИТАНИЯ", 1.6, "AL_ArialBd"), (name, 1.5, "AL_Arial"), (f"Арт. {art}", 1.1, "AL_ArialBd"),
                     ("Вход: AC 100-240V, 50/60Hz", 1.0, "AL_Arial"), (f"Выход: DC 24-29V, {I}A, {P}W", 1.0, "AL_Arial"),
                     ("ta: -30...+60°C   IP20", 1.0, "AL_Arial")):
        b.text(s, 2.3, y, h, pr, TA.MIDDLE_LEFT, st, 0.72 if W < 35 else 0.8)
        y -= h * 1.55
    certs(b, 2.3, 30.0, 2.0, pr)
    b.rect(0, 0, W, H, r=0.8)
    return W, H


# ================================================================ 4. корпус KNX IARL 72x90 (KNX-104/204, SR-KN041CC)
# по фото KNX-104-DIM-DIN анфас: 9,78 px/мм, начало — левый нижний угол корпуса
def knx72(lay, m):
    name, art, top, bot, txt, brand = m
    W, H = 72.0, 90.0
    b = B(lay, H / 2)
    g, gl, pr = KNX_GREY, KNX_GREY_L, (54, 54, 56)
    b.frect(0, 0, W, H, g, r=0.6)
    b.frect(1.6, 77.7, W - 1.6, H - 0.6, sh(g, 1.04), r=1.5)            # верхняя полка клемм
    b.frect(1.6, 0.6, W - 1.6, 11.8, sh(g, 1.04), r=1.5)                # нижняя полка
    b.rect(1.6, 77.7, W - 1.6, H - 0.6, sh(g, 0.8), r=1.5)
    b.rect(1.6, 0.6, W - 1.6, 11.8, sh(g, 0.8), r=1.5)
    # верхние клеммы: 12 винтов шагом 5,05
    xs = [6.1 + i * 5.05 for i in range(6)] + [40.9 + i * 5.05 for i in range(6)]
    for x in xs:
        screw(b, x, 81.3, 2.25, head=(70, 96, 84), ring=(38, 44, 40), slot=(26, 30, 28))
    for x, s in zip(xs, top):
        if s:
            b.text(s, x, 86.6, 1.4, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.8)
    if top[0] in ("A", ""):                              # скобки «A»/«B» над парами (как на корпусе)
        for i0 in (0, 4, 6, 10):
            x0, x1 = xs[i0], xs[i0 + 1]
            b.pline([(x0, 86.0), (x0, 87.4), (x1, 87.4), (x1, 86.0)], pr)
            if not top[0]:
                b.text("ABCD"[(0, 4, 6, 10).index(i0)], (x0 + x1) / 2, 88.7, 1.4, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.85)
    vents(b, 3.0, W - 3.0, 75.4, 77.4, w=0.85, pitch=2.55, rgb=(60, 60, 62))
    vents(b, 3.0, W - 3.0, 68.3, 69.3, w=0.7, pitch=2.55, rgb=(120, 120, 120))
    # лицевая крышка
    b.frect(0, 22.0, W, 67.5, gl, r=0.8)
    b.rect(0, 22.0, W, 67.5, sh(g, 0.78), r=0.8)
    b.fpoly([(0.4, 22.6), (5.6, 22.6), (0.4, 34.5)], (40, 170, 90))   # зелёный уголок
    if brand == "IARL":
        intelligent_logo(b, 3.4, 61.6, 2.1, (40, 40, 42))
    else:
        arlight_logo(b, 6.0, 59.5, 2.3, (40, 40, 42))
    # светодиоды
    leds = [(30.1, 59.3), (39.1, 59.3), (12.7, 52.5), (21.5, 52.5), (30.2, 52.5), (39.1, 52.5)]
    for x, y in leds:
        b.fcircle(x, y, 1.6, (88, 88, 90))
    b.text("OT.", 32.4, 58.6, 1.25, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.8)
    b.text("OV.", 41.4, 58.6, 1.25, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.8)
    qr(b, 53.6, 48.8, 15.2)
    y = 47.2
    for s, h, st in txt:
        if st != "small":
            b.text(s, 12.0, y, h, pr, TA.MIDDLE_LEFT, "AL_Arial" if st == "reg" else "AL_ArialBd", 0.82)
            y -= h * 1.35
    yy = 43.3
    for s, h, st in txt:
        if st == "small":
            b.text(s, 42.6, yy, h, pr, TA.MIDDLE_LEFT, "AL_Arial", 0.8)
            yy -= h * 1.5
    # кнопки каналов
    for x, s in zip((13.0, 21.7, 30.4, 39.1), "ABCD"):
        b.fcircle(x, 32.0, 1.8, RED_BTN)
        b.circle(x, 32.0, 1.8, (150, 40, 50))
        b.text(s, x, 36.2, 1.7, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.9)
    b.fcircle(50.8, 33.2, 1.4, (88, 88, 90))
    b.fcircle(59.2, 33.4, 1.6, (40, 20, 22))
    b.text("Prog.", 59.2, 37.3, 1.4, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.85)
    b.text("KNX", 12.2, 26.3, 2.2, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 1.0)
    b.text("EIB", 24.5, 26.3, 2.4, pr, TA.MIDDLE_LEFT, "AL_ArialBdI", 1.0)
    certs(b, 34.5, 26.3, 2.6, pr, ("box", "EAC", "CE"))
    b.text("RoHS", 60.9, 27.4, 1.6, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.85)
    b.text("Made in PRC", 60.9, 24.9, 0.9, pr, TA.MIDDLE_LEFT, "AL_Arial", 0.85)
    # низ
    vents(b, 3.0, W - 3.0, 18.5, 19.5, w=0.7, pitch=2.55, rgb=(120, 120, 120))
    vents(b, 3.0, 54.0, 12.2, 14.4, w=0.85, pitch=2.55, rgb=(60, 60, 62))
    xb = [6.6 + i * 5.05 for i in range(6)]
    for x in xb:
        screw(b, x, 8.9, 2.25, head=(70, 96, 84), ring=(38, 44, 40), slot=(26, 30, 28))
    for x, s in zip(xb, bot):
        b.text(s, x, 3.4, 1.8, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.9)
    # клемма шины KNX
    b.frect(56.4, 2.5, 70.6, 15.0, (175, 175, 172))
    b.frect(57.7, 3.1, 63.7, 14.3, (36, 36, 38))
    b.frect(63.7, 3.1, 69.7, 14.3, (206, 40, 46))
    b.frect(59.2, 8.0, 62.2, 12.6, (20, 20, 22))
    b.frect(65.2, 8.0, 68.2, 12.6, (150, 26, 32))
    b.text("-", 60.7, 16.4, 2.2, (110, 110, 110), TA.MIDDLE_CENTER, "AL_Arial", 1.0)
    b.text("+", 66.7, 16.4, 2.2, (110, 110, 110), TA.MIDDLE_CENTER, "AL_Arial", 1.0)
    b.rect(0, 0, W, H, r=0.6)
    return W, H


# ================================================================ 5. DALI-504-TE-DIN 145x90 (8 мод.)
# по фото анфас: X 6,04 px/мм, Y 5,78 px/мм
def dali504(lay, m):
    name, art = m[:2]
    W, H = 145.0, 90.0
    b = B(lay, H / 2)
    g, gl, pr = (226, 226, 224), (240, 240, 238), (54, 54, 56)
    b.frect(0, 0, W, H, g, r=0.6)
    b.frect(1.5, 74.8, W - 1.5, H - 0.6, sh(g, 1.03), r=1.2)
    b.rect(1.5, 74.8, W - 1.5, H - 0.6, sh(g, 0.8), r=1.2)
    for xa in (9.27, 83.8):                               # две колодки по 8 винтов
        xs = [xa + i * 7.4 for i in range(8)]
        b.frect(xs[0] - 3.7, 75.4, xs[-1] + 3.7, 82.6, (112, 186, 124), r=0.8)
        b.rect(xs[0] - 3.7, 75.4, xs[-1] + 3.7, 82.6, (60, 130, 76), r=0.8)
        for x in xs:
            screw(b, x, 79.0, 2.9, head=(150, 176, 156), ring=(64, 128, 78), slot=(60, 80, 66))
    tl = ["L", "L4", "L", "L3", "L", "L2", "L", "L1"]
    tr = ["L", "N", "NC", "NC", "DA", "DA", "DA", "DA"]
    for i in range(8):
        b.text(tl[i], 9.27 + i * 7.4, 84.8, 1.5, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.85)
        b.text(tr[i], 83.8 + i * 7.4, 84.8, 1.5, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.85)
        sym = "↓" if tl[i] == "L" else "~"
        b.text("v" if sym == "↓" else "~", 9.27 + i * 7.4, 87.3, 1.3, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.9)
    b.text("v", 83.8, 87.3, 1.3, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.9)
    b.pline([(113.4, 86.4), (113.4, 87.6), (128.2, 87.6), (128.2, 86.4)], pr)
    b.pline([(120.8, 86.4), (120.8, 88.4), (135.6, 88.4), (135.6, 86.4)], pr)
    # средняя ступень и лицевая крышка
    b.line(1.5, 74.8, W - 1.5, 74.8, sh(g, 0.8))
    b.frect(0, 19.4, W, 67.0, gl, r=0.6)
    b.rect(0, 19.4, W, 67.0, sh(g, 0.8), r=0.6)
    intelligent_logo(b, 12.9, 57.7, 2.6, (100, 100, 102))
    y = 54.6
    for s, h, st in (("PHASE CUT DIMMER", 2.3, "AL_ArialBd"), (name, 2.4, "AL_Arial"), (f"Art. {art}", 1.6, "AL_ArialBd")):
        b.text(s, 60.6, y, h, pr, TA.MIDDLE_LEFT, st, 0.82)
        y -= h * 1.6
    b.text("AC 100-230V, 50/60Hz", 60.6, 42.1, 1.35, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.8)
    b.text("Temp. range: -20...+45 °C", 60.6, 39.6, 1.35, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.8)
    b.frect(104.8, 47.1, 120.5, 63.0, DISP)
    seg7(b, 106.6, 49.6, 9.8, 2)
    for x, s in ((108.1, "0-6"), (116.7, "0-9")):
        b.fcircle(x, 42.8, 1.8, (36, 36, 38))
        b.text(s, x, 38.0, 1.5, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.85)
    b.rect(125.0, 53.4, 139.6, 57.6, pr, r=2.0)
    b.text("DALI", 132.3, 55.5, 1.9, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 1.1)
    for j, s in enumerate(("1×800W, 3.5A", "2×600W, 2×2.5A", "3×500W, 3×2A", "4×400W, 4×1.5A")):
        b.text(s, 125.5, 41.5 - j * 1.9, 1.15, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.75)
    b.fcircle(131.6, 26.3, 1.9, (36, 36, 38))
    b.text("Min.set", 134.3, 26.0, 1.4, pr, TA.MIDDLE_LEFT, "AL_Arial", 0.85)
    b.text("RoHS", 104.0, 26.6, 2.0, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.85)
    b.text("200424V5", 117.5, 28.0, 1.2, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.8)
    b.text("Made in P.R.C.", 117.5, 25.4, 1.1, pr, TA.MIDDLE_LEFT, "AL_Arial", 0.8)
    certs(b, 20.5, 26.3, 3.0, (110, 110, 112), ("EAC", "CE"))
    b.circle(45.5, 26.3, 2.2, (110, 110, 112))
    # условные знаки нагрузки
    sc = (130, 130, 132)
    b.rect(56.8, 24.6, 63.5, 33.0, sc)
    b.line(60.1, 24.6, 60.1, 33.0, sc)
    b.circle(66.5, 28.8, 1.7, sc)
    b.rect(73.5, 24.6, 81.0, 33.0, sc)
    b.line(73.5, 24.6, 81.0, 33.0, sc)
    b.circle(84.5, 28.8, 1.7, sc)
    b.circle(94.5, 28.8, 1.7, sc)
    for a in range(0, 360, 45):
        r0 = math.radians(a)
        b.line(94.5 + 2.3 * math.cos(r0), 28.8 + 2.3 * math.sin(r0), 94.5 + 3.2 * math.cos(r0), 28.8 + 3.2 * math.sin(r0), sc)
    vents(b, 1.0, 14.5, 23.0, 27.5, w=1.2, pitch=2.0, rgb=(206, 206, 204))
    # нижние ступени
    b.line(0.6, 13.9, W - 0.6, 13.9, sh(g, 0.8))
    b.frect(9.5, 1.3, W - 9.5, 12.5, sh(g, 1.02), r=1.0)
    b.rect(9.5, 1.3, W - 9.5, 12.5, sh(g, 0.86), r=1.0)
    b.rect(0, 0, W, H, r=0.6)
    return W, H


# ================================================================ 6. корпус SMART 48 x 110/115 (поперёк рейки, ушки сверху и снизу)
# Профиль — чертёж из инструкции SMART-K36-DMX (110 с защёлкой, глубина 67, по рейке 48, паз рейки 52,4 от торца без защёлки);
# вид сверху — схемы подключения SMART-DALI-104-62 и SMART-D23 (масштаб по ширине 48).
SM_SLOT = 48.0 / 8                                       # шаг поля подписей (8 мест)


def smart(lay, m):
    name, art, H, black, brand, top, bot, tlab, blab, disp, lines = m
    W = 48.0
    L = H - 4.0                                          # корпус с ушками; 4 мм — защёлка снизу
    Y = lambda t: H - t * L / 106.0                      # мм от верха (по чертежу при L=106) -> y блока
    b = B(lay, H - 52.4 * L / 106.0)
    body = (54, 54, 57) if black else WHITE
    pr = PRINT_L if black else PRINT_D
    edge = (96, 96, 100) if black else (150, 150, 148)
    hole = (14, 14, 16) if black else HOLE
    # защёлка
    b.frect(W / 2 - 5.0, 0.0, W / 2 + 5.0, 4.6, sh(body, 0.9 if not black else 1.1), r=0.6)
    b.rect(W / 2 - 5.0, 0.0, W / 2 + 5.0, 4.6, edge, r=0.6)
    # основание с ушками, ступени
    b.frect(0, Y(106), W, Y(0), sh(body, 0.92 if not black else 1.0), r=2.6)
    b.frect(0, Y(94.4), W, Y(11.2), sh(body, 0.96 if not black else 1.08))
    b.frect(0, Y(85.1), W, Y(20.5), sh(body, 0.98 if not black else 1.14))
    b.frect(0, Y(76.7), W, Y(29.3), sh(body, 1.0 if not black else 1.24))
    for t in (11.2, 20.5, 29.3, 76.7, 85.1, 94.4):
        b.line(0, Y(t), W, Y(t), edge)
    for t in (5.6, 100.4):                               # крепёжные отверстия в ушках
        for x in (5.65, 42.35):
            b.fcircle(x, Y(t), 2.9, sh(body, 0.8 if not black else 1.5))
            b.fcircle(x, Y(t), 1.4, hole)
    # винты клемм и «окна» проводов (8 с каждого торца)
    for ts, tq in ((17.3, 22.6), (89.1, 83.3)):
        b.line(4.6, Y(tq), W - 4.6, Y(tq), edge)
        for i in range(8):
            x = 7.6 + i * 4.77
            screw(b, x, Y(ts), 2.15, head=(176, 178, 180), ring=hole)
            b.frect(x - 0.95, Y(tq) - 0.95, x + 0.95, Y(tq) + 0.95, sh(body, 0.85 if not black else 1.6))
            b.rect(x - 0.95, Y(tq) - 0.95, x + 0.95, Y(tq) + 0.95, edge)
    # поле подписей клемм на верхней грани
    for row, t0, labs in ((top, 31.7, tlab), (bot, 72.4, blab)):
        for i, s in enumerate(row):
            if s is None:
                continue
            x0 = 2.9 + i * (42.5 / 8)
            b.rect(x0 + 0.15, Y(t0 + 3.0), x0 + 42.5 / 8 - 0.15, Y(t0), pr)
            b.text(s, x0 + 42.5 / 16, Y(t0 + 1.5), 1.25 if len(s) < 3 else 0.95, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.78)
        for s, x in labs:
            b.text(s, x, Y(t0 + 4.8) if t0 < 50 else Y(t0 - 1.8), 0.95, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.8)
    # логотип, надписи, индикатор
    if brand == "IARL":
        intelligent_logo(b, 2.8, Y(42.6), 1.35, pr)
    else:
        arlight_logo(b, 3.2, Y(43.0), 1.75, pr, mark=False)
        b.arc(2.4, Y(42.5), 1.25, 200, 470, pr)        # «семёрка» в кольце перед словом
    t = 47.4
    for s, h, st in lines:
        b.text(s, 3.2, Y(t), h, pr, TA.MIDDLE_LEFT, st, 0.74)
        t += h * 1.55
    if disp:
        b.frect(20.4, Y(57.7), 45.1, Y(41.4), DISP, r=0.6)
        on = {"R": SEG_R, "W": SEG_W, "B": (90, 190, 240)}[disp]
        seg7(b, 22.4, Y(55.4), 9.5, 3, on=on)
        for x, s in ((23.0, "M"), (32.2, "<"), (41.4, ">")):
            b.text(s, x, Y(60.3), 1.4, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.9)
            b.fcircle(x, Y(64.2), 1.6, (20, 20, 22) if not black else (8, 8, 10))
            b.circle(x, Y(64.2), 1.6, edge)
        b.text("Made in P.R.C.", 45.0, Y(69.2), 0.8, pr, TA.MIDDLE_RIGHT, "AL_Arial", 0.8)
    else:
        b.text("EAC", 37.5, Y(61.0), 3.6, pr, TA.MIDDLE_LEFT, "AL_ArialN", 0.8)
        b.text("RoHS", 36.6, Y(65.6), 1.7, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.9)
    # контур — ПоСлою: корпус с ушками + защёлка
    b.rect(0, Y(106), W, Y(0), r=2.6)
    return W, H


# ================================================================ 7. SMART-DALI-104-82-RGBW-DT6-DT8-RDM-DIN 115 x 97 (+1 разъёмы)
# по габаритному чертежу из инструкции (рис. 1): X 3,19 px/мм, Y 3,04 px/мм; надписи и цвета — по фото
def smart_dali82(lay, m):
    name, art = m[:2]
    W, H = 115.0, 97.0
    b = B(lay, H / 2 + 1.0)
    base = (34, 34, 36)
    pr = (214, 214, 214)
    edge = (86, 86, 90)
    # разъёмные клеммы снизу (выступают на 1 мм) — рисуем первыми, корпус перекроет верх
    for xa, n, pitch in ((14.1, 4, 5.88), (58.0, 6, 5.75)):
        b.frect(xa, -1.0, xa + n * pitch, 6.5, GREEN, r=0.4)
        b.rect(xa, -1.0, xa + n * pitch, 6.5, GREEN_D, r=0.4)
        for i in range(n):
            x = xa + pitch * (i + 0.5)
            b.fcircle(x, 2.6, 1.7, (176, 182, 176))
            b.line(x - 1.0, 3.3, x + 1.0, 1.9, (70, 80, 70))
    b.frect(0, 0.5, W, H, base, r=0.8)
    b.rect(0, 0.5, W, H, edge, r=0.8)
    # верх: откидная крышка над винтовыми клеммами
    b.frect(1.6, 75.0, W - 1.6, H - 1.5, (28, 28, 30))
    b.rect(1.6, 75.0, W - 1.6, H - 1.5, edge)
    b.line(1.6, H - 5.0, W - 1.6, H - 5.0, edge)
    xs = [8.0 + i * 8.6 for i in range(12)]
    labs = ["-", "-", "+", "+", "+", "-", "+", "-", "+", "-", "+", "-"]
    for x, s in zip(xs, labs):
        b.fcircle(x, 86.0, 1.8, (110, 112, 116))
        b.line(x - 1.1, 86.0, x + 1.1, 86.0, (50, 50, 52))
        b.text(s, x, 91.0, 1.5, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.9)
    b.text("НАПРЯЖЕНИЕ ПИТАНИЯ DC 12-48V", 20.9, 79.5, 1.05, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.72)
    for i in range(4):
        b.text(f"CH{i + 1}", (xs[4 + 2 * i] + xs[5 + 2 * i]) / 2, 79.5, 1.2, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.8)
    # лицевая панель
    b.frect(1.0, 27.3, W - 1.0, 75.0, (42, 42, 44))
    b.rect(1.0, 27.3, W - 1.0, 75.0, edge)
    b.frect(11.0, 45.4, 38.6, 63.5, (16, 22, 52), r=0.6)                 # OLED
    b.rect(11.0, 45.4, 38.6, 63.5, (90, 90, 96), r=0.6)
    b.text("DMX    DIM-4CH", 13.0, 60.0, 1.6, (120, 170, 255), TA.MIDDLE_LEFT, "AL_Arial", 0.8)
    b.text("Addr:001  8bit", 13.0, 56.8, 1.6, (120, 170, 255), TA.MIDDLE_LEFT, "AL_Arial", 0.8)
    for x, s in ((15.0, "M"), (24.1, "<"), (33.5, ">")):
        b.fcircle(x, 41.1, 2.2, (20, 20, 22))
        b.circle(x, 41.1, 2.2, (110, 110, 114))
        b.text(s, x, 44.0, 1.2, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.9)
    b.text("DALI ВХОД  DALI ВЫХОД", 11.0, 36.0, 1.05, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.72)
    for j, s in enumerate(["DA1", "DA2", "DA1", "DA2"]):
        b.rect(11.0 + j * 6.0, 30.2, 16.6 + j * 6.0, 33.0, pr)
        b.text(s, 13.8 + j * 6.0, 31.6, 1.0, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.72)
    for i, x in enumerate((50.8, 66.5, 82.1, 97.8)):                     # K1…K4 и светодиоды каналов
        b.fcircle(x, 60.2, 0.8, (110, 110, 110))
        b.fcircle(x, 52.0, 4.7, (20, 20, 22))
        b.circle(x, 52.0, 4.7, (120, 120, 124))
        b.text(f"K{i + 1}", x, 52.0, 1.8, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.85)
    intelligent_logo(b, 45.0, 42.4, 1.5, pr)
    for j, (s, h, st) in enumerate((("ДИММЕР", 1.4, "AL_ArialBd"), ("SMART-DALI-104-82-RGBW-DT6/DT8-RDM-DIN", 1.15, "AL_Arial"),
                                    (f"(12-48V, 4x12A, DMX)   Арт. {art}", 1.1, "AL_Arial"))):
        b.text(s, 45.0, 38.6 - j * 1.9, h, pr, TA.MIDDLE_LEFT, st, 0.75)
    b.text("DMX ВХОД   DMX ВЫХОД", 72.0, 31.8, 1.0, pr, TA.MIDDLE_LEFT, "AL_ArialBd", 0.72)
    for j, s in enumerate(["GND", "D-", "D+", "GND", "D-", "D+"]):
        b.rect(72.0 + j * 5.2, 28.0, 76.8 + j * 5.2, 30.4, pr)
        b.text(s, 74.4 + j * 5.2, 29.2, 0.9, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.7)
    b.text("EAC  CE  RoHS", W - 3.0, 70.5, 1.6, pr, TA.MIDDLE_RIGHT, "AL_ArialBd", 0.85)
    # нижняя полоса: индикаторы шин, RJ45
    b.line(1.0, 27.3, W - 1.0, 27.3, edge)
    for x, s, al in ((6.9, "DALI COM", TA.MIDDLE_LEFT), (107.5, "DMX COM", TA.MIDDLE_RIGHT)):
        b.fcircle(x, 19.7, 0.9, (120, 120, 120))
        b.text(s, x + (2.0 if al == TA.MIDDLE_LEFT else -2.0), 19.7, 1.1, (160, 160, 160), al, "AL_ArialBd", 0.8)
    b.frect(96.0, 0.5, 107.0, 5.0, (190, 190, 192))                    # RJ45 (торец)
    b.frect(97.5, 0.5, 105.5, 3.5, (40, 40, 42))
    b.frect(50.6, 0.5, 56.0, 3.2, (24, 24, 26), r=0.5)                  # защёлка рейки
    return W, H


# ================================================================ 8. корпус SR 53 x 98 (+ крючок и лапки до 110), рычажные клеммы
# по схеме подключения SR-2108FA-RJ45-DIN из инструкции (масштаб: ширина 53, высота с лапками 110) и фото SR-2303
def sr(lay, m):
    name, art, kind, top, bot, ndig, buttons, lines = m
    W, H = 53.0, 110.0
    Y = lambda t: H - t                                    # мм от верха
    b = B(lay, H / 2)
    g, pr, edge = (230, 230, 228), (60, 60, 62), (150, 150, 148)
    # крючок сверху, лапки защёлки снизу
    b.frect(20.9, Y(6.0), 30.6, Y(0.4), sh(g, 0.92), r=1.2)
    b.rect(20.9, Y(6.0), 30.6, Y(0.4), edge, r=1.2)
    for x0, x1 in ((9.0, 16.4), (35.8, 43.3)):
        b.frect(x0, Y(110.0), x1, Y(103.0), sh(g, 0.92), r=0.5)
        b.rect(x0, Y(110.0), x1, Y(103.0), edge, r=0.5)
        b.rect(x0 + 1.2, Y(108.6), x1 - 1.2, Y(106.4), edge)
    b.frect(22.4, Y(109.0), 29.9, Y(103.0), sh(g, 0.88), r=2.5)
    # корпус 6…104 от верха
    b.frect(0, Y(104.0), W, Y(6.0), g, r=0.8)
    b.frect(1.2, Y(24.0), W - 1.2, Y(7.2), (50, 50, 52))             # карман верхних клемм
    b.frect(1.2, Y(102.8), W - 1.2, Y(86.0), (50, 50, 52))           # карман нижних клемм
    for t in (24.0, 32.2, 77.8, 86.0):
        b.line(0, Y(t), W, Y(t), edge)
    for t in (24.9, 85.2):
        for i in range(10):
            x = 5.0 + i * 4.8
            b.frect(x - 0.8, Y(t + 0.8), x + 0.8, Y(t - 0.8), sh(g, 0.8))
    def levers(x0, n, t0, t1, labs, up):
        pw = 5.25
        for i in range(n):
            x = x0 + i * pw
            s = labs[i] if i < len(labs) else ""
            col = ORANGE if s in ("+", "V+") else (156, 158, 160)
            b.frect(x + 0.15, Y(t1), x + pw - 0.15, Y(t0), col)
            b.rect(x + 0.15, Y(t1), x + pw - 0.15, Y(t0), sh(col, 0.7))
            lt = (t0 + 0.8, t0 + 6.0) if up else (t1 - 6.0, t1 - 0.8)
            b.frect(x + 0.9, Y(lt[1]), x + pw - 0.9, Y(lt[0]), (250, 250, 250), r=0.4)
            ht = t1 - 2.4 if up else t0 + 2.4
            b.fcircle(x + pw / 2, Y(ht), 1.5, sh(col, 0.45))
    # верхний карман
    if kind == "RJ45":
        for x0 in (4.5, 16.8):                              # два разъёма RJ45 (DMX вход/выход)
            b.frect(x0, Y(22.5), x0 + 12.0, Y(8.5), (196, 196, 198), r=0.4)
            b.frect(x0 + 1.4, Y(20.5), x0 + 10.6, Y(11.5), (34, 34, 36))
            b.frect(x0 + 4.0, Y(22.5), x0 + 8.0, Y(20.5), (34, 34, 36))
        levers(38.8, 2, 9.0, 22.5, ["V+", "V-"], True)
    else:
        levers(W / 2 - 3 * 5.25, 6, 9.0, 22.5, top, True)
    levers(W / 2 - len(bot) * 5.25 / 2, len(bot), 87.5, 101.5, bot, False)
    # верхняя грань
    b.frect(1.6, Y(77.8), W - 1.6, Y(32.2), (238, 238, 236))
    b.rect(1.6, Y(77.8), W - 1.6, Y(32.2), edge)
    for labs, t0, x0 in ((top, 34.6, 40.3 if kind == "RJ45" else W - 2.5 - len(top) * 5.0), (bot, 73.0, 12.7)):
        if kind == "RJ45" and t0 < 50:
            labs = ["V+", "V-"]
        for i, s in enumerate(labs):
            b.rect(x0 + i * 5.2, Y(t0 + 3.0), x0 + i * 5.2 + 4.9, Y(t0), pr)
            b.text(s, x0 + i * 5.2 + 2.45, Y(t0 + 1.5), 1.1, pr, TA.MIDDLE_CENTER, "AL_ArialBd", 0.72)
    if kind == "RJ45":
        b.text("Power input", 45.6, Y(39.2), 0.95, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.8)
        for x0 in (7.0, 17.0):
            b.rect(x0, Y(39.7), x0 + 6.5, Y(36.0), pr)
        b.text("Pin1:Data+ Pin2:Data-", 25.5, Y(35.0), 0.75, pr, TA.MIDDLE_LEFT, "AL_Arial", 0.72)
        b.text("Pin7:GND  Pin8:GND", 25.5, Y(36.6), 0.75, pr, TA.MIDDLE_LEFT, "AL_Arial", 0.72)
        b.text("DMX512 & RDM", 3.0, Y(46.8), 0.95, pr, TA.MIDDLE_LEFT, "AL_Arial", 0.8)
        b.text("decoder", 3.0, Y(48.4), 0.95, pr, TA.MIDDLE_LEFT, "AL_Arial", 0.8)
        dx0, dx1, dt0, dt1 = 17.2, 51.0, 41.2, 54.6
    else:
        b.text("arlight", 3.0, Y(37.8), 2.4, (214, 40, 40), TA.MIDDLE_LEFT, "AL_ArialBdI", 1.0)
        dx0, dx1, dt0, dt1 = 32.0, 50.5, 47.0, 57.0
    t = 41.5 if kind != "RJ45" else 51.0
    for s, h, st in lines:
        b.text(s, 3.0, Y(t), h, pr, TA.MIDDLE_LEFT, st, 0.74)
        t += h * 1.5
    b.frect(dx0, Y(dt1), dx1, Y(dt0), DISP, r=0.8)
    hh = (dt1 - dt0) - 3.5
    seg7(b, dx0 + 1.6, Y(dt1 - 1.6), hh, ndig)
    bt0 = dt1 + 3.2
    if kind == "RJ45":
        b.rect(16.4, Y(67.4), 51.5, Y(56.9), pr, r=1.5)
        bt0 = 61.1
    for j, s in enumerate(buttons):
        x = (20.5 + j * 8.23) if kind == "RJ45" else (dx0 + 4.0 + j * 9.0)
        b.fcircle(x, Y(bt0), 1.6, (30, 30, 32))
        b.text(s, x, Y(bt0 + 3.0), 0.95, pr, TA.MIDDLE_CENTER, "AL_Arial", 0.8)
    b.rect(0, Y(104.0), W, Y(6.0), r=0.8)
    return W, H


# ================================================================ модели
DR = [
    ("ARV-DR15-12", "034670", 17.5, 12, 1.25, 15, False), ("ARV-DR30-12", "034671", 35.0, 12, 2, 24, False),
    ("ARV-DR60-12", "034672", 52.5, 12, 4.5, 54, False), ("ARV-DR100-12", "034673", 70.0, 12, 7.5, 90, False),
    ("ARV-DR135-12", "035699", 105.0, 12, 11.3, 135, False),
    ("ARV-DR15-24", "031084", 17.5, 24, 0.63, 15, False), ("ARV-DR30-24", "031085", 35.0, 24, 1.25, 30, False),
    ("ARV-DR60-24", "031086", 52.5, 24, 2.5, 60, False), ("ARV-DR100-24", "031087", 70.0, 24, 4.2, 100, False),
    ("ARV-DR150-24", "035700", 105.0, 24, 6.25, 150, False),
    ("ARV-DR15-48", "034674", 17.5, 48, 0.32, 15, False), ("ARV-DR30-48", "034675", 35.0, 48, 0.6, 30, False),
    ("ARV-DR60-48", "034676", 52.5, 48, 1.25, 60, False), ("ARV-DR100-48", "034677", 70.0, 48, 2.1, 100, False),
    ("ARV-DR150-48", "035701", 105.0, 48, 3.2, 150, False),
    ("ARV-DRS-15-24", "057073", 17.5, 24, 0.63, 15, True), ("ARV-DRS-36-24", "059900", 35.0, 24, 1.5, 36, True),
    ("ARV-DRS-60-24", "059901", 52.5, 24, 2.5, 60, True), ("ARV-DRS-100-24", "059902", 70.0, 24, 4.2, 100, True),
    ("ARV-DRS-150-24", "057074", 105.0, 24, 6.25, 150, True),
]
DRP = [
    ("ARV-DRP75-24", "040232", 42.0, 24, 3.15, 75, False), ("ARV-DRP120-24-B", "037156", 42.0, 24, 5, 120, False),
    ("ARV-DRP240-24-B", "037155", 63.0, 24, 10, 240, True), ("ARV-DRP480-PFC-24", "040231", 84.0, 24, 20, 480, True),
    ("ARV-DRP960-24", "053072", 150.0, 24, 40, 960, True),
    ("ARV-DRP75-48", "052805", 42.0, 48, 1.6, 75, False), ("ARV-DRP120-48-B", "052806", 42.0, 48, 2.5, 120, False),
    ("ARV-DRP240-PFC-48", "044599", 63.0, 48, 5, 240, True), ("ARV-DRP480-PFC-48", "044601", 84.0, 48, 10, 480, True),
]
DRPY = [
    ("ARV-DRP75-24-Y", "061926", 31.0, 124.0, 3.2, 75, 4), ("ARV-DRP120-24-Y", "061927", 31.0, 124.0, 5, 120, 4),
    ("ARV-DRP240-24-PFC-Y", "061930", 41.0, 124.0, 10, 240, 6), ("ARV-DRP480-24-PFC-Y", "061931", 51.0, 130.0, 20, 480, 6),
]
KNX = [
    ("KNX-104-DIM-DIN", "025658", ["", "", "+", "+", "", "", "", "", "+", "+", "", ""], ["+", "+", "+", "–", "–", "–"],
     [("DIMMER", 2.2, "bold"), ("KNX-104-DIM-DIN", 2.4, "reg"), ("Art. 025658", 1.8, "bold"),
      ("Input: DC 12-24V", 1.2, "small"), ("Output: DC 12-24V, 4×4A, PWM (CV)", 1.2, "small")], "IARL"),
    ("KNX-204-DIM-DIN", "025660", ["A1", "A2", "+", "+", "B1", "B2", "C1", "C2", "+", "+", "D1", "D2"],
     ["+", "+", "+", "–", "–", "–"],
     [("KNX-204-DIM-DIN", 2.2, "bold"), ("Art. 025660", 1.8, "reg"), ("Input: DC 12-48V", 1.2, "small"),
      ("Output: 4×350mA, 2×700mA, 1×1000mA", 1.2, "small")], "IARL"),
    ("SR-KN041CC-DIN", "023042", ["A1", "A2", "+", "+", "B1", "B2", "C1", "C2", "+", "+", "D1", "D2"],
     ["+", "+", "+", "–", "–", "–"],
     [("SR-KN041CC-DIN", 2.2, "bold"), ("Art. 023042", 1.8, "reg"), ("Input: DC 12-48V", 1.2, "small"),
      ("Output: 4×350mA / 2×700mA", 1.2, "small")], "AR"),
]
_N = None
_DALI62_TOP = ["+", "-", _N, _N, "DA2", "DA1", "DA2", "DA1"]
_DALI62_TL = [("POWER INPUT", 8.2), ("SIGNAL INPUT", 35.0)]
SMART = [  # имя, арт, высота (с защёлкой), чёрный, бренд, верх, низ, подписи верх, подписи низ, индикатор, строки
    ("SMART-D23-DIM-PUSH-DIN", "033758", 110.0, False, "AR",
     ["+", "-", "+/-", "+/-", "1", "2", "3", "4"], ["-", "+", "+", "+", "-", "-", "-", "-"],
     [("Push-Dim", 8.2), ("Input 0/1-10V", 30.0)], [("INPUT DC 12-36V", 8.2), ("LED OUTPUT", 30.0)], None,
     [("DIMMER", 1.5, "AL_ArialBd"), ("SMART-D23-DIM-PUSH-DIN", 1.4, "AL_Arial"), ("Art. 033758", 1.3, "AL_Arial"),
      ("Input: DC 12-36V", 1.0, "AL_Arial"), ("Output: DC 4×(12-36)V, 4×5A, 240-720W", 1.0, "AL_Arial"),
      ("Ta: -20...+45°C   IP20", 1.0, "AL_Arial"), ("Made in P.R.C.", 1.0, "AL_Arial")]),
    ("SMART-DALI-104-62-ADDR-DIM-DT6-DIN White", "028422", 110.0, False, "IARL",
     _DALI62_TOP, ["+", "1", "+", "2", "+", "3", "+", "4"], _DALI62_TL, [("LED OUTPUT", 24.3)], "R",
     [("DIMMER", 1.5, "AL_ArialBd"), ("SMART-DALI-104-62-", 1.3, "AL_Arial"), ("ADDR-DIM-DT6-DIN", 1.3, "AL_Arial"),
      ("White (12-24V, 4x5A)", 1.1, "AL_Arial"), ("Art. 028422", 1.1, "AL_ArialBd")]),
    ("SMART-DALI-104-62-ADDR-DIM-DT6-DIN Black", "039147", 110.0, True, "IARL",
     _DALI62_TOP, ["+", "1", "+", "2", "+", "3", "+", "4"], _DALI62_TL, [("LED OUTPUT", 24.3)], "W",
     [("DIMMER", 1.5, "AL_ArialBd"), ("SMART-DALI-104-62-", 1.3, "AL_Arial"), ("ADDR-DIM-DT6-DIN", 1.3, "AL_Arial"),
      ("Black (12-24V, 4x5A)", 1.1, "AL_Arial"), ("Art. 039147", 1.1, "AL_ArialBd")]),
    ("SMART-DALI-104-62-ADDR-RGBW-DT6-DT8-PD-DIN White", "028422(1)", 110.0, False, "IARL",
     _DALI62_TOP, ["+", "R", "+", "G", "+", "B", "+", "W"], _DALI62_TL, [("LED OUTPUT", 24.3)], "R",
     [("DIMMER", 1.5, "AL_ArialBd"), ("SMART-DALI-104-62-", 1.3, "AL_Arial"), ("ADDR-RGBW-DT6/DT8-PD", 1.2, "AL_Arial"),
      ("White (12-48V, 4x5A)", 1.1, "AL_Arial"), ("Art. 028422(1)", 1.1, "AL_ArialBd")]),
    ("SMART-DALI-104-62-ADDR-RGBW-DT6-DT8-PD-DIN Black", "039147(1)", 110.0, True, "IARL",
     _DALI62_TOP, ["+", "R", "+", "G", "+", "B", "+", "W"], _DALI62_TL, [("LED OUTPUT", 24.3)], "W",
     [("DIMMER", 1.5, "AL_ArialBd"), ("SMART-DALI-104-62-", 1.3, "AL_Arial"), ("ADDR-RGBW-DT6/DT8-PD", 1.2, "AL_Arial"),
      ("Black (12-48V, 4x5A)", 1.1, "AL_Arial"), ("Art. 039147(1)", 1.1, "AL_ArialBd")]),
    ("SMART-K3-RGBW", "022493", 115.0, False, "AR",
     [_N] * 8, ["-", "+", "+", "+", "R", "G", "B", "W"], [], [("POWER", 8.2), ("ВЫХОД", 30.0)], "R",
     [("КОНТРОЛЛЕР", 1.5, "AL_ArialBd"), ("SMART-K3-RGBW", 1.3, "AL_Arial"), ("(12-36V, 4x5A,", 1.1, "AL_Arial"),
      ("DIN, 2.4G)", 1.1, "AL_Arial"), ("Арт. 022493", 1.1, "AL_ArialBd")]),
    ("SMART-K36-DMX", "028409", 110.0, False, "AR",
     ["GND", "D-", "D+", "GND", "D-", "D+", _N, _N], ["-", "+", "+", "+", "R", "G", "B", "W"],
     [("DMX IN", 8.2), ("DMX OUT", 24.1)], [("POWER", 8.2), ("ВЫХОД", 30.0)], "R",
     [("ДЕКОДЕР DMX512", 1.4, "AL_ArialBd"), ("SMART-K36-DMX", 1.3, "AL_Arial"), ("(12-24V, 4x5A,", 1.1, "AL_Arial"),
      ("DIN)", 1.1, "AL_Arial"), ("Арт. 028409", 1.1, "AL_ArialBd")]),
    ("SMART-RGBW-DIN", "025169", 115.0, False, "AR",
     ["-", "+", _N, "+", "R", "G", "B", "W"], [_N, _N, _N, "+", "R", "G", "B", "W"],
     [("POWER", 5.5), ("ВХОД", 32.6)], [("ВЫХОД", 32.6)], None,
     [("УСИЛИТЕЛЬ", 1.5, "AL_ArialBd"), ("SMART-RGBW-DIN", 1.4, "AL_Arial"), ("Input: DC 12-36V", 1.0, "AL_Arial"),
      ("Output: 4×5A, 240-720W", 1.0, "AL_Arial"), ("Art. 025169", 1.1, "AL_ArialBd"), ("Made in P.R.C.", 1.0, "AL_Arial")]),
]
SR = [  # имя, арт, вариант, верх, низ, цифр, кнопки, строки
    ("SR-2303DIN-PD", "022539", "LEV", ["D1", "D2", "D1", "D2", "V+", "V-"], ["4", "3", "2", "1", "+"], 2, ["0-6", "0-9"],
     [("SR-2303DIN", 1.6, "AL_ArialBd"), ("PRI: Uin=12-36VDC", 1.0, "AL_Arial"), ("Iin=20.5A", 1.0, "AL_Arial"),
      ("SEC: Uout=4×(12-36)VDC", 1.0, "AL_Arial"), ("Iout=4×5A", 1.0, "AL_Arial"),
      ("TEMP RANGE: -20...+50°C", 0.9, "AL_Arial")]),
    ("SR-2108FA-RJ45-DIN", "021847", "RJ45", ["V+", "V-"], ["W-", "B-", "G-", "R-", "+"], 4,
     ["Back", "Enter", "Up", "Down"], []),
]
DALI504 = [("DALI-504-TE-DIN", "029279")]
DALI82 = [("SMART-DALI-104-82-RGBW-DT6-DT8-RDM-DIN", "058998")]

# ряды листа (как в базе): (подпись ряда, [(имя, функция, модель)])
ROWS = [
    ("ARV-DR 12V / 24V", [(m[0], arv_dr, m) for m in DR[:10]]),
    ("ARV-DR 48V / ARV-DRS", [(m[0], arv_dr, m) for m in DR[10:]]),
    ("ARV-DRP", [(m[0], arv_drp, m) for m in DRP]),
    ("ARV-DRP-Y / KNX / DALI-504", [(m[0], arv_drpy, m) for m in DRPY] + [(m[0], knx72, m) for m in KNX] +
     [(m[0], dali504, m) for m in DALI504]),
    ("SMART / SR", [(m[0], smart, m) for m in SMART] + [(m[0], smart_dali82, m) for m in DALI82] +
     [(m[0], sr, m) for m in SR]),
]

msp = doc.modelspace()
sizes = {}
for _, items in ROWS:
    for name, fn, m in items:
        sizes[name] = fn(doc.blocks.new(name, base_point=(0, 0)), m)
y = 0.0
for title, items in ROWS:
    x = 0.0
    hmax = max(sizes[n][1] for n, _, _ in items)
    t = msp.add_text(title, height=6, dxfattribs={"style": "AL_Arial"})
    t.set_placement((-10, y), align=TA.MIDDLE_RIGHT)
    for name, fn, m in items:
        W, H = sizes[name]
        msp.add_blockref(name, (x, y))
        tt = msp.add_text(name, height=2.2, dxfattribs={"style": "AL_Arial"})
        tt.set_placement((x, y + hmax / 2 + 6), align=TA.BOTTOM_LEFT)
        x += max(W, 30) + 14
    y -= hmax + 40
doc.saveas(OUT)
print("ok", len(sizes), "blocks")
if LIST:
    for n, (W, H) in sizes.items():
        print(f"{n}\t{W}\t{H}")

# ---------------------------------------------------------------- LISP для базы
#   py make_arlight.py lib.dxf lsp <папка с Arlight_DIN.dwg> <файл .lsp>
# Ряды под низом базы (EXTMIN после ZOOM E, минус 20), ячейки 117,2 (156 — для шириной > 97), не длиннее рейки 895,8.
BASE_ROWS = [  # (высота ряда, рейка над низом ряда, имена)
    (155.8, 82.69, [m[0] for m in DR[0:5]]), (155.8, 82.69, [m[0] for m in DR[5:10]]),
    (155.8, 82.69, [m[0] for m in DR[10:15]]), (155.8, 82.69, [m[0] for m in DR[15:20]]),
    (200.0, 100.0, [m[0] for m in DRP[0:5]]), (200.0, 100.0, [m[0] for m in DRP[5:9]]),
    (200.0, 100.0, [m[0] for m in DRPY]),
    (155.8, 82.69, [m[0] for m in KNX] + [DALI504[0][0], DALI82[0][0]] + [m[0] for m in SR]),
    (155.8, 82.69, [m[0] for m in SMART[0:5]]), (155.8, 82.69, [m[0] for m in SMART[5:8]]),
]
if len(sys.argv) > 4 and sys.argv[2] == "lsp":
    d = sys.argv[3].replace("\\", "/").rstrip("/") + "/"
    GX, RX, RY = 2517.4, 2527.54323342981, 1123.489645820265
    L = [";;; Пополнение «База блоков компоновки.dwg»: Arlight — блоки питания на DIN-рейку и 4-канальные диммеры.",
         ";;; Сгенерировано src/make_arlight.py (режим lsp). Ряды под низом базы (T0), рейка — копия «Дин.рейка» ряда VPR.",
         '(setvar "CMDECHO" 0)', '(setvar "OSMODE" 0)', '(setvar "ATTREQ" 0)',
         "(defun cell (x0 y0 x1 y1)",
         "  (entmake (list '(0 . \"LWPOLYLINE\") '(100 . \"AcDbEntity\") '(8 . \"0\") '(100 . \"AcDbPolyline\")",
         "                 '(90 . 4) '(70 . 1) (list 10 x0 y1) (list 10 x1 y1) (list 10 x1 y0) (list 10 x0 y0))))",
         "(defun cap (s x y w)",
         "  (entmake (list '(0 . \"MTEXT\") '(100 . \"AcDbEntity\") '(8 . \"0\") '(100 . \"AcDbMText\")",
         "                 (list 10 x y 0.0) '(40 . 5.0) (cons 41 w) '(71 . 2) '(72 . 1) (cons 1 s) '(7 . \"Standard\"))))",
         "(defun ins (name x y)",
         "  (entmake (list '(0 . \"INSERT\") '(100 . \"AcDbEntity\") '(8 . \"0\") '(100 . \"AcDbBlockReference\")",
         "                 (cons 2 name) (list 10 x y 0.0) '(41 . 1.0) '(42 . 1.0) '(43 . 1.0) '(50 . 0.0))))",
         "(setq rail nil ss (ssget \"X\" '((0 . \"INSERT\"))) i 0)",
         "(while (< i (sslength ss))",
         "  (setq e (ssname ss i) p (cdr (assoc 10 (entget e))))",
         "  (if (and (equal (car p) 2527.543 0.01) (equal (cadr p) 1123.4896 0.01)) (setq rail e))",
         "  (setq i (1+ i)))",
         '(if (null rail) (princ "\\nRAIL-NOT-FOUND\\n"))',
         ";; верх новых рядов — низ чертежа минус 20 (базу дополняют и другие наборы)",
         '(command "_.ZOOM" "_E")', '(setq T0 (- (cadr (getvar "EXTMIN")) 20.0))',
         "(defun Y (d) (+ T0 d))",
         ";; блоки из Arlight_DIN.dwg: вставить, расчленить один раз, лист-каталог удалить",
         f'(command "_.-INSERT" "ALTMP={d}Arlight_DIN.dwg" (list -5000.0 -5000.0) 1 1 0)',
         "(setq mark (entlast))", '(command "_.EXPLODE" mark)',
         "(setq e (entnext mark))", "(while e (setq nx (entnext e)) (entdel e) (setq e nx))",
         '(command "_.-PURGE" "_B" "ALTMP" "_N")']
    top = 0.0                                             # от T0 вниз
    for rh, ry, names in BASE_ROWS:
        yb = top - rh
        yr = yb + ry
        L.append(f'(command "_.COPY" rail "" (list {RX} {RY}) (list {RX} (Y {yr:.4f})))')
        x0 = GX
        for n in names:
            W, H = sizes[n]
            cw = 117.2 if W <= 97 else 156.0
            cx = x0 + cw / 2
            L.append(f"(cell {x0:.3f} (Y {yb:.3f}) {x0 + cw:.3f} (Y {top:.3f}))")
            cap = n
            if len(n) > 30:                                # длинное имя — перенос по дефису у середины (\P в MTEXT)
                k = min((i for i, ch in enumerate(n) if ch == "-"), key=lambda i: abs(i - len(n) / 2))
                cap = n[:k + 1] + "\\\\P" + n[k + 1:]
            L.append(f'(cap "{cap}" {cx:.3f} (Y {top - 3.0:.3f}) {cw - 4.0:.1f})')
            L.append(f'(ins "{n}" {cx - W / 2:.3f} (Y {yr:.4f}))')
            x0 += cw
        assert x0 - GX <= 906.0, (names, x0 - GX)          # рейка 895,8 от X 2527,5
        top = yb
    L += ['(command "_.AUDIT" "_Y")', '(command "_.ZOOM" "_E")', '(princ "\\nADD-ARLIGHT-DONE\\n")', "(princ)"]
    with open(sys.argv[4], "w", encoding="utf-8", newline="\r\n") as f:
        f.write("\n".join(L) + "\n")
    print("lsp", sys.argv[4], "rows", len(BASE_ROWS), "bottom", round(top, 2))
