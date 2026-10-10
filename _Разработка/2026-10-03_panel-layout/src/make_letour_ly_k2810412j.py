# Letour LY/K2810412J — релейный актуатор KNX, 28 каналов 10 А. Вид спереди 1:1, мм.
# Габарит 216 x 90 (12 мод., knx24.com); положение элементов и цвета — по фото изделия (knx24.com),
# масштаб фото по габариту: X 216/935 px, Y 90/360 px.
# Точка вставки: левый край, Y=0 — ось DIN-рейки. Блок «LY-K2810412J» (в имени блока «/» недопустим).
import math, sys
import ezdxf
from ezdxf.enums import TextEntityAlignment as TA

OUT = sys.argv[1]
NAME = "LY-K2810412J"
W = 216.0

C_PANEL = (222, 222, 223)       # лицевая крышка
C_BAND = (206, 206, 207)        # ступени с клеммами (рамка)
C_BANDIN = (228, 228, 229)      # ступени, внутреннее поле
C_BANDLO = (204, 205, 205)      # нижняя ступень (в тени крышки)
C_EDGE = (160, 160, 162)        # грани
C_BTN = (58, 58, 60)            # кнопки
C_BTNHI = (110, 110, 114)
C_LED = (188, 188, 190)         # светодиоды (погашены)
C_TEXT = (64, 64, 66)
C_LOGO = (28, 28, 30)
C_GREEN = (150, 212, 162)       # клеммные колодки
C_GREEND = (78, 142, 96)
C_SCREW = (244, 244, 244)
C_SCREWD = (150, 150, 154)
C_SCHEM = (120, 120, 124)       # схема контактов на корпусе
C_KNXR = (204, 40, 48)
C_KNXB = (40, 42, 45)
C_HOLE = (40, 40, 42)

doc = ezdxf.new("R2018", setup=False)
doc.units = ezdxf.units.MM
doc.header["$INSUNITS"] = 4
doc.header["$INSBASE"] = (0, 0, 0)
doc.styles.new("LT_Arial", dxfattribs={"font": "arial.ttf"})
doc.styles.new("LT_ArialBd", dxfattribs={"font": "arialbd.ttf"})
doc.styles.new("LT_ArialBdI", dxfattribs={"font": "arialbi.ttf"})
RAW = len(sys.argv) > 2 and sys.argv[2] == "raw"   # raw: геометрия прямо в модели (для переопределения блока в базе)
blk = doc.modelspace() if RAW else doc.blocks.new(NAME, base_point=(0, 0))
A = {"layer": "0"}


def tc(e, rgb):
    if rgb:
        e.rgb = rgb
    return e


def rr_pts(x0, y0, x1, y1, r):
    b = math.tan(math.pi / 8)
    return [(x0 + r, y0, 0), (x1 - r, y0, b), (x1, y0 + r, 0), (x1, y1 - r, b),
            (x1 - r, y1, 0), (x0 + r, y1, b), (x0, y1 - r, 0), (x0, y0 + r, b)]


def fill_rect(x0, y0, x1, y1, rgb, r=0):
    h = blk.add_hatch(color=8, dxfattribs=dict(A))
    h.rgb = rgb
    if r:
        h.paths.add_polyline_path(rr_pts(x0, y0, x1, y1, r), is_closed=True)
    else:
        h.paths.add_polyline_path([(x0, y0), (x1, y0), (x1, y1), (x0, y1)], is_closed=True)
    return h


def fill_poly(pts, rgb):
    h = blk.add_hatch(color=8, dxfattribs=dict(A))
    h.rgb = rgb
    h.paths.add_polyline_path(pts, is_closed=True)
    return h


def fill_circle(x, y, r, rgb):
    h = blk.add_hatch(color=8, dxfattribs=dict(A))
    h.rgb = rgb
    h.paths.add_edge_path().add_arc((x, y), r, 0, 360)
    return h


def rrect(x0, y0, x1, y1, r, rgb=None):
    return tc(blk.add_lwpolyline(rr_pts(x0, y0, x1, y1, r), format="xyb", close=True, dxfattribs=dict(A)), rgb)


def rect(x0, y0, x1, y1, rgb=None):
    return tc(blk.add_lwpolyline([(x0, y0), (x1, y0), (x1, y1), (x0, y1)], close=True, dxfattribs=dict(A)), rgb)


def line(x0, y0, x1, y1, rgb=None):
    return tc(blk.add_line((x0, y0), (x1, y1), dxfattribs=dict(A)), rgb)


def circle(x, y, r, rgb=None):
    return tc(blk.add_circle((x, y), r, dxfattribs=dict(A)), rgb)


def arc(x, y, r, a0, a1, rgb=None):
    return tc(blk.add_arc((x, y), r, a0, a1, dxfattribs=dict(A)), rgb)


def text(s, x, y, h, align=TA.MIDDLE_CENTER, style="LT_Arial", rgb=C_TEXT, wf=0.9):
    t = blk.add_text(s, height=h, dxfattribs={**A, "style": style, "width": wf})
    t.set_placement((x, y), align=align)
    return tc(t, rgb)


# ---------------------------------------------------------------- геометрия
YT, YM1, YM0, YB = 45.0, 23.5, -22.5, -45.0       # верх, крышка (верх/низ), низ
XB0, XB1 = 0.0, W                                 # ступени с клеммами — во всю ширину, как крышка
PITCH = 7.23                                       # шаг клемм
GROUPS_TOP = [13.85, 82.25, 151.75]                # первый винт группы (по 8)
GROUPS_BOT = [(13.85, 8), (82.25, 8), (151.75, 2)]
YS_TOP, YS_BOT = 36.4, -35.85                      # оси винтов
COLS = [10.17 + 13.043 * k for k in range(16)]     # колонки кнопок
DIV = 181.1                                        # перегородка отсека KNX внизу справа

top_scr = [g + PITCH * i for g in GROUPS_TOP for i in range(8)]
bot_scr = [g + PITCH * i for g, n in GROUPS_BOT for i in range(n)]

# ---------------------------------------------------------------- заливки
fill_rect(XB0, YM1 - 0.5, XB1, YT, C_BAND, r=0.8)
fill_rect(XB0 + 0.6, YM1 + 0.6, XB1 - 0.6, YT - 0.6, C_BANDIN, r=0.5)
fill_rect(XB0, YB, XB1, YM0 + 0.5, C_BAND, r=0.8)
fill_rect(XB0 + 0.6, YB + 0.6, DIV - 0.4, YM0 - 0.6, C_BANDLO, r=0.5)
fill_rect(DIV + 0.4, YB + 0.6, XB1 - 0.6, YM0 - 0.6, C_BANDLO, r=0.5)


def strip(x_first, n, ys):
    """зелёная колодка: перемычка, кольца, торцевые лапки"""
    x0, x1 = x_first - 3.1, x_first + PITCH * (n - 1) + 3.1
    fill_rect(x0 - 1.0, ys - 1.4, x1 + 1.0, ys + 1.4, C_GREEND)       # лапки
    fill_rect(x0, ys - 1.7, x1, ys + 1.7, C_GREEN, r=0.4)
    for i in range(n):
        fill_circle(x_first + PITCH * i, ys, 2.55, C_GREEN)


for g in GROUPS_TOP:
    strip(g, 8, YS_TOP)
for g, n in GROUPS_BOT:
    strip(g, n, YS_BOT)
for x in top_scr:
    fill_circle(x, YS_TOP, 1.85, C_SCREW)
for x in bot_scr:
    fill_circle(x, YS_BOT, 1.85, C_SCREW)

# крышка
fill_rect(0, YM0, W, YM1, C_PANEL, r=1.0)
for x in COLS:
    fill_circle(x, 12.7, 2.05, C_BTN)
    fill_circle(x, 4.4, 1.55, C_LED)
for k in list(range(13)) + [15]:
    fill_circle(COLS[k], -12.5, 2.05, C_BTN)
    fill_circle(COLS[k], -4.1, 1.55, C_LED)

# клеммник KNX в отсеке справа внизу: чёрный −, красный +
fill_rect(195.6, -44.3, 201.4, -32.6, C_KNXB)
fill_rect(201.4, -44.3, 206.8, -32.6, C_KNXR)
fill_rect(206.8, -38.6, 207.8, -34.4, C_KNXR)
fill_circle(204.9, -26.3, 1.0, C_HOLE)            # окно светодиода/отверстие

# ---------------------------------------------------------------- линии
rrect(XB0, YM1 - 0.5, XB1, YT, 0.8)                 # контур — ПоСлою
rrect(XB0, YB, XB1, YM0 + 0.5, 0.8)
rrect(XB0 + 0.6, YM1 + 0.6, XB1 - 0.6, YT - 0.6, 0.5, C_EDGE)
rrect(XB0 + 0.6, YB + 0.6, DIV - 0.4, YM0 - 0.6, 0.5, C_EDGE)
rrect(DIV + 0.4, YB + 0.6, XB1 - 0.6, YM0 - 0.6, 0.5, C_EDGE)
rrect(0, YM0, W, YM1, 1.0)
line(0.6, YM0 + 0.7, W - 0.6, YM0 + 0.7, C_EDGE)    # нижняя кромка крышки

# колодки: контуры колец, винты PZ
for xs, ys in [(top_scr, YS_TOP), (bot_scr, YS_BOT)]:
    for x in xs:
        circle(x, ys, 2.55, C_GREEND)
        circle(x, ys, 1.85, C_SCREWD)
        line(x - 1.15, ys, x + 1.15, ys, C_SCREWD)
        line(x, ys - 1.15, x, ys + 1.15, C_SCREWD)


def schem(xs, ys, first, sgn):
    """схема контактов у клемм: тройка (общий, кан. n, кан. n+1); sgn=+1 — под верхним рядом, −1 — над нижним"""
    yt = ys - sgn * 2.8                 # от кольца
    yc = ys - sgn * 5.2                 # верх контакта
    yo = ys - sgn * 7.0                 # низ контакта
    yb = ys - sgn * 9.1                 # общая шина
    for j in range(0, len(xs) - 2, 3):
        xc, xa, xb = xs[j:j + 3]
        line(xc, yt, xc, yb, C_SCHEM)
        line(xc, yb, xb, yb, C_SCHEM)
        for n, xo in enumerate((xa, xb)):
            line(xo, yt, xo, yc, C_SCHEM)
            line(xo, yo, xo, yb, C_SCHEM)
            line(xo, yo, xo - 1.4, yc + sgn * 0.3, C_SCHEM)    # подвижный контакт
            text(f"S{first + j // 3 * 2 + n}", xo - 3.4, (yc + yo) / 2, 1.4, rgb=C_SCHEM)


schem(top_scr, YS_TOP, 1, +1)
schem(bot_scr, YS_BOT, 17, -1)

# крышка: подписи, кнопки
for k, x in enumerate(COLS):
    circle(x, 12.7, 2.05, (40, 40, 42))
    arc(x, 12.7, 1.5, 100, 190, C_BTNHI)
    text(f"S{k + 1}", x, 8.6, 1.5)
for k in list(range(13)) + [15]:
    x = COLS[k]
    circle(x, -12.5, 2.05, (40, 40, 42))
    arc(x, -12.5, 1.5, 100, 190, C_BTNHI)
    text({12: "FS", 15: "Prog"}.get(k, f"S{k + 17}"), x, -8.4, 1.5)

text("LETOUR", 3.2, 18.45, 3.0, TA.MIDDLE_LEFT, style="LT_ArialBdI", rgb=C_LOGO, wf=1.08)
text("LY/K2810412J", 24.2, 18.3, 2.2, TA.MIDDLE_LEFT, rgb=C_TEXT, wf=1.15)
text("EIB", 193.6, 18.4, 2.1, TA.MIDDLE_LEFT, style="LT_ArialBdI", rgb=C_LOGO, wf=1.1)
arc(202.6, 17.2, 2.6, 45, 135, C_LOGO)              # знак KNX
text("KNX", 202.6, 18.1, 1.6, style="LT_ArialBd", rgb=C_LOGO, wf=1.0)
text("CE", 210.4, 17.7, 3.2, rgb=C_LOGO, wf=0.95)
text("µ", 185.6, -17.9, 2.8, rgb=C_LOGO)
text("10Ax28unit 250V~", 213.0, -17.8, 2.0, TA.MIDDLE_RIGHT, rgb=C_LOGO, wf=1.05)

# отсек KNX: значок фиксатора, клеммник
blk.add_lwpolyline([(184.2, -24.9, 0), (194.3, -24.9, -1), (194.3, -27.7, 0), (184.2, -27.7, 0)],
                   format="xyb", dxfattribs=dict(A)).rgb = C_LOGO
rect(195.6, -44.3, 201.4, -32.6, (20, 20, 22))
blk.add_lwpolyline([(201.4, -44.3), (206.8, -44.3), (206.8, -38.6), (207.8, -38.6), (207.8, -34.4),
                    (206.8, -34.4), (206.8, -32.6), (201.4, -32.6)], close=True, dxfattribs=dict(A)).rgb = (130, 22, 30)
circle(204.9, -26.3, 1.0, (20, 20, 22))

if not RAW:
    doc.modelspace().add_blockref(NAME, (0, 0), dxfattribs={"layer": "0"})
doc.saveas(OUT)
print("ok", len(blk))
