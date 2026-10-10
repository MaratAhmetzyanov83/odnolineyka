# KNX-20E-640 (Mean Well), вид спереди 1:1, мм.
# Геометрия — даташит KNX-20E-640-SPEC (Case No. KNX-20), цвета и шильдик — по фото изделия.
# Точка вставки: левый край, Y=0 — ось DIN-рейки (как у блоков Zennio в базе). Габарит 52.5 x 90.
# Результат: блок «KNX-20E-640» и одно его вхождение в 0,0.
import math, sys
import ezdxf
from ezdxf.enums import TextEntityAlignment as TA

OUT = sys.argv[1]
NAME = "KNX-20E-640"
W, H = 52.5, 90.0
DY = -45.0                      # y чертежа Mean Well (0 внизу) -> ось рейки

# цвета (RGB) по фото
C_BODY = (46, 46, 50)           # корпус, антрацит
C_EDGE = (82, 82, 88)           # грани ступеней (блик)
C_SLOT = (20, 20, 22)           # вентиляционные щели
C_HOLE = (14, 14, 16)           # гнёзда клемм
C_LABEL = (74, 76, 81)          # шильдик
C_STRIPE = (204, 32, 46)        # красные полосы шильдика
C_MW = (222, 60, 30)            # логотип MEAN WELL
C_KNXR = (208, 40, 48)          # клеммник KNX, красная часть
C_KNXB = (30, 30, 32)           # клеммник KNX, чёрная часть
C_SCREW = (92, 92, 98)          # винты
C_LED = (238, 238, 238)         # окна светодиодов
C_PRINT = (232, 232, 232)       # белая печать на корпусе
C_WHITE = (255, 255, 255)

doc = ezdxf.new("R2018", setup=False)
doc.units = ezdxf.units.MM
doc.header["$INSUNITS"] = 4
doc.header["$INSBASE"] = (0, 0, 0)
doc.styles.new("MW_Arial", dxfattribs={"font": "arial.ttf"})
doc.styles.new("MW_ArialBd", dxfattribs={"font": "arialbd.ttf"})
doc.styles.new("MW_ArialBdI", dxfattribs={"font": "arialbi.ttf"})
doc.styles.new("MW_Times", dxfattribs={"font": "times.ttf"})
doc.styles.new("MW_TimesBd", dxfattribs={"font": "timesbd.ttf"})
blk = doc.blocks.new(NAME, base_point=(0, 0))
A = {"layer": "0"}              # всё на 0: контур ПоСлою (берёт слой вставки), остальное — истинные цвета


def P(x, y):
    return (x, y + DY)


def tc(e, rgb):
    if rgb:
        e.rgb = rgb
    return e


def rr_pts(x0, y0, x1, y1, r):
    """скруглённый прямоугольник: (x, y, bulge)"""
    b = math.tan(math.pi / 8)
    pts = [(x0 + r, y0, 0), (x1 - r, y0, b), (x1, y0 + r, 0), (x1, y1 - r, b),
           (x1 - r, y1, 0), (x0 + r, y1, b), (x0, y1 - r, 0), (x0, y0 + r, b)]
    return [(x, y + DY, bb) for x, y, bb in pts]


def fill(path_xyb, rgb):
    """сплошная заливка истинным цветом; path — [(x,y)] в координатах MW или [(x,y,bulge)] уже со сдвигом"""
    h = blk.add_hatch(color=250, dxfattribs=dict(A))
    h.rgb = rgb
    if len(path_xyb[0]) == 3:
        h.paths.add_polyline_path(path_xyb, is_closed=True)
    else:
        h.paths.add_polyline_path([P(*p) for p in path_xyb], is_closed=True)
    return h


def fill_rect(x0, y0, x1, y1, rgb, r=0):
    if r:
        return fill(rr_pts(x0, y0, x1, y1, r), rgb)
    return fill([(x0, y0), (x1, y0), (x1, y1), (x0, y1)], rgb)


def fill_circle(x, y, r, rgb):
    h = blk.add_hatch(color=250, dxfattribs=dict(A))
    h.rgb = rgb
    h.paths.add_edge_path().add_arc(P(x, y), r, 0, 360)
    return h


def pline(pts, rgb=None, close=False):
    return tc(blk.add_lwpolyline([P(*p) for p in pts], close=close, dxfattribs=dict(A)), rgb)


def rrect(x0, y0, x1, y1, r, rgb=None):
    return tc(blk.add_lwpolyline(rr_pts(x0, y0, x1, y1, r), format="xyb", close=True, dxfattribs=dict(A)), rgb)


def rect(x0, y0, x1, y1, rgb=None):
    return pline([(x0, y0), (x1, y0), (x1, y1), (x0, y1)], rgb, close=True)


def line(x0, y0, x1, y1, rgb=None):
    return tc(blk.add_line(P(x0, y0), P(x1, y1), dxfattribs=dict(A)), rgb)


def circle(x, y, r, rgb=None):
    return tc(blk.add_circle(P(x, y), r, dxfattribs=dict(A)), rgb)


def arc(x, y, r, a0, a1, rgb=None):
    return tc(blk.add_arc(P(x, y), r, a0, a1, dxfattribs=dict(A)), rgb)


def text(s, x, y, h, align=TA.MIDDLE_CENTER, style="MW_Arial", rgb=C_PRINT, wf=0.9):
    t = blk.add_text(s, height=h, dxfattribs={**A, "style": style, "width": wf})
    t.set_placement(P(x, y), align=align)
    return tc(t, rgb)


def slots(y0, y1, x0, x1, w=1.8, pitch=3.0):
    """ряд вентиляционных щелей"""
    n = int((x1 - x0 + pitch - w) // pitch)
    s = x0 + ((x1 - x0) - (n * pitch - (pitch - w))) / 2
    for i in range(n):
        fill_rect(s + i * pitch, y0, s + i * pitch + w, y1, C_SLOT, r=0.25)


# ================= заливки (рисуются первыми, закрывают рейку и короба)
fill(rr_pts(0, 0, W, H, 0.8), C_BODY)

# вентиляционные щели на ступенях корпуса (ряды как на чертеже MW)
slots(88.3, 89.3, 1.5, W - 1.5)
slots(76.6, 77.5, 17.5, W - 1.5)
slots(66.0, 66.8, 1.5, W - 1.5)
slots(22.8, 23.7, 1.5, W - 1.5)
slots(14.0, 14.9, 1.5, W - 1.5)
slots(1.0, 1.7, 1.5, 15.5)
slots(1.0, 1.7, 37.0, W - 1.5)

# шильдик
WIN = (1.5, 26.3, 50.9, 63.6)
fill_rect(*WIN, C_LABEL, r=0.6)
fill_rect(2.1, 61.6, 50.3, 62.8, C_STRIPE)
fill_rect(2.1, 27.1, 50.3, 28.3, C_STRIPE)

# клеммник шины KNX: красный +, чёрный −
RED = [(3.9, 77.0), (9.8, 77.0), (9.8, 87.5), (3.9, 87.5), (3.9, 82.6), (3.0, 82.6), (3.0, 79.6), (3.9, 79.6)]
BLK = [(9.8, 77.0), (15.8, 77.0), (15.8, 87.5), (9.8, 87.5)]
fill(RED, C_KNXR)
fill(BLK, C_KNXB)
fill_rect(15.8, 79.4, 16.5, 82.8, C_KNXB)
for x0, c in ((4.6, (150, 25, 32)), (7.1, (150, 25, 32)), (10.5, C_HOLE), (13.0, C_HOLE)):
    fill_rect(x0, 85.6, x0 + 1.7, 86.9, c)                       # гнёзда проводов
fill_rect(5.3, 78.0, 8.3, 79.2, (150, 25, 32))                     # окна защёлок
fill_rect(11.6, 78.0, 14.6, 79.2, C_HOLE)

# доп. выход +V/−V (гнездо «восьмёрка») и кнопка Reset
TY = 82.55
for x in (32.75, 37.75):
    fill_circle(x, TY, 2.2, C_HOLE)
fill_rect(32.75, TY - 1.2, 37.75, TY + 1.2, C_HOLE)
fill_circle(45.0, TY, 2.0, C_HOLE)
fill_circle(45.0, TY, 1.3, (36, 36, 40))

# сетевые клеммы L, N, PE
TERM = (13.44, 20.94, 28.44)
for x in TERM:
    fill_circle(x, 7.45, 2.5, C_HOLE)
    fill_circle(x, 7.45, 1.6, C_SCREW)

# светодиоды
LED = [(58.4, "ON"), (51.4, "Reset"), (44.3, "I>Imax")]
for y, _ in LED:
    fill_rect(42.4, y - 0.85, 46.7, y + 0.85, C_LED)

# логотип MEAN WELL
fill_rect(2.8, 50.7, 9.6, 57.3, C_MW)

# ================= линии
rrect(0, 0, W, H, 0.8)                                            # контур — ПоСлою
for x in (0.6, W - 0.6):
    line(x, 0.8, x, H - 0.8, C_EDGE)
for y in (89.2, 76.4, 75.8, 67.5, 66.9, 25.3, 22.7, 13.9, 13.3, 0.8):
    line(0.6, y, W - 0.6, y, C_EDGE)
rrect(WIN[0] - 0.5, WIN[1] - 0.5, WIN[2] + 0.5, WIN[3] + 0.5, 0.9, C_SLOT)  # ниша шильдика

# клеммник KNX — контуры
pline(RED, (120, 20, 26), close=True)
pline(BLK, (70, 70, 74), close=True)
rect(15.8, 79.4, 16.5, 82.8, (70, 70, 74))
text("+V", 6.9, 73.6, 1.9)
text("-V", 12.8, 73.6, 1.9)
text("BUS", 6.9, 70.3, 1.4)
text("30VDC", 13.0, 70.3, 1.4)

# доп. выход, Reset
text("+V", 32.75, 86.6, 1.6, style="MW_ArialBd")
text("-V", 37.75, 86.6, 1.6, style="MW_ArialBd")
text("Reset", 45.0, 86.6, 1.6, style="MW_ArialBd")
circle(45.0, TY, 1.3, (60, 60, 66))

# сетевые клеммы: шлицы PZ, надписи
for x in TERM:
    line(x - 1.0, 7.45, x + 1.0, 7.45, C_HOLE)
    line(x, 6.45, x, 8.45, C_HOLE)
    circle(x, 7.45, 2.5, C_EDGE)
text("L", TERM[0], 3.3, 1.8)
text("N", TERM[1], 3.3, 1.8)
gx, gy = TERM[2], 4.4                                              # знак заземления
line(gx, gy, gx, gy - 0.9, C_PRINT)
for i, w in enumerate((2.2, 1.4, 0.6)):
    line(gx - w / 2, gy - 0.9 - 0.5 * i, gx + w / 2, gy - 0.9 - 0.5 * i, C_PRINT)

# ================= шильдик
# логотип: «MW» и плашка MEAN WELL
text("MW", 6.2, 55.0, 3.0, style="MW_ArialBdI", rgb=C_WHITE, wf=1.0)
rect(3.3, 51.2, 9.1, 52.6, C_WHITE)
text("MEAN WELL", 6.2, 51.9, 0.85, style="MW_ArialBd", rgb=C_WHITE, wf=0.8)
text("KNX-20E-640", 10.4, 54.0, 3.8, TA.MIDDLE_LEFT, style="MW_TimesBd", rgb=C_WHITE, wf=0.85)

# светодиоды — окна и подписи справа
for y, s in LED:
    rect(42.4, y - 0.85, 46.7, y + 0.85, (120, 120, 125))
    text(s, 47.0, y, 1.0, TA.MIDDLE_LEFT, style="MW_TimesBd", rgb=C_WHITE, wf=0.65)

# структурная схема: вход ~ / выход =, дроссель, выходы
bx0, by0, bx1, by1 = 13.6, 36.6, 21.6, 44.6
rect(bx0, by0, bx1, by1, C_WHITE)
line(bx0, by1, bx1, by0, C_WHITE)
text("~", bx0 + 2.2, by0 + 2.2, 2.4, rgb=C_WHITE)
line(bx1 - 3.6, by1 - 2.0, bx1 - 1.6, by1 - 2.0, C_WHITE)          # «=»
line(bx1 - 3.6, by1 - 2.8, bx1 - 1.6, by1 - 2.8, C_WHITE)
line(11.8, 40.6, bx0, 40.6, C_WHITE)                                # вход
rect(23.0, 42.6, 25.6, 45.6, C_WHITE)                               # дроссель
rect(23.7, 43.3, 24.9, 44.9, C_WHITE)
line(bx1, 44.1, 23.0, 44.1, C_WHITE)
line(25.6, 45.1, 26.4, 45.1, C_WHITE)
line(26.4, 45.1, 26.4, 47.2, C_WHITE)
line(26.4, 47.2, 29.6, 47.2, C_WHITE)                               # -> KNX BUS
line(25.6, 43.2, 29.6, 43.2, C_WHITE)                               # -> I1
line(bx1, 39.2, 29.6, 39.2, C_WHITE)                                # -> I2
text("KNX BUS 30VDC", 30.0, 47.2, 1.25, TA.MIDDLE_LEFT, style="MW_Times", rgb=C_WHITE, wf=0.8)
text("I1", 30.0, 43.9, 1.25, TA.MIDDLE_LEFT, style="MW_Times", rgb=C_WHITE, wf=0.8)
text("30VDC", 30.0, 42.3, 1.25, TA.MIDDLE_LEFT, style="MW_Times", rgb=C_WHITE, wf=0.8)
text("I2", 30.0, 39.2, 1.25, TA.MIDDLE_LEFT, style="MW_Times", rgb=C_WHITE, wf=0.8)
text("I1+I2<640mA", 30.0, 37.0, 1.25, TA.MIDDLE_LEFT, style="MW_Times", rgb=C_WHITE, wf=0.8)
text("200-240VAC", 7.2, 42.4, 1.25, style="MW_Times", rgb=C_WHITE, wf=0.8)
text("0.25A", 7.2, 40.6 - 0.0, 1.25, style="MW_Times", rgb=C_WHITE, wf=0.8)
text("50-60Hz", 7.2, 38.8, 1.25, style="MW_Times", rgb=C_WHITE, wf=0.8)

# знаки: корпус с предохранителем, «для помещений», трансформатор, KNX, CE, сертификат
rect(3.2, 30.6, 6.2, 34.6, C_WHITE)
line(3.2, 31.6, 4.7, 30.0, C_WHITE)
line(6.2, 31.6, 4.7, 30.0, C_WHITE)
circle(4.7, 32.6, 0.9, C_WHITE)
pline([(7.4, 30.2), (11.0, 30.2), (11.0, 32.8), (9.2, 34.6), (7.4, 32.8)], C_WHITE, close=True)
circle(14.2, 32.4, 1.9, C_WHITE)
line(12.9, 31.1, 15.5, 33.7, C_WHITE)
arc(22.0, 30.6, 3.6, 40, 140, C_WHITE)                              # дуга знака KNX
text("KNX", 22.0, 31.6, 2.2, style="MW_ArialBd", rgb=C_WHITE, wf=1.0)
text("CE", 29.6, 32.3, 3.8, style="MW_Arial", rgb=C_WHITE, wf=0.95)
rect(33.6, 30.0, 37.0, 34.8, C_WHITE)
pline([(34.0, 32.0), (36.6, 32.0), (35.3, 34.3)], C_WHITE, close=True)
for i in range(4):
    line(37.8, 34.4 - 0.9 * i, 40.6, 34.4 - 0.9 * i, C_WHITE)
text("MEAN WELL ENTERPRISE CO., LTD   MADE IN CHINA", 26.2, 29.0, 0.8, style="MW_Times", rgb=C_WHITE, wf=0.85)

# ================= вхождение в 0,0
doc.modelspace().add_blockref(NAME, (0, 0), dxfattribs={"layer": "0"})
doc.saveas(OUT)
print("ok", len(blk))
