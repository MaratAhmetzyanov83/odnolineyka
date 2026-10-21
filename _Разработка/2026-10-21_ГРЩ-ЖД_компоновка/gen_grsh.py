# Генерирует LISP для компоновки ГРЩ-ЖД в копии «Компоновки.dwg» (запуск в accoreconsole).
import sys
out = sys.argv[1]

FX, FY = 31000.0, 3033.48          # низ-лево рамки в модели (масштаб 1:10)
SC = 10.0
CAB = [(FX + 450, FY + 1851), (FX + 1450, FY + 1851)]   # низ корпуса CQE (без цоколя)

# образцы с листа ЩОВ-001: (handle, точка вставки)
SRC_CAB = ("13E977", (6157.958076956492, 3986.963122293474))
SRC_PLINTH = ("13E987", (6157.958076956515, 3885.96330784117))
SRC_PL_A = ("13E99B", (6207.958076956522, 5565.715218564964))   # пара пластин под крупный аппарат
SRC_PL_B = ("13E9AE", (6206.358079674494, 5644.573750216012))
SRC_DPL = ("13E9FA", (6207.958076956522, 4532.968302298534))    # пластина под DIN
SRC_DIN = ("13EA0D", (6209.55807423855, 4570.46830378865))      # DIN-рейка
SRC_FRAME = ("13E954", (5441.762355350729, 3033.480591588035))  # рамка 1:8
CAB0 = SRC_CAB[1]

# габариты блоков: extmin x, extmin y, w, h
BB = {
    "A$C5c1f3258": (-74646.28, -967.68, 247.0, 240.0),
    "NM8N-400": (0.0, -127.5, 140.0, 255.0),
    "NM8N-250": (-0.8, -108.3, 106.6, 201.39),
    "VPR-16-GY": (0.0, -37.7, 12.2, 75.4), "VPR-16-BU": (0.0, -37.7, 12.2, 75.4), "VPR-16-PE-YG": (0.0, -37.7, 12.2, 75.4),
    "VPR-4-GY": (0.0, -27.95, 6.2, 55.9), "VPR-4-BU": (0.0, -27.95, 6.2, 55.9), "VPR-4-PE-YG": (0.0, -27.95, 6.2, 55.9),
    "КБР-400": (0.0, -48.5, 48.0, 97.0),
    "S203": (0.0, -19.0, 54.0, 85.0), "S201": (0.0, -19.0, 18.0, 85.0), "DS201": (0.0, -64.0, 36.0, 83.0),
}

L = []  # строки LISP


def s(t):
    return '"' + t.replace("\\", "\\\\").replace('"', '\\"') + '"'


def pt(p):
    return "'(%.3f %.3f 0.0)" % p


def copy(src, dst):
    h, p0 = src
    L.append('(command "_.COPY" (handent "%s") "" %s %s)' % (h, pt(p0), pt(dst)))


def ins(name, cx, cy, sx=1.0, sy=1.0, layer="devices"):
    """Вставка блока центром габарита в (cx, cy)."""
    ex, ey, w, h = BB[name]
    x = cx - (ex + w / 2) * sx
    y = cy - (ey + h / 2) * sy
    L.append("(entmake (list '(0 . \"INSERT\") (cons 8 %s) (cons 2 %s) (cons 10 '(%.3f %.3f 0.0)) (cons 41 %.4f) (cons 42 %.4f) '(43 . 1.0)))"
             % (s(layer), s(name), x, y, sx, sy))
    return w * sx, h * sy


def text(p, t, h=25.0, rot=0.0, just="ML", layer="0"):
    hj, vj = {"ML": (0, 2), "MC": (1, 2), "BL": (0, 1), "BC": (1, 1), "MR": (2, 2)}[just]
    L.append("(entmake (list '(0 . \"TEXT\") (cons 8 %s) (cons 10 %s) (cons 11 %s) (cons 40 %.2f) (cons 1 %s) '(7 . \"Standard\") (cons 50 %.6f) '(72 . %d) '(73 . %d)))"
             % (s(layer), pt(p), pt(p), h, s(t), rot * 3.14159265358979 / 180, hj, vj))


def line(a, b, layer="0"):
    L.append("(entmake (list '(0 . \"LINE\") (cons 8 %s) (cons 10 %s) (cons 11 %s)))" % (s(layer), pt(a), pt(b)))


def rect(x0, y0, x1, y1, layer="0"):
    L.append("(entmake (list '(0 . \"LWPOLYLINE\") '(100 . \"AcDbEntity\") (cons 8 %s) '(100 . \"AcDbPolyline\") '(90 . 4) '(70 . 1) "
             "(cons 10 '(%.3f %.3f)) (cons 10 '(%.3f %.3f)) (cons 10 '(%.3f %.3f)) (cons 10 '(%.3f %.3f))))"
             % (s(layer), x0, y0, x1, y0, x1, y1, x0, y1))


def label(px, py, lines, h=25.0):
    """Подпись в 1-2 строки: низ первой строки над точкой (px, py)."""
    for i, t in enumerate(reversed(lines)):
        text((px, py + i * h * 1.45), t, h, just="BL")


# ---------- рамка, корпуса ----------
L.append('(vl-load-com)')
L.append('(setvar "3DOSMODE" 1) (setvar "AUTOSNAP" 0)')
L.append('(setvar "OSMODE" 0) (setvar "CMDECHO" 0) (setvar "ATTREQ" 0)')
copy(SRC_FRAME, (FX, FY))
L.append('(command "_.SCALE" (entlast) "" %s 1.25)' % pt((FX, FY)))

ROW_A, ROW_B, ROW_C, ROW_D, ROW_E = 1578.75, 1198.75, 900.0, 650.0, 450.0   # от низа корпуса: пластины / DIN


def cabinet(cx, cy, din_rows):
    dx, dy = cx - CAB0[0], cy - CAB0[1]
    copy(SRC_CAB, (cx, cy))
    copy(SRC_PLINTH, (SRC_PLINTH[1][0] + dx, SRC_PLINTH[1][1] + dy))
    for r in (ROW_A, ROW_B):
        sh = r - (SRC_PL_A[1][1] - CAB0[1])
        copy(SRC_PL_A, (SRC_PL_A[1][0] + dx, SRC_PL_A[1][1] + dy + sh))
        copy(SRC_PL_B, (SRC_PL_B[1][0] + dx, SRC_PL_B[1][1] + dy + sh))
    for r in din_rows:
        sh = r - (SRC_DPL[1][1] - CAB0[1])
        copy(SRC_DPL, (SRC_DPL[1][0] + dx, SRC_DPL[1][1] + dy + sh))
        copy(SRC_DIN, (SRC_DIN[1][0] + dx, SRC_DIN[1][1] + dy + sh))


cabinet(*CAB[0], (ROW_C, ROW_D, ROW_E))
cabinet(*CAB[1], (ROW_C, ROW_D, ROW_E))

PL_MID = lambda cy, r: cy + r + 77.0          # середина пары пластин
DIN_Y = lambda cy, r: cy + r + 37.5           # ось DIN-рейки


def big(name, cx, cy, sx, sy, short):
    w, h = ins(name, cx, cy, sx, sy, layer="0")
    text((cx + w / 2 + 14, cy), short[0], 25, just="ML")
    return w, h


def pline(pts, width=6.0, layer="devices", ltype=None):
    lt = "(cons 6 %s) " % s(ltype) if ltype else ""
    L.append("(entmake (append (list '(0 . \"LWPOLYLINE\") '(100 . \"AcDbEntity\") (cons 8 %s) %s'(100 . \"AcDbPolyline\") (cons 90 %d) '(70 . 0) (cons 43 %.2f))"
             " (mapcar '(lambda (p) (cons 10 p)) '(%s))))"
             % (s(layer), lt, len(pts), width, " ".join("(%.3f %.3f)" % q for q in pts)))


def dot(p, r=7.0):
    L.append("(entmake (list '(0 . \"CIRCLE\") '(8 . \"devices\") (cons 10 '(%.3f %.3f 0.0)) (cons 40 %.2f)))" % (p[0], p[1], r))


BUS_Y = [1865.0, 1895.0, 1925.0]      # оси шин L1, L2, L3 от низа корпуса


def busbars(c, sec):
    for k, yy in enumerate(BUS_Y):
        rect(c[0] + 60, c[1] + yy - 5, c[0] + 940, c[1] + yy + 5, layer="devices")
        text((c[0] + 945, c[1] + yy), "L%d" % (k + 1), 12, just="ML")
    text((c[0] + 430, c[1] + 1840), "L1-L3: ШМТ 60х8, секция %d" % sec, 14, just="ML")


def flex(x, y0, y1, width=6.0):
    pline([(x, y0), (x, y1)], width)


POLES = (-70.0, 0.0, 70.0)            # оси полюсов NM8N-400 х1,5


def din_row(x0, y, items, gap=0.0):
    """items: (имя блока, sx, подпись) — подпись вертикально над аппаратом."""
    x = x0
    for name, sx, lab in items:
        w = BB[name][2] * sx
        ins(name, x + w / 2, y, sx, 1.0)
        if lab:
            text((x + w / 2, y + 62), lab, 14, rot=90, just="ML")
        x += w + gap
    return x


def terminals(x0, y, names, tag, pitch=None):
    x = x0
    for n in names:
        w = BB[n][2]
        ins(n, x + w / 2, y, layer="CP_DEVICES 0")
        x += (pitch or w)
    text(((x0 + x) / 2, y + BB[names[0]][3] / 2 + 25), tag, 22, just="MC")
    return x


def box(x0, y, w, h, tag, lines):
    rect(x0, y - h / 2, x0 + w, y + h / 2, layer="devices")
    for i, t in enumerate(lines):
        if t:
            text((x0 + w / 2, y + 15 - i * 30), t, (14 if i else 18) if w >= 60 else 9, just="MC", layer="devices")
    if w < 60:
        text((x0 + w / 2, y + 62), tag, 14, rot=90, just="ML")
    else:
        text((x0 + w / 2, y + h / 2 + 25), tag, 22 if len(tag) < 5 else 17, just="MC")
    return x0 + w


T16 = ["VPR-16-GY"] * 3 + ["VPR-16-BU", "VPR-16-PE-YG"]
T4 = ["VPR-4-GY"] * 3 + ["VPR-4-BU", "VPR-4-PE-YG"]
M3, M1 = ("S203", 1.0), ("S201", 1.0)
M3W = ("S203", 1.5)    # NXB-125 3P, 27 мм на полюс

# ---------- шкаф 1: ввод 1, секция 1 ----------
cx, cy = CAB[0]
busbars((cx, cy), 1)
big("NM8N-400", cx + 290, PL_MID(cy, ROW_A), 1.5, 1.08, ["QF1"])
big("NM8N-400", cx + 760, PL_MID(cy, ROW_A), 1.0, 1.0, ["QF24"])
big("NM8N-400", cx + 290, PL_MID(cy, ROW_B), 1.5, 1.1, ["QS1"])
y = DIN_Y(cy, ROW_C)
x = terminals(cx + 90, y, T16, "XT1")
x = terminals(x + 40, y, T4, "XT4", pitch=8.0)
x = box(x + 60, y, 36, 85, "KV1", ["U1", "РНПП"])
x = box(x + 40, y, 36, 85, "KV2", ["U2", "РНПП"])
x = box(x + 40, y, 81, 90, "FU1-FU3", ["gG 80А"])
x = box(x + 20, y, 144, 90, "FV1-FV3", ["УЗИП 1+2"])
x = din_row(x + 20, y, [(*M3, "SF5"), (*M3, "SF6")], gap=9.0)
y = DIN_Y(cy, ROW_D)
sec1 = [(*M3W, "QF3"), (*M3, "QF4"), (*M3, "QF5"), (*M3, "QF6"), (*M3, "QF7"), (*M3, "QF8"),
        (*M3, "QF9"), (*M3W, "QF10")]
x = din_row(cx + 90, y, sec1, gap=9.0)
x = box(x - 9, y, 18, 85, "YA1", [""])
x = din_row(x + 9, y, [(*M3, "QF11"), (*M3W, "QF12")], gap=9.0)
x = box(x - 9, y, 18, 85, "YA2", [""])
# АВР: питание цепей управления от двух вводов, контроллер, режим, сигнализация
y = DIN_Y(cy, ROW_E)
x = din_row(cx + 90, y, [("S201", 2.0, "SF1"), ("S201", 2.0, "SF2")], gap=9.0)
x = box(x + 20, y, 36, 85, "KM1", ["KM"])
x = box(x + 9, y, 36, 85, "KM2", ["KM"])
x = din_row(x + 20, y, [("S201", 1.0, "SF3"), ("S201", 1.0, "SF4")], gap=9.0)
x = box(x + 20, y, 123, 90, "A1", ["ПР200-24", "ОВЕН"])
x = box(x + 9, y, 72, 90, "A2", ["ПРМ-24", "ОВЕН"])
x = box(x + 20, y, 36, 85, "SA1", ["А-0-Р"])
for k in range(3):
    x = box(x + (20 if k == 0 else 4), y, 18, 85, "HL%d" % (k + 1), [""])
x = terminals(x + 25, y, ["VPR-4-GY"] * 24, "XT10", pitch=5.4)

# ---------- шкаф 2: ввод 2, секция 2 ----------
cx, cy = CAB[1]
busbars((cx, cy), 2)
big("NM8N-400", cx + 290, PL_MID(cy, ROW_A), 1.5, 1.08, ["QF2"])
big("NM8N-250", cx + 760, PL_MID(cy, ROW_A), 1.0, 1.0, ["QF14"])
big("NM8N-400", cx + 290, PL_MID(cy, ROW_B), 1.5, 1.1, ["QS2"])
y = DIN_Y(cy, ROW_C)
x = terminals(cx + 90, y, T16, "XT2")
x = terminals(x + 40, y, T16, "XT13")
x = terminals(x + 40, y, ["КБР-400"] * 5, "XT14", pitch=56.0)
x = box(x + 15, y, 81, 90, "FU4-FU6", ["gG 80А"])
x = box(x + 15, y, 144, 90, "FV4-FV6", ["УЗИП 1+2"])
y = DIN_Y(cy, ROW_D)
sec2 = [(*M3, "QF13"), (*M3, "QF15"), (*M3, "QF16"), (*M3, "QF17"), (*M1, "QF18"), (*M1, "QF19"),
        ("DS201", 1.0, "QFD20"), (*M1, "QF21"), (*M3, "QF22"), (*M1, "QF23")]
din_row(cx + 90, y, sec2, gap=9.0)
y = DIN_Y(cy, ROW_E)
x = box(cx + 90, y, 108, 90, "G1", ["БП+ИБП", "24 В"])
x = box(x + 20, y, 180, 90, "GB1", ["АКБ 24 В"])

# ---------- шины N и PE (общие, по низу обоих шкафов) ----------
for i, (tag, yy) in enumerate((("N", 330.0), ("PE", 220.0))):
    for c in CAB:
        rect(c[0] + 130, c[1] + yy - 15, c[0] + 930, c[1] + yy + 15, layer="devices")
        for k in range(12):
            L.append("(entmake (list '(0 . \"CIRCLE\") '(8 . \"devices\") (cons 10 '(%.3f %.3f 0.0)) '(40 . 5.0)))"
                     % (c[0] + 165 + k * 65, c[1] + yy))
    for c in CAB:
        text((c[0] + 80, c[1] + yy), tag, 25, just="ML")

# ---------- силовые связи (гибкая шина 10х32х1) и вводы ----------
A_MID = PL_MID(0, ROW_A); B_MID = PL_MID(0, ROW_B)
QF_TOP, QF_BOT = A_MID + 255 * 1.08 / 2, A_MID - 255 * 1.08 / 2
QS_TOP, QS_BOT = B_MID + 255 * 1.1 / 2, B_MID - 255 * 1.1 / 2
for n, c in enumerate(CAB):
    X = c[0] + 290
    for k, dx in enumerate(POLES):
        flex(X + dx, c[1] + QS_TOP, c[1] + QF_BOT)                  # QS -> QF
        flex(X + dx, c[1] + QF_TOP, c[1] + BUS_Y[k])                # QF -> шины
        dot((X + dx, c[1] + BUS_Y[k]))
    # ввод: кабели снизу к QS
    pline([(c[0] + 38, c[1] + 40), (c[0] + 38, c[1] + QS_BOT - 40), (X, c[1] + QS_BOT - 40), (X, c[1] + QS_BOT)], 10.0)
    text((c[0] + 62, c[1] + 30), "Ввод %d: 2х(5х150) снизу" % (n + 1), 14, rot=90, just="ML")
    text((X + 110, c[1] + (QS_TOP + QF_BOT) / 2), "гибкая шина 10х32х1", 13, just="ML")
# QF24: вход от шин секции 1, выход на шины секции 2
X24 = CAB[0][0] + 760
for k, dx in enumerate((-47.0, 0.0, 47.0)):
    flex(X24 + dx, CAB[0][1] + A_MID + 127.5, CAB[0][1] + BUS_Y[k])
    dot((X24 + dx, CAB[0][1] + BUS_Y[k]))
    yy = CAB[0][1] + A_MID - 127.5 - 30 - k * 16
    xx = CAB[1][0] + 30 + k * 14
    pline([(X24 + dx, CAB[0][1] + A_MID - 127.5), (X24 + dx, yy), (xx, yy), (xx, CAB[1][1] + BUS_Y[k])], 6.0)
    dot((xx, CAB[1][1] + BUS_Y[k]))
text((X24 - 80, CAB[0][1] + A_MID - 127.5 - 100), "QF24 -> шины секции 2", 13, just="ML")
# QF14: от шин секции 2
X14 = CAB[1][0] + 760
for k, dx in enumerate((-35.0, 0.0, 35.0)):
    flex(X14 + dx, CAB[1][1] + A_MID + 100.7, CAB[1][1] + BUS_Y[k], 5.0)
    dot((X14 + dx, CAB[1][1] + BUS_Y[k]))
# перемычки N и PE между шкафами
for yy in (330.0, 220.0):
    rect(CAB[0][0] + 930, CAB[0][1] + yy - 6, CAB[1][0] + 130, CAB[1][1] + yy + 6, layer="devices")
text((CAB[0][0] + 140, CAB[0][1] + 330 + 30), "N: ШМТ 60х8", 14, just="ML")
text((CAB[0][0] + 140, CAB[0][1] + 220 + 30), "PE: ШМТ 40х5 (ГЗШ)", 14, just="ML")

# подписи шкафов
for c, t in zip(CAB, ("Шкаф 1. Ввод 1, секция 1", "Шкаф 2. Ввод 2, секция 2")):
    text((c[0] + 500, c[1] + 2000 + 45), t, 30, just="BC")

# размеры
L.append('(setvar "DIMSCALE" 10.0)')
L.append('(command "_.DIMLINEAR" "_non" %s "_non" %s "_non" %s)' % (pt((CAB[0][0], CAB[0][1] + 2000)), pt((CAB[1][0] + 1000, CAB[1][1] + 2000)),
                                               pt((CAB[0][0] + 1000, CAB[0][1] + 2000 + 160))))
L.append('(command "_.DIMLINEAR" "_non" %s "_non" %s "_non" %s)' % (pt((CAB[0][0], CAB[0][1] + 2000)), pt((CAB[0][0] + 1000, CAB[0][1] + 2000)),
                                               pt((CAB[0][0] + 500, CAB[0][1] + 2000 + 100))))
L.append('(command "_.DIMLINEAR" "_non" %s "_non" %s "_non" %s)' % (pt((CAB[1][0], CAB[1][1] + 2000)), pt((CAB[1][0] + 1000, CAB[1][1] + 2000)),
                                               pt((CAB[1][0] + 500, CAB[1][1] + 2000 + 100))))
L.append('(command "_.DIMLINEAR" "_non" %s "_non" %s "_T" "2100" "_non" %s)' % (pt((CAB[1][0] + 1000, CAB[1][1] + 2000)), pt((CAB[1][0] + 1000, CAB[1][1] - 101)),
                                               pt((CAB[1][0] + 1130, CAB[1][1] + 1000))))
L.append('(setvar "DIMSCALE" 1.0)')

# ---------- перечень аппаратов ----------
SQ = "\\U+00B2"   # «²» кодом AutoCAD: в cp1251 его нет
ROWS = [
    ("Поз.", "Наименование (по спецификации, раздел 1)", "Кол."),
    ("QS1, QS2", "Выключатель-разъединитель NM8NSD-800, 3P, 800 А", "2"),
    ("QF1, QF2", "Выключатель NM8N-630, 3P, LSI (Ir=520 А, tsd=0,2 с), OF/SD, MX", "2"),
    ("QF24", "Выключатель секционный NM8N-400, 3P, LSI (Ir=400 А, tsd=0,1 с), OF/SD", "1"),
    ("M1-M3", "Электропривод MOD к NM8N, ~230 В (QF1, QF2, QF24)", "3"),
    ("A1", "Контроллер АВР ОВЕН ПР200-24 (=24 В), программа АВР", "1"),
    ("A2", "Модуль расширения ОВЕН ПРМ-24.1", "1"),
    ("KV1, KV2", "Реле контроля напряжения РНПП-311М (вводы 1, 2)", "2"),
    ("SF1, SF2", "NB1-63H 2P C6 - питание цепей АВР от ввода 1 / 2", "2"),
    ("KM1, KM2", "Контактор NCH8-20/20 - АВР цепей управления", "2"),
    ("SF3-SF6", "NB1-63H 1P C4 (G1), C10 (M1-M3); 3P C2 (KV1, KV2)", "4"),
    ("G1, GB1", "Источник питания 24 В с ИБП и АКБ", "1"),
    ("SA1, HL1-HL3", "Переключатель «Авто-0-Ручн.»; лампы ND9", "1+3"),
    ("FU1-FU6, FV1-FV6", "Предохранители gG 80 А; УЗИП тип 1+2, 3P+N", "2 к."),
    ("QF3, QF10, QF12", "NXB-125, 3P, C80 / C100 / C80", "3"),
    ("QF4-QF9, QF11, QF13,", "NB1-63H, 3P, C20...C63", "12"),
    ("QF15-QF17, QF22", "", ""),
    ("QF18, QF19, QF21, QF23", "NB1-63H, 1P, C6...C25", "4"),
    ("QF14", "NM8N-250, 3P, 250 А", "1"),
    ("QFD20", "АВДТ 2P C32 30 мА", "1"),
    ("YA1, YA2", "Независимый расцепитель 230 В к QF10, QF12 (СПС)", "2"),
    ("XT1, XT2, XT13 / XT4 / XT10", "Клеммы 16 / 4 / 2,5 мм" + SQ, "15/5/24"),
    ("XT14", "Колодка КБР-400 (кабель 1х185)", "5"),
    ("", "Корпус CQE 2000х1000х600, цоколь 100. Шины, провода, наконечники - спецификация", "2 шк."),
]
TX0, TY_TOP, RH = FX + 300, FY + 1620, 40.0
CW = [640, 1700, 180]
xs = [TX0, TX0 + CW[0], TX0 + CW[0] + CW[1], TX0 + sum(CW)]
n = len(ROWS)
text((TX0, TY_TOP + 25), "Перечень аппаратов ГРЩ-ЖД. Алгоритм АВР — см. принципиальную схему 0,4 кВ, лист ЖД", 26, just="BL")
for i in range(n + 1):
    line((xs[0], TY_TOP - i * RH), (xs[3], TY_TOP - i * RH))
for xx in xs:
    line((xx, TY_TOP), (xx, TY_TOP - n * RH))
for i, r in enumerate(ROWS):
    yc = TY_TOP - i * RH - RH / 2
    for j, t in enumerate(r):
        if t:
            if j == 2 or i == 0:
                text(((xs[j] + xs[j + 1]) / 2, yc), t, 20, just="MC")
            else:
                text((xs[j] + 12, yc), t, 20, just="ML")

# ---------- лист ----------
L.append('(command "_.-LAYOUT" "_Copy" "ЩОВ-001" "ГРЩ-ЖД")')
L.append('(setvar "CTAB" "ГРЩ-ЖД") (command "_.PSPACE")')
L.append('''(setq ss (ssget "_X" '((0 . "VIEWPORT") (410 . "ГРЩ-ЖД"))) i 0)
(while (< i (sslength ss))
  (setq e (ssname ss i) d (entget e))
  (if (> (cadr (assoc 12 d)) 1000.0)
    (progn (setq d (subst (cons 12 '(%.3f %.3f 0.0)) (assoc 12 d) d))
           (setq d (subst (cons 45 %.3f) (assoc 45 d) d))
           (entmod d)))
  (setq i (1+ i)))''' % (FX + 156.4 * SC, FY + 210.0 * SC, 410.0 * SC))
L.append('''(setq ss (ssget "_X" '((0 . "MTEXT") (410 . "ГРЩ-ЖД"))) i 0)
(while (and ss (< i (sslength ss)))
  (setq e (ssname ss i) d (entget e))
  (if (vl-string-search "1:8" (cdr (assoc 1 d))) (entmod (subst (cons 1 "М1:10") (assoc 1 d) d)))
  (setq i (1+ i)))''')
L.append('''(setq ss (ssget "_X" '((0 . "INSERT") (410 . "ГРЩ-ЖД"))) i 0)
(while (and ss (< i (sslength ss)))
  (setq e (entnext (ssname ss i)))
  (while (and e (= (cdr (assoc 0 (setq d (entget e)))) "ATTRIB"))
    (cond ((= (cdr (assoc 2 d)) "ОПИСАНИЕ_ЛИСТА") (entmod (subst (cons 1 "Схема расположения оборудования в ГРЩ-ЖД") (assoc 1 d) d)))
          ((member (cdr (assoc 2 d)) '("ЛИСТ_№" "ЛИСТ№")) (entmod (subst (cons 1 "108") (assoc 1 d) d))))
    (setq e (entnext e)))
  (setq i (1+ i)))''')
L.append('(princ "\\nGRSH-DONE\\n")')

open(out, "w", encoding="cp1251").write("\n".join(L) + "\n")
print(len(L), "lines")
