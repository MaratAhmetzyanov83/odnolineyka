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



PL_MID = lambda cy, r: cy + r + 77.0
ROW_A = 1578.75
L.append('(setvar "OSMODE" 0) (setvar "CMDECHO" 0) (setvar "3DOSMODE" 1)')
# ---------- электроприводы MOD на лицевой стороне QF1, QF2, QF24 ----------
def motor(cx, cy, w, h, tag):
    x0, y0, x1, y1 = cx - w / 2, cy - h / 2 + 10, cx + w / 2, cy + h / 2 + 10
    L.append('(command "_.WIPEOUT" "_non" %s "_non" %s "_non" %s "_non" %s "")'
             % (pt((x0, y0)), pt((x1, y0)), pt((x1, y1)), pt((x0, y1))))
    rect(x0, y0, x1, y1, layer="devices")
    rect(cx - 18, y0 + 18, cx + 18, y0 + 44, layer="devices")          # окно положения
    text((cx, y1 - 28), tag, 24, just="MC", layer="devices")
    text((cx, y1 - 58), "MOD", 17, just="MC", layer="devices")
    text((cx, y1 - 82), "эл. привод", 13, just="MC", layer="devices")


motor(CAB[0][0] + 290, CAB[0][1] + PL_MID(0, ROW_A), 170, 150, "M1")
motor(CAB[1][0] + 290, CAB[1][1] + PL_MID(0, ROW_A), 170, 150, "M2")
motor(CAB[0][0] + 760, CAB[0][1] + PL_MID(0, ROW_A), 116, 140, "M3")


L.append('(princ "\nMOT-DONE\n")')
open(out, 'w', encoding='cp1251').write('\n'.join(L) + '\n')
print(len(L))
