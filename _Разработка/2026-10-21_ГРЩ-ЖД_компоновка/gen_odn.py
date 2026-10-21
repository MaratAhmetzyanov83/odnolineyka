# LISP: дорисовать на однолинейке ГРЩ-ЖД (ручной узел вводов) УЗИП, KV, АВР ПР200, YA1/YA2.
import sys
out = sys.argv[1]
L = []
LAY, ST = "ОДН_РУЧН", "SX_GOST"


def s(t):
    return '"' + t.replace("\\", "\\\\").replace('"', '\\"') + '"'


def line(a, b, lt=None):
    L.append("(entmake (list '(0 . \"LINE\") (cons 8 %s) %s(cons 10 '(%.3f %.3f 0.0)) (cons 11 '(%.3f %.3f 0.0))))"
             % (s(LAY), ("(cons 6 %s) " % s(lt)) if lt else "", a[0], a[1], b[0], b[1]))


def rect(x0, y0, x1, y1):
    for a, b in (((x0, y0), (x1, y0)), ((x1, y0), (x1, y1)), ((x1, y1), (x0, y1)), ((x0, y1), (x0, y0))):
        line(a, b)


def dot(p, r=0.6):
    L.append("(entmake (list '(0 . \"LWPOLYLINE\") '(100 . \"AcDbEntity\") (cons 8 %s) '(100 . \"AcDbPolyline\") '(90 . 2) '(70 . 1) '(43 . %.2f)"
             " (cons 10 '(%.3f %.3f)) '(42 . 1.0) (cons 10 '(%.3f %.3f)) '(42 . 1.0)))" % (s(LAY), r, p[0] - r / 2, p[1], p[0] + r / 2, p[1]))


def text(p, t, h=2.5, just="ML"):
    hj = {"ML": 0, "MC": 1, "MR": 2}[just]
    L.append("(entmake (list '(0 . \"TEXT\") (cons 8 %s) (cons 7 %s) (cons 10 '(%.3f %.3f 0.0)) (cons 11 '(%.3f %.3f 0.0)) (cons 40 %.2f) (cons 1 %s) (cons 72 %d) '(73 . 2)))"
             % (s(LAY), s(ST), p[0], p[1], p[0], p[1], h, s(t), hj))


def spd(xin, xb, ytap, tags, side):
    """Отвод от линии ввода xin на УЗИП: точка ytap, ветвь по x = xb вниз."""
    dot((xin, ytap))
    line((xin, ytap), (xb, ytap))
    line((xb, ytap), (xb, 1500))
    rect(xb - 1.25, 1492, xb + 1.25, 1500)            # FU
    line((xb, 1492), (xb, 1484))
    rect(xb - 2.0, 1474, xb + 2.0, 1484)              # FV
    line((xb, 1482.5), (xb, 1476))                    # стрелка разряда
    line((xb, 1476), (xb - 0.9, 1478))
    line((xb, 1476), (xb + 0.9, 1478))
    line((xb, 1474), (xb, 1468))
    for w, dy in ((3.0, 0.0), (2.0, 1.0), (1.0, 2.0)):  # заземление
        line((xb - w, 1468 - dy), (xb + w, 1468 - dy))
    tx = xb - 3.5 if side == "L" else xb + 3.5
    j = "MR" if side == "L" else "ML"
    text((tx, 1497.5), tags[0], 2.2, j)
    text((tx, 1494.5), "gG 80А", 2.0, j)
    text((tx, 1480.5), tags[1], 2.2, j)
    text((tx, 1477.5), "тип 1+2", 2.0, j)


L.append('(setvar "CMDECHO" 0) (setvar "OSMODE" 0)')
spd(125.0, 103.0, 1506.0, ("FU1-FU3", "FV1-FV3"), "L")
spd(265.0, 287.0, 1506.0, ("FU4-FU6", "FV4-FV6"), "R")


def ya(xq, tag):
    x0 = xq - 9.5
    rect(x0, 1521.0, x0 + 2.5, 1524.5)
    line((x0 + 2.5, 1522.75), (xq - 1.3, 1522.75), "HIDDEN2")
    text((x0 + 1.25, 1519.3), tag, 1.8, "MC")
    text((x0 + 1.25, 1517.0), "от СПС", 1.5, "MC")


ya(565.0, "YA1")
ya(605.0, "YA2")
text((195.0, 1457.5), "FU, FV - УЗИП тип 1+2 с предохр. gG 80А", 2.5, "ML")

# правка существующих надписей слоя ОДН_РУЧН
REPL = [("U1", "KV1"), ("U2", "KV2"),
        ("контроллер по ТУ", "A1 ПР200-24, A2 ПРМ-24.1"),
        ("изготовителя НКУ", "ИБП 24 В: G1, GB1"),
        ("U1, U2 - контроль напряжения вводов", "KV1, KV2 - РНПП-311М (SF5, SF6 3P C2)")]
L.append("(setq rp (list %s))" % " ".join("(cons %s %s)" % (s(a), s(b)) for a, b in REPL))
L.append('''(setq ss (ssget "_X" (list '(0 . "TEXT") (cons 8 "ОДН_РУЧН"))) i 0 n 0)
(while (< i (sslength ss))
  (setq e (ssname ss i) d (entget e) a (assoc (cdr (assoc 1 d)) rp))
  (if a (progn (entmod (subst (cons 1 (cdr a)) (assoc 1 d) d)) (setq n (1+ n))))
  (setq i (1+ i)))
(princ (strcat "\\nODN-REPL=" (itoa n) "\\n"))''')
L.append('(princ "\\nODN-DONE\\n")')
open(out, "w", encoding="cp1251").write("\n".join(L) + "\n")
print(len(L), "lines")
