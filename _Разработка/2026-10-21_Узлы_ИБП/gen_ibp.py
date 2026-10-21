# Генерирует LISP для узлов ИБП HIDEN KC3360H (098 ЖД) в копии однолинейки проекта (запуск в accoreconsole).
# Лист A3 проекта: рамка *U382 + внешняя ссылка «Штамп». Рисуем в пространстве листа, мм.
# Узел 1 — однолинейная схема ГРЩ → ИБП → ЩГП/ЩБП; узел 2 — шкаф АКБ C-32.
import sys
out, target = sys.argv[1], sys.argv[2]

L = []
LAY1, LAY2 = "Узел 1 ИБП", "Узел 2 АКБ"
LAYER = "ЭОМ_Узлы_ИБП"
DASHDOT, DASH = "ACAD_ISO10W100", "ACAD_ISO02W100"
PW, TW = 50, 25          # вес линий: силовые цепи 0,5 мм, прочие 0,25 мм


def enc(t):
    """Символы вне cp1251 — кодом AutoCAD \\U+XXXX."""
    r = []
    for ch in t:
        try:
            ch.encode("cp1251")
            r.append(ch)
        except UnicodeEncodeError:
            r.append("\\U+%04X" % ord(ch))
    return "".join(r)


def s(t):
    return '"' + enc(t).replace("\\", "\\\\").replace('"', '\\"') + '"'


def p2(p):
    return "'(%.3f %.3f 0.0)" % (p[0], p[1])


def common(lw, lt):
    r = "'(67 . 1) (cons 8 %s) (cons 370 %d)" % (s(LAYER), lw)
    if lt:
        r += " (cons 6 %s) '(48 . 1.0)" % s(lt)
    return r


def line(a, b, lw=PW, lt=None):
    L.append("(entmake (list '(0 . \"LINE\") %s (cons 10 %s) (cons 11 %s)))" % (common(lw, lt), p2(a), p2(b)))


def pline(pts, lw=PW, lt=None, closed=False, width=None, bulges=None):
    v = ""
    for i, p in enumerate(pts):
        v += " (cons 10 '(%.3f %.3f))" % p
        if bulges:
            v += " '(42 . %.4f)" % bulges[i]
    L.append("(entmake (list '(0 . \"LWPOLYLINE\") '(100 . \"AcDbEntity\") %s '(100 . \"AcDbPolyline\") '(90 . %d) '(70 . %d)%s%s))"
             % (common(lw, lt), len(pts), 1 if closed else 0, " '(43 . %.3f)" % width if width is not None else "", v))


def path(*pts, lw=PW, lt=None):
    pline(list(pts), lw, lt)


def rect(x0, y0, x1, y1, lw=TW, lt=None):
    pline([(x0, y0), (x1, y0), (x1, y1), (x0, y1)], lw, lt, closed=True)


def circle(c, r, lw=TW):
    L.append("(entmake (list '(0 . \"CIRCLE\") %s (cons 10 %s) '(40 . %.3f)))" % (common(lw, None), p2(c), r))


def dot(c, d=1.2):
    """Точка соединения (ГОСТ 2.702) — залитый кружок."""
    x, y = c
    pline([(x - d / 4, y), (x + d / 4, y)], lw=TW, closed=True, width=d / 2, bulges=[1.0, 1.0])


def term(c):
    """Вывод (клемма) — кружок."""
    circle(c, 0.8)


JUST = {"BL": (0, 0), "BC": (1, 0), "BR": (2, 0), "ML": (0, 2), "MC": (1, 2), "MR": (2, 2), "TL": (0, 3), "TC": (1, 3)}


def text(p, t, h=2.5, just="BL", rot=0.0):
    hj, vj = JUST[just]
    L.append("(entmake (list '(0 . \"TEXT\") '(67 . 1) (cons 8 %s) (cons 10 %s) (cons 11 %s) (cons 40 %.2f) (cons 1 %s) "
             "'(7 . \"SX_GOST\") (cons 50 %.6f) '(72 . %d) '(73 . %d)))"
             % (s(LAYER), p2(p), p2(p), h, s(t), rot * 3.14159265358979 / 180, hj, vj))


def lines(p, rows, h=2.5, step=None, just="BL"):
    step = step or h * 1.6
    for i, t in enumerate(rows):
        text((p[0], p[1] - i * step), t, h, just)


def qf_v(x, yt):
    """Автоматический выключатель (ГОСТ 2.755: замыкающий контакт + функция выключателя «×»), поток вниз.
    Вход — (x, yt), выход — (x, yt-8.7)."""
    line((x - 0.8, yt - 0.8), (x + 0.8, yt + 0.8))
    line((x - 0.8, yt + 0.8), (x + 0.8, yt - 0.8))
    line((x, yt - 8.7), (x - 3.8, yt - 0.6))
    return yt - 8.7


def qf_h(xl, y):
    """То же, поток вправо: вход (xl, y), выход (xl+8.7, y); нож вверх."""
    line((xl - 0.8, y - 0.8), (xl + 0.8, y + 0.8))
    line((xl - 0.8, y + 0.8), (xl + 0.8, y - 0.8))
    line((xl + 8.7, y), (xl + 0.6, y + 3.8))
    return xl + 8.7


def conv(x0, y0, a, top, bot):
    """Преобразователь (ГОСТ 2.721): квадрат с диагональю, род тока входа слева вверху, выхода справа внизу."""
    rect(x0, y0, x0 + a, y0 + a, lw=PW)
    line((x0, y0), (x0 + a, y0 + a), lw=TW)
    for sym, cx, cy in ((top, x0 + a * 0.28, y0 + a * 0.72), (bot, x0 + a * 0.72, y0 + a * 0.28)):
        if sym == "=":
            line((cx - 2, cy + 0.6), (cx + 2, cy + 0.6), lw=TW)
            line((cx - 2, cy - 0.6), (cx + 2, cy - 0.6), lw=TW)
        else:
            text((cx, cy), "~", 3.5, "MC")


def triac(cx, cy):
    """Симистор (ГОСТ 2.730) — статический ключ байпаса, поток вправо, выводы cx±6."""
    line((cx - 6, cy), (cx - 2, cy))
    line((cx + 2, cy), (cx + 6, cy))
    line((cx - 2, cy - 3.2), (cx - 2, cy + 3.2), lw=TW)
    line((cx + 2, cy - 3.2), (cx + 2, cy + 3.2), lw=TW)
    pline([(cx - 2, cy + 0.3), (cx - 2, cy + 3.0), (cx + 2, cy + 1.65)], lw=TW, closed=True)
    pline([(cx + 2, cy - 0.3), (cx + 2, cy - 3.0), (cx - 2, cy - 1.65)], lw=TW, closed=True)
    line((cx + 2, cy - 2.6), (cx + 4.2, cy - 4.8), lw=TW)


def cell_h(cx, y, pitch):
    """Аккумулятор (ГОСТ 2.768): длинная тонкая черта «+» слева, короткая толстая «–» справа; выводы на ±pitch/2."""
    line((cx - 1, y - 4), (cx - 1, y + 4), lw=TW)
    line((cx + 1, y - 2), (cx + 1, y + 2), lw=70)
    line((cx - pitch / 2, y), (cx - 1, y))
    line((cx + 1, y), (cx + pitch / 2, y))


def cell_v(x, cy):
    """Аккумулятор, «+» сверху."""
    line((x - 4, cy + 1), (x + 4, cy + 1), lw=TW)
    line((x - 2, cy - 1), (x + 2, cy - 1), lw=70)


def thermistor(x, y0):
    """Терморезистор (ГОСТ 2.728), вертикально, низ (x, y0), высота 10."""
    rect(x - 2, y0, x + 2, y0 + 10, lw=TW)
    path((x - 5, y0 - 0.5), (x - 3.5, y0 - 0.5), (x + 3.5, y0 + 10.5), lw=TW)
    text((x + 3, y0 + 1), "t°", 2.5, "BL")


def slash(c, n, horiz=False):
    """Число проводников однолинейной линии (ГОСТ 2.702): наклонная черта и цифра."""
    x, y = c
    line((x - 1.2, y - 1.2), (x + 1.2, y + 1.2), lw=TW)
    if horiz:
        text((x + 1.0, y + 1.6), str(n), 2.0, "BL")
    else:
        text((x + 1.8, y - 0.2), str(n), 2.0, "ML")


def set_layout(name):
    L.append('(setvar "CTAB" %s)' % s(name))


# ---------------------------------------------------------------- подготовка файла
L.append('(vl-load-com)')
L.append('(setvar "OSMODE" 0) (setvar "CMDECHO" 0) (setvar "FILEDIA" 0) (setvar "EXPERT" 5)')
L.append('''(foreach lt (list %s %s) (if (not (tblsearch "LTYPE" lt)) (command "_.-LINETYPE" "_Load" lt "acadiso.lin" "")))'''
         % (s(DASHDOT), s(DASH)))
L.append('(if (not (tblsearch "LAYER" %s)) (entmake (list \'(0 . "LAYER") \'(100 . "AcDbSymbolTableRecord") '
         '\'(100 . "AcDbLayerTableRecord") (cons 2 %s) \'(70 . 0) \'(62 . 7) \'(6 . "Continuous"))))' % (s(LAYER), s(LAYER)))
# модель — очистить (там была схема ЩСПЗ-ЖД)
L.append('''(setq ss (ssget "_X" '((410 . "Model"))) i 0)
(while (and ss (< i (sslength ss))) (entdel (ssname ss i)) (setq i (1+ i)))''')
L.append('(command "_.-LAYOUT" "_Rename" "ЩСПЗ-ЖД" %s)' % s(LAY1))
L.append('(command "_.-LAYOUT" "_Rename" "ЩСПЗ-ЖД (2)" %s)' % s(LAY2))


def clean_sheet(lay, title):
    # лишний видовой экран (id ≠ 1), боковик однолинейки; атрибуты рамки; масштаб «—»
    L.append('''(setq ss (ssget "_X" (list (cons 410 %s))) i 0)
(while (and ss (< i (sslength ss)))
  (setq e (ssname ss i) d (entget e))
  (cond ((and (= (cdr (assoc 0 d)) "VIEWPORT") (> (cdr (assoc 69 d)) 1)) (vl-catch-all-apply 'entdel (list e)))
        ((and (= (cdr (assoc 0 d)) "INSERT") (wcmatch (cdr (assoc 2 d)) "_Боковик*")) (vl-catch-all-apply 'entdel (list e)))
        ((and (= (cdr (assoc 0 d)) "MTEXT") (equal (cdr (assoc 10 d)) '(407.5 2.5 0.0) 0.5)) (entdel e))
        ((= (cdr (assoc 0 d)) "INSERT")
          (setq a (entnext e))
          (while (and a (= (cdr (assoc 0 (setq ad (entget a)))) "ATTRIB"))
            (if (= (cdr (assoc 2 ad)) "ОПИСАНИЕ_ЛИСТА") (entmod (subst (cons 1 %s) (assoc 1 ad) ad)))
            (setq a (entnext a)))))
  (setq i (1+ i)))''' % (s(lay), s(title)))
    text((407.5, 2.5), "—", 2.5, "MC")



NOTES_COMMON = [
    "1. Сечение W2 принято по выходному автомату QF2 63 А, а не по номиналу ИБП",
    "   (производитель для 60 кВА рекомендует 25 мм²).",
    "2. Клеммную колодку ИБП, наличие клеммы N для АКБ и тип QFB уточнить",
    "   у поставщика HIDEN.",
    "3. Помещение АКБ: вентиляция, +15...+25 °C, ограниченный доступ,",
    "   нагрузка на пол около 1 т.",
]

# ================================================================ УЗЕЛ 1
set_layout(LAY1)
clean_sheet(LAY1, "Узел 1. Схема подключения ИБП")
text((25, 280), "Узел 1. Схема подключения ИБП (однолинейная)", 5.0)

# ГРЩ (ВРУ)
rect(25, 228, 115, 272, lt=DASHDOT)
text((28, 266.5), "ГРЩ (ВРУ)", 3.5)
lines((28, 262), ["секция I особой категории,", "резерв — ДГУ AKSA AD 630"])
line((30, 248), (110, 248), lw=70)
text((66, 249.5), "3L+N+PE ~380/220 В", 2.5)
dot((60, 248))
line((60, 248), (60, 244))
yb = qf_v(60, 244)
lines((64, 242), ["QF1", "3P+N, C160 А"])
line((60, yb), (60, 205))
slash((60, 222), 5)
lines((64, 219), ["W1  ВВГнг(А)-LS 5х35", "(L1–L3, N, PE)"])

# ИБП A1
rect(35, 100, 250, 205, lt=DASHDOT)
term((60, 205))
text((64, 199), "A1  ИБП HIDEN KC3360H, 60 кВА / 54 кВт", 3.5)
text((64, 194.5), "on-line, двойное преобразование; вход и выход 3L+N+PE ~380/220 В")
line((60, 205), (60, 155))
dot((60, 185))
# байпас (общий вход с выпрямителем)
line((60, 185), (68, 185))
xb = qf_h(68, 185)
lines((66, 191), ["Q2 байпас"])
line((xb, 185), (154, 185))
rect(152, 178, 172, 192)
triac(162, 185)
line((168, 185), (200, 185))
text((162, 174.5), "Статический байпас", 2.5, "TC")
line((200, 185), (200, 155))
dot((200, 155))
# основной канал
line((60, 155), (68, 155))
xb = qf_h(68, 155)
text((66, 161), "Q1 вход")
line((xb, 155), (92, 155))
conv(92, 148, 14, "~", "=")
text((99, 145), "Выпрямитель", 2.5, "TC")
line((106, 155), (124, 155))
conv(124, 148, 14, "=", "~")
text((131, 145), "Инвертор", 2.5, "TC")
line((138, 155), (212, 155))
xb = qf_h(212, 155)
text((210, 161), "Q3 выход")
line((xb, 155), (250, 155))
term((250, 155))
lines((38, 131), ["Вход выпрямителя и байпаса общий,", "нейтраль общая (раздельный", "байпас — опция)"])
lines((38, 107), ["Q1, Q2, Q3 — встроенные выключатели ИБП", "(рис. 3-1 руководства: поз. 23, 22, 24)"])
# звено постоянного тока → АКБ
dot((115, 155))
line((115, 155), (115, 100))
slash((115, 118), 3)
text((119, 116), "±192 В DC (+, N, –)", 2.5, "BL", 90)
term((115, 100))

# порты слаботочки (рис. 3-1): низ корпуса ИБП
PORTS = [(135, "п. 4 — датчик t° АКБ"), (152, "п. 10 — мониторинг QFB"),
         (195, "п. 5, 6 — RS485"), (210, "п. 14 — слот SNMP (опция)"), (225, "п. 17 — EPO")]
for x, t in PORTS:
    term((x, 100))
    text((x + 1.25, 102), t, 2.5, "BL", 90)

# W3 → шкаф АКБ
line((115, 100), (115, 75))
lines((112, 93), ["W3  3х(1х50) + 1х25 (PE), Cu", "марку уточнить; Iразр ≈ 168 А"], just="BR")
slash((115, 80), "3+PE")

# шкаф АКБ A2 (упрощённо, подробно — узел 2)
rect(40, 15, 185, 75, lt=DASHDOT)
text((43, 69), "A2  Шкаф АКБ C-32 (см. узел 2)", 3.5)
yb = qf_v(115, 58)
line((115, 75), (115, 58))
lines((111, 56), ["QFB  3P (+ / N / –), ≈200 А,", "≥250 В DC на полюс", "(тип — уточнить у поставщика)"], just="BR")
line((115, yb), (115, 42))
cell_v(115, 41)
line((115, 39), (115, 37.5), lw=TW, lt=DASH)
line((115, 37.5), (115, 34))
cell_v(115, 33)
line((115, 31), (115, 25))
lines((122, 40), ["GB1...GB32  12 В, 100 А·ч, GEL,", "32 шт. последовательно, ±192 В"])
# датчик температуры и доп. контакт QFB
thermistor(135, 60)
text((138, 66), "BK1", 2.5)
line((135, 70), (135, 100), lw=TW)
text((138.5, 82), "W7", 2.5, "BL", 90)
line((113.1, 53.4), (146, 53.4), lw=TW, lt=DASH)
rect(146, 49, 158, 57)
text((152, 53), "конт.", 2.0, "MC")
text((160, 52), "QFB доп.", 2.5)
line((152, 57), (152, 100), lw=TW)
text((155.5, 82), "W8", 2.5, "BL", 90)
lines((43, 24), ["BK1 — датчик t° АКБ (комплектность — уточнить у поставщика)",
                 "W7, W8 — марку кабеля уточнить"], h=2.5)

# выход → QF2 → W2 → ЩГП/ЩБП
line((250, 155), (262, 155))
xb = qf_h(262, 155)
lines((258, 168), ["QF2  3P+N, 63 А", "(характеристику и место установки уточнить)"])
line((xb, 155), (330, 155))
slash((300, 155), 5, horiz=True)
lines((275, 150), ["W2  ВВГнг(А)-LS 5х16", "(L1–L3, N, PE), см. примеч. 1"])
rect(330, 130, 405, 180, lt=DASHDOT)
text((345, 172), "ЩГП / ЩБП", 3.5)
lines((345, 166), ["Рр ≈ 35 кВт, Кс = 0,81"])
line((338, 140), (338, 170), lw=70)
dot((338, 155))

# слаботочные связи
path((195, 100), (195, 79), (300, 79), lw=TW)
text((213, 80.5), "W4  RS485, экранированная витая пара — марку уточнить")
path((210, 100), (210, 89), (300, 89), lw=TW)
text((213, 90.5), "W5  Ethernet (SNMP) — марку уточнить")
path((225, 100), (225, 96), (285, 96), (285, 107), (300, 107), lw=TW)
text((228, 97.5), "W6  EPO — марку уточнить")
rect(300, 74, 395, 94)
text((303, 84.5), "АСУ (диспетчеризация)", 3.5)
text((303, 77), "RS485 — W4; SNMP — W5 (карта SNMP — опция)", 2.2)
rect(300, 101, 395, 113)
text((303, 108.5), "SB1 — кнопка EPO (аварийное отключение ИБП)", 2.5)
text((303, 104), "место установки и тип контакта уточнить у поставщика", 2.2)

# примечания
text((258, 284), "Примечания", 3.5)
lines((258, 278), NOTES_COMMON + [
    "4. Номера портов (п. ...) — по рис. 3-1 руководства HIDEN KC3300H 50–200 кВА.",
    "5. QF1 — характеристика C, 125 % входного тока ИБП 125 А (руководство п. 3.5).",
    "6. Руководство указывает АКБ типа AGM VRLA; применение GEL",
    "   уточнить у поставщика HIDEN.",
    "7. Марки кабелей W3–W8 уточнить.",
], step=4.4)

# ================================================================ УЗЕЛ 2
set_layout(LAY2)
clean_sheet(LAY2, "Узел 2. Подключение шкафа АКБ")
text((30, 278), "Узел 2. Подключение шкафа АКБ", 5.0)
rect(30, 115, 405, 265, lt=DASHDOT)
text((33, 259), "A2  Шкаф АКБ C-32", 3.5)
text((33, 254), "GB1...GB32 — АКБ GEL 12 В, 100 А·ч, 32 шт. последовательно; двухполярная батарея ±192 В")

P, X0, YB = 10.5, 50.0, 240.0
for k in range(32):
    cx = X0 + k * P
    cell_h(cx, YB, P)
    text((cx, 246), "GB%d" % (k + 1), 2.2, "BC")
XP, XN, XM = X0 - P / 2, X0 + 15.5 * P, X0 + 31 * P + P / 2     # «+», средняя точка N, «–»
text((XP - 2, YB), "+", 3.5, "MR")
text((XM + 2, YB), "–", 3.5, "ML")
dot((XN, YB))
text((60, 230), "Положительная группа GB1...GB16: +192 В отн. N")
text((232, 230), "Отрицательная группа GB17...GB32: –192 В отн. N")
lines((XN + 2.5, 210), ["N — средняя точка:", "«–» GB16 / «+» GB17"])

# QFB 3P: полюса +, N, –
QP, QN, QM, YQ = 150.0, XN, 275.5, 165.0
path((XP, YB), (XP, 180), (QP, 180), (QP, YQ))
line((XN, YB), (XN, YQ))
path((XM, YB), (XM, 180), (QM, 180), (QM, YQ))
for x in (QP, QN, QM):
    yb = qf_v(x, YQ)
    line((x, yb), (x, 115))
    term((x, 115))
line((QP - 1.9, 160.35), (300, 160.35), lw=TW, lt=DASH)
lines((146, 165), ["QFB", "3P (+ / N / –), ≈200 А,", "не менее 250 В DC на полюс", "(вариант — предохранитель-",
                   "разъединитель NH); тип уточнить", "у поставщика"], just="BR", step=3.9)
text((QP + 1.5, 117), "+", 2.5)
text((QN + 1.5, 117), "N", 2.5)
text((QM + 1.5, 117), "–", 2.5)
# доп. контакт QFB → ИБП п. 10
rect(300, 155, 322, 166)
text((311, 160.5), "доп. конт.", 2.0, "MC")
lines((324, 162), ["QFB — доп. контакт", "(НО/НЗ — уточнить)"])
line((311, 155), (311, 86), lw=TW)
text((314.5, 90), "W8", 2.5, "BL", 90)
# PE шкафа
line((330, 125), (365, 125), lw=70)
text((367, 125), "PE (корпус шкафа)", 2.5, "ML")
dot((345, 125))
line((345, 125), (345, 86))
text((348.5, 90), "PE 1х25 (W3)", 2.5, "BL", 90)
# датчик температуры BK1 → ИБП п. 4
thermistor(393, 248)
text((393, 260), "BK1", 2.5, "BC")
line((393, 248), (393, 86), lw=TW)
text((396.5, 120), "W7  датчик t° АКБ (комплектность уточнить)", 2.2, "BL", 90)

# W3 и клеммы ИБП
for x in (QP, QN, QM):
    line((x, 115), (x, 86))
lines((147, 108), ["W3  3х(1х50) Cu (+, N, –)", "+ 1х25 Cu (PE), марку уточнить", "Iразр ≈ 168 А (54 кВт / 0,955 / 336 В)"],
      just="BR", step=4.0)
rect(40, 66, 405, 86)
text((43, 78), "A1  ИБП HIDEN KC3360H", 3.5)
text((43, 71.5), "клеммы батареи +, N, –; PE; порты п. 4, п. 10 (рис. 3-1)", 2.5)
for x, t in ((QP, "+"), (QN, "N"), (QM, "–"), (311, "п. 10"), (345, "PE"), (393, "п. 4")):
    term((x, 86))
    text((x + 1.5, 81.5), t, 2.5)

# примечания
text((22, 56), "Примечания", 3.5)
lines((22, 51), NOTES_COMMON[:2] + NOTES_COMMON[2:] + [
    "4. Средняя точка N — между «–» GB16 и «+» GB17 (руководство п. 3.8).",
    "5. При монтаже одну и более межблочных перемычек не устанавливать до ПНР (руководство п. 3.8).",
    "6. Руководство указывает АКБ типа AGM VRLA; применение GEL уточнить у поставщика.",
], h=2.2, step=3.7)

# ---------------------------------------------------------------- сохранить
L.append('(setvar "CTAB" %s)' % s(LAY1))
L.append('(repeat 3 (command "_.-PURGE" "_All" "*" "_N"))')
L.append('(command "_.SAVEAS" "2018" %s)' % s(target))
L.append('(princ "\\nIBP-DONE\\n")')

open(out, "w", encoding="cp1251").write("\n".join(L) + "\n")
print(len(L), "lines")
