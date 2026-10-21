# Генерирует LISP для узлов ИБП HIDEN KC3360H (098 ЖД) в копии однолинейки проекта (запуск в accoreconsole).
# Лист A3 проекта: рамка *U382 + внешняя ссылка «Штамп». Рисуем в пространстве листа, мм.
# Узел 1 — однолинейная схема ГРЩ → ИБП → ЩГП/ЩБП; узел 2 — шкаф АКБ C-32.
import sys
out, target = sys.argv[1], sys.argv[2]

L = []
LAY1, LAY2 = "Узел ИБП", None
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



L.append('(command "_.-LAYOUT" "_Delete" "ЩСПЗ-ЖД (2)")')

# ================================================================ лист: упрощённая схема + кабели
set_layout(LAY1)
clean_sheet(LAY1, "Схема подключения ИБП. Кабели")
text((25, 280), "Схема подключения ИБП HIDEN KC3360H (упрощённая)", 5.0)

# ЩГП
rect(25, 140, 125, 265, lt=DASHDOT)
text((28, 259), "ЩГП", 3.5)
text((28, 254.5), "питание от ГРЩ-ЖД; I особая категория")
line((45, 212), (45, 245), lw=70)
text((43, 247), "шина «Сеть»", 2.5, "BL")
dot((45, 230))
line((45, 230), (80, 230))
xb = qf_h(80, 230)
lines((76, 241), ["QF1 3P C63 А", "«На ИБП»"])
line((xb, 230), (125, 230))
line((45, 160), (45, 200), lw=70)
lines((28, 155), ["шина «ИБП» — через переключатель", "сервисного байпаса I-0-II 4P 80 А"], h=2.2)
dot((45, 190))
line((45, 190), (80, 190))
xb = qf_h(80, 190)
lines((76, 201), ["QF2 3P C63 А", "«От ИБП»"])
line((xb, 190), (125, 190))

# ИБП
rect(170, 170, 260, 250, lw=PW)
text((215, 236), "A1  ИБП HIDEN KC3360H", 3.5, "MC")
text((215, 230), "60 кВА / 54 кВт, on-line", 2.5, "MC")
text((215, 225.5), "вход / выход 3L+N+PE ~380/220 В", 2.5, "MC")
line((125, 230), (170, 230))
text((128, 231.5), "W1  5х16", 2.5)
line((170, 190), (125, 190))
text((128, 191.5), "W2  5х16", 2.5)

# шкаф АКБ
rect(170, 80, 260, 135, lw=PW)
text((215, 124), "A2  Шкаф АКБ C-32", 3.5, "MC")
text((215, 118), "32 х 12 В, 100 А·ч, GEL; ±192 В", 2.5, "MC")
text((215, 113.5), "QFB 3P (+ / N / –), ≈200 А, ≥250 В DC", 2.5, "MC")
text((215, 109), "N — между «–» АКБ 16 и «+» АКБ 17", 2.5, "MC")
line((195, 170), (195, 135))
text((193, 152), "W3, W3.1", 2.5, "BR")
line((230, 170), (230, 135), lw=TW)
text((232, 152), "W7", 2.5)
line((248, 170), (248, 135), lw=TW)
text((250, 152), "W8", 2.5)

# АСУ, кнопка EPO
rect(315, 215, 405, 245)
text((360, 233), "АСУ (диспетчеризация)", 3.0, "MC")
text((360, 226), "RS485 — W4; SNMP — W5", 2.2, "MC")
line((260, 238), (315, 238), lw=TW)
text((263, 239.5), "W4  RS485")
line((260, 222), (315, 222), lw=TW)
text((263, 223.5), "W5  SNMP (опция)")
rect(315, 172, 405, 195)
text((360, 186), "SB1 — кнопка EPO", 3.0, "MC")
text((360, 179), "аварийное отключение ИБП", 2.2, "MC")
line((260, 183), (315, 183), lw=TW)
text((263, 184.5), "W6  EPO")

# таблица кабелей
COLS = [("Поз.", 11), ("Откуда", 30), ("Куда", 31), ("Марка", 42), ("Сечение", 26), ("L, м", 11), ("Примечание", 54)]
ROWS = [
    ("W1", "ЩГП, QF1", "ИБП A1, вход", "ВВГнг(А)-LS", "5х16", "16", "по книге ЩГП"),
    ("W2", "ИБП A1, выход", "ЩГП, QF2", "ВВГнг(А)-LS", "5х16", "16", "по книге ЩГП; см. примеч. 1"),
    ("W3", "Шкаф АКБ, QFB", "ИБП A1, клеммы АКБ", "ПуГВнг(А)-LS", "3х(1х50)", "", "+, N, –; Iразр ≈ 168 А"),
    ("W3.1", "Шкаф АКБ, PE", "ИБП A1, PE", "ПуГВнг(А)-LS", "1х25", "", "жёлто-зелёный"),
    ("W4", "ИБП A1, RS485", "АСУ", "КИПЭВнг(А)-LS", "1х2х0,60", "", ""),
    ("W5", "ИБП A1, SNMP", "АСУ", "U/UTP кат. 5e нг(А)-LS", "4х2х0,52", "", "если заказана карта SNMP"),
    ("W6", "ИБП A1, EPO", "Кнопка SB1", "КВВГнг(А)-LS", "4х1,0", "", ""),
    ("W7", "ИБП A1, датчик t°", "Шкаф АКБ", "—", "—", "", "в комплекте ИБП (проверить)"),
    ("W8", "ИБП A1, мон. QFB", "Шкаф АКБ, доп. конт. QFB", "КВВГнг(А)-LS", "4х1,0", "", ""),
]
TX, TY, RH = 22.0, 62.0, 5.2
xs = [TX]
for _, w in COLS:
    xs.append(xs[-1] + w)
n = len(ROWS) + 1
text((TX, TY + 2), "Кабели узла ИБП (длины — по плану)", 3.0)
for i in range(n + 1):
    line((xs[0], TY - i * RH), (xs[-1], TY - i * RH), lw=PW if i in (0, 1, n) else TW)
for x in xs:
    line((x, TY), (x, TY - n * RH), lw=PW if x in (xs[0], xs[-1]) else TW)
for j, (h, _) in enumerate(COLS):
    text(((xs[j] + xs[j + 1]) / 2, TY - RH / 2), h, 2.2, "MC")
for i, r in enumerate(ROWS):
    yc = TY - (i + 1.5) * RH
    for j, t in enumerate(r):
        if t:
            if j in (0, 4, 5):
                text(((xs[j] + xs[j + 1]) / 2, yc), t, 2.0, "MC")
            else:
                text((xs[j] + 1.0, yc), t, 2.0, "ML")

# примечания
text((272, 160), "Примечания", 3.5)
lines((272, 154), [
    "1. Сечение W2 принято по выходному автомату QF2 63 А, а не по",
    "   номиналу ИБП (производитель для 60 кВА рекомендует 25 мм²).",
    "2. Клеммную колодку ИБП, наличие клеммы N для АКБ и тип QFB",
    "   уточнить у поставщика HIDEN.",
    "3. Помещение АКБ: вентиляция, +15...+25 °C, ограниченный доступ,",
    "   нагрузка на пол около 1 т.",
    "4. QF1 63 А и W1 5х16 — по нагрузке ЩГП (≈35 кВт). На полную",
    "   мощность 60 кВА руководство требует C160 А и 35 мм².",
    "5. Марки W3–W8 приняты проектом; длины — по плану.",
], step=4.4)

L.append('(setvar "CTAB" %s)' % s(LAY1))
L.append('(repeat 3 (command "_.-PURGE" "_All" "*" "_N"))')
L.append('(command "_.SAVEAS" "2018" %s)' % s(target))
L.append('(princ "\nIBP-DONE\n")')
open(out, "w", encoding="cp1251").write("\n".join(L) + "\n")
print(len(L), "lines")
