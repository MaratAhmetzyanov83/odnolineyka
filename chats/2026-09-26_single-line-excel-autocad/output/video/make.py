"""Видеоинструкция: таблица «Однолинейка» -> схема в AutoCAD. Кадры PIL -> ffmpeg (H.264)."""
import math, subprocess, sys
from PIL import Image, ImageDraw, ImageFont

W, H, FPS = 1920, 1080, 25
OUT = sys.argv[1] if len(sys.argv) > 1 else "instr.mp4"
PREVIEW = "--preview" in sys.argv

F = "/usr/share/fonts/truetype/"
def font(name, size): return ImageFont.truetype(F + name, size)
CAL = lambda s: font("crosextra/Carlito-Regular.ttf", s)
CALB = lambda s: font("crosextra/Carlito-Bold.ttf", s)
DJ = lambda s: font("dejavu/DejaVuSans.ttf", s)
DJB = lambda s: font("dejavu/DejaVuSans-Bold.ttf", s)
MONO = lambda s: font("dejavu/DejaVuSansMono.ttf", s)

# палитра (цвета листа «Однолинейка»)
BG = (236, 239, 243); INK = (32, 38, 46); MUTED = (98, 108, 120)
C_IN = (255, 242, 204); C_RES = (226, 239, 218); C_SRC = (242, 242, 242); C_HEAD = (221, 235, 247)
C_GRID = (208, 212, 218); ACC = (230, 110, 20); ACC2 = (31, 110, 190); OK = (0, 97, 0); BAD = (156, 0, 6)
EXCEL_GREEN = (33, 115, 70)
KIND = {"in": C_IN, "res": C_RES, "src": C_SRC, "hid": (250, 250, 250)}


def ease(x): x = max(0.0, min(1.0, x)); return x * x * (3 - 2 * x)
def lerp(a, b, t): return a + (b - a) * t
def mix(c1, c2, t): return tuple(int(lerp(a, b, t)) for a, b in zip(c1, c2))


def text_w(d, s, f): return d.textlength(s, font=f)


def wrap(d, s, f, maxw):
    out = []
    for para in s.split("\n"):
        line = ""
        for w in para.split(" "):
            t = (line + " " + w).strip()
            if text_w(d, t, f) <= maxw: line = t
            else:
                if line: out.append(line)
                line = w
        out.append(line)
    return out


# ------------------------------------------------------------------ общие элементы кадра
def frame_base(step, title, sub=None):
    im = Image.new("RGB", (W, H), BG); d = ImageDraw.Draw(im)
    d.rectangle([0, 0, W, 92], fill=(26, 32, 40))
    if step:
        d.rounded_rectangle([40, 22, 40 + 70, 70], 10, fill=ACC)
        d.text((75, 46), str(step), font=DJB(30), fill="white", anchor="mm")
        d.text((135, 46), title, font=DJB(34), fill="white", anchor="lm")
    else:
        d.text((40, 46), title, font=DJB(34), fill="white", anchor="lm")
    d.text((W - 40, 46), "Однолинейка ЩР ЭОМ · SX_Schema", font=DJ(20), fill=(150, 160, 172), anchor="rm")
    if sub: d.text((40, 118), sub, font=DJ(26), fill=MUTED, anchor="lm")
    return im, d


def caption(d, lines, t, t0=0.4):
    """нижняя плашка с пояснением (появляется плавно)"""
    a = ease((t - t0) / 0.5)
    if a <= 0: return
    f = DJ(32)
    ls = []
    for s in lines: ls += wrap(d, s, f, W - 220)
    hbox = 40 + 46 * len(ls)
    y0 = H - 40 - hbox
    col = mix(BG, (255, 255, 255), a)
    d.rounded_rectangle([60, y0, W - 60, H - 40], 18, fill=col, outline=mix(BG, (200, 205, 212), a), width=2)
    d.rectangle([60, y0 + 14, 68, H - 54], fill=mix(BG, ACC, a))
    for i, s in enumerate(ls):
        d.text((100, y0 + 22 + 46 * i), s, font=f, fill=mix(BG, INK, a))


def cursor(d, x, y, click=0.0):
    pts = [(x, y), (x, y + 34), (x + 9, y + 26), (x + 16, y + 41), (x + 22, y + 38), (x + 15, y + 24), (x + 26, y + 24)]
    if click > 0:
        r = 10 + 26 * click
        d.ellipse([x - r, y - r, x + r, y + r], outline=mix(ACC, BG, click), width=4)
    d.polygon(pts, fill="white", outline=(0, 0, 0))


def path_pos(keys, t):
    """keys: [(t, x, y)] -> позиция курсора"""
    if t <= keys[0][0]: return keys[0][1:]
    for (t0, x0, y0), (t1, x1, y1) in zip(keys, keys[1:]):
        if t0 <= t <= t1:
            k = ease((t - t0) / max(1e-6, t1 - t0)); return lerp(x0, x1, k), lerp(y0, y1, k)
    return keys[-1][1:]


# ------------------------------------------------------------------ таблица «как в Excel»
class Grid:
    def __init__(self, x, y, cols, rows, first_row=21, rh=46, hh=78, fs=24, maxw=None, smax=1.55):
        tot = 56 + sum(c[2] for c in cols)
        sc = min(smax, ((maxw or (W - 2 * x)) / tot))
        self.sc = sc
        self.x, self.y, self.cols, self.rows, self.first = x, y, cols, rows, first_row
        self.rh, self.hh, self.fs, self.lw, self.th = rh * sc, hh * sc, int(fs * sc), 56 * sc, 34 * sc
        self.cx = {}
        cx = x + self.lw
        for c in cols:
            self.cx[c[0]] = (cx, cx + c[2] * sc); cx += c[2] * sc
        self.right = cx

    def cell(self, col, i):
        x0, x1 = self.cx[col]; y0 = self.y + self.th + self.hh + i * self.rh
        return x0, y0, x1, y0 + self.rh

    def center(self, col, i):
        x0, y0, x1, y1 = self.cell(col, i); return (x0 + x1) / 2, (y0 + y1) / 2

    def draw(self, d, over=None, typed=None, glow=None):
        over = over or {}; typed = typed or {}; glow = glow or {}
        x, y = self.x, self.y
        # строка букв столбцов
        sc, th, lw = self.sc, self.th, self.lw
        d.rectangle([x, y, self.right, y + th], fill=(230, 232, 235))
        for c in self.cols:
            x0, x1 = self.cx[c[0]]
            d.text(((x0 + x1) / 2, y + th / 2), c[0], font=CAL(int(20 * sc)), fill=MUTED, anchor="mm")
            d.line([x1, y, x1, y + th], fill=C_GRID)
        # заголовки
        yh = y + th
        d.rectangle([x, yh, x + lw, yh + self.hh], fill=(230, 232, 235))
        for c in self.cols:
            x0, x1 = self.cx[c[0]]
            d.rectangle([x0, yh, x1, yh + self.hh], fill=C_HEAD, outline=C_GRID)
            ls = c[1].split("\n")
            for k, s in enumerate(ls):
                d.text(((x0 + x1) / 2, yh + self.hh / 2 + (k - (len(ls) - 1) / 2) * 24 * sc), s, font=CALB(int(21 * sc)), fill=INK, anchor="mm")
        # строки
        for i, row in enumerate(self.rows):
            y0 = yh + self.hh + i * self.rh
            d.rectangle([x, y0, x + lw, y0 + self.rh], fill=(230, 232, 235), outline=C_GRID)
            if self.first is not None:
                d.text((x + lw / 2, y0 + self.rh / 2), str(self.first + i), font=CAL(int(20 * sc)), fill=MUTED, anchor="mm")
            for c in self.cols:
                x0, x1 = self.cx[c[0]]
                fill = KIND[c[3]]
                if (c[0], i) in glow: fill = mix(fill, (255, 214, 170), glow[(c[0], i)])
                d.rectangle([x0, y0, x1, y0 + self.rh], fill=fill, outline=C_GRID)
                val = over.get((c[0], i), row.get(c[0], ""))
                if (c[0], i) in typed: val = typed[(c[0], i)]
                if val is None: val = ""
                col = INK
                if val in ("ok",): col = OK
                if isinstance(val, str) and any(w in val for w in ("ПЕРЕГРУЗ", "НЕ ЗАЩИЩ", "ЛИШНИЙ", "ДУБЛЬ")): col = BAD
                f = CAL(self.fs)
                s = str(val)
                while text_w(d, s, f) > (x1 - x0 - 12) and len(s) > 3: s = s[:-2] + "…"
                al = c[4] if len(c) > 4 else "c"
                if al == "l": d.text((x0 + 10 * sc, y0 + self.rh / 2), s, font=f, fill=col, anchor="lm")
                else: d.text(((x0 + x1) / 2, y0 + self.rh / 2), s, font=f, fill=col, anchor="mm")

    def ring(self, d, col, i, a, col2=None, color=ACC):
        if a <= 0: return
        x0, y0, x1, y1 = self.cell(col, i)
        if col2:
            x1 = self.cx[col2][1]
        pad = 3 + 4 * (1 - a)
        d.rounded_rectangle([x0 - pad, y0 - pad, x1 + pad, y1 + pad], 6, outline=mix(BG, color, a), width=4)

    def ring_rows(self, d, col, i0, i1, a, color=ACC):
        if a <= 0: return
        x0, y0, _, _ = self.cell(col, i0); _, _, x1, y1 = self.cell(col, i1)
        d.rounded_rectangle([x0 - 4, y0 - 4, x1 + 4, y1 + 4], 6, outline=mix(BG, color, a), width=4)


def typing(s, t, t0, dur):
    k = max(0.0, min(1.0, (t - t0) / dur))
    n = int(round(len(s) * k))
    return s[:n] + ("|" if 0 < k < 1 else "")


def fade(im, t, dur):
    a = min(1.0, t / 0.35, (dur - t) / 0.35)
    if a >= 1: return im
    return Image.blend(Image.new("RGB", (W, H), BG), im, max(0.0, a))


# ------------------------------------------------------------------ данные примера (ЖК «Остров», ЩР)
ROWS_R = [
    {"B": "25", "E": "", "G": "QFD6", "H": "2P С16А 30мА", "P": "XT6", "R": "R.1.05.1", "S": "Розеточная группа Пом. №05", "T": "0,42", "U": "1,91"},
    {"B": "26", "E": "", "G": "QFD7", "H": "2P С16А 30мА", "P": "XT7", "R": "R.1.05.2", "S": "Розеточная группа Пом. №05", "T": "0,06", "U": "0,28"},
    {"B": "27", "E": "", "G": "QFD8", "H": "2P С16А 30мА", "P": "XT8", "R": "R.1.05.3", "S": "Розеточная группа Пом. №05", "T": "0,36", "U": "1,64"},
    {"B": "28", "E": "", "G": "QFD9", "H": "2P С16А 30мА", "P": "XT9", "R": "R.1.05.4", "S": "Варочная поверхность", "T": "0,06", "U": "0,28"},
]


# ================================================================== сцены
def s_intro(t, dur):
    im = Image.new("RGB", (W, H), (26, 32, 40)); d = ImageDraw.Draw(im)
    a = ease(t / 0.8)
    d.text((W / 2, 400), "Однолинейка с умным домом", font=DJB(76), fill=mix((26, 32, 40), (255, 255, 255), a), anchor="mm")
    d.text((W / 2, 500), "от таблицы Excel до готовой схемы в AutoCAD", font=DJ(40), fill=mix((26, 32, 40), (200, 208, 218), a), anchor="mm")
    b = ease((t - 0.8) / 0.8)
    d.rounded_rectangle([W / 2 - 360, 600, W / 2 + 360, 606], 3, fill=mix((26, 32, 40), ACC, b))
    d.text((W / 2, 680), "Таблица «Однолинейка ЩР ЭОМ»  ·  программа SX_Schema.lsp", font=DJ(28), fill=mix((26, 32, 40), (150, 160, 172), b), anchor="mm")
    return im


def s_flow(t, dur):
    im, d = frame_base(1, "Как это устроено")
    boxes = [("Исходные данные", "потребители, мощность,\nкабель, длины"),
             ("Разбивка по щитам", "линии по местам\nв щите"),
             ("Лист «Однолинейка»", "автоматы, аппараты,\nмодули, каналы, клеммы"),
             ("AutoCAD: SXDRAW", "схема + нижняя\nтаблица из ваших блоков")]
    bw, bh, gap = 380, 230, 70
    x0 = (W - (4 * bw + 3 * gap)) / 2; y0 = 330
    for k, (h, s) in enumerate(boxes):
        a = ease((t - 0.3 - 0.9 * k) / 0.6)
        if a <= 0: continue
        x = x0 + k * (bw + gap)
        hl = k == 2
        d.rounded_rectangle([x, y0, x + bw, y0 + bh], 18, fill=mix(BG, (255, 255, 255) if not hl else (255, 246, 226), a),
                            outline=mix(BG, ACC if hl else (190, 196, 204), a), width=4 if hl else 2)
        d.text((x + bw / 2, y0 + 62), h, font=DJB(30), fill=mix(BG, INK, a), anchor="mm")
        for j, line in enumerate(s.split("\n")):
            d.text((x + bw / 2, y0 + 130 + 38 * j), line, font=DJ(26), fill=mix(BG, MUTED, a), anchor="mm")
        if k > 0:
            ax = x - gap + 12
            d.line([ax, y0 + bh / 2, x - 14, y0 + bh / 2], fill=mix(BG, INK, a), width=5)
            d.polygon([(x - 14, y0 + bh / 2 - 12), (x - 14, y0 + bh / 2 + 12), (x - 2, y0 + bh / 2)], fill=mix(BG, INK, a))
    caption(d, ["Вы заполняете «Исходные данные» как обычно. Лист «Однолинейка» сам подтягивает линии "
                "и считает автоматы. Вы дописываете только то, что знает проектировщик: модули, каналы, аппараты. "
                "Дальше одна команда в AutoCAD."], t, 4.0)
    return im


def s_colors(t, dur):
    im, d = frame_base(2, "Лист «Однолинейка»: что заполнять")
    cols = [("B", "№\nместа", 80, "src"), ("D", "Шина", 110, "in"), ("E", "Автомат", 130, "in"), ("F", "30 мА", 90, "in"),
            ("G", "QF", 100, "res"), ("H", "Автомат\nитог", 200, "res"), ("I", "Аппарат 1", 150, "in"),
            ("L", "Модуль", 130, "in"), ("N", "Канал", 90, "in"), ("P", "Клемма", 100, "res"),
            ("R", "Линия", 130, "src"), ("S", "Наименование", 330, "src", "l")]
    g = Grid(60, 150, cols, ROWS_R, first_row=27)
    rows = [dict(r) for r in ROWS_R]; rows[0]["E"] = "авто"
    g.rows = rows
    g.draw(d)
    # легенда
    lg = [(C_IN, "жёлтые — заполняете вы"), (C_RES, "зелёные — считаются сами"), (C_SRC, "серые — из исходных данных")]
    for k, (c, s) in enumerate(lg):
        a = ease((t - 1.0 - 1.4 * k) / 0.5)
        yy = int(g.y + g.th + g.hh + g.rh * 4 + 40 + 56 * k)
        d.rounded_rectangle([90, yy, 140, yy + 40], 6, fill=c, outline=mix(BG, INK, a), width=2)
        d.text((160, yy + 20), s, font=DJ(30), fill=mix(BG, INK, a), anchor="lm")
    for k, cs in enumerate([("D", "F"), ("G", "H")]):
        a = ease((t - 1.0 - 1.4 * k) / 0.5) * (1 - ease((t - 2.2 - 1.4 * k) / 0.5))
        x0 = g.cx[cs[0]][0]; x1 = g.cx[cs[1]][1]
        d.rounded_rectangle([x0 - 4, g.y + g.th, x1 + 4, g.y + g.th + g.hh + g.rh * 4 + 4], 8, outline=mix(BG, ACC, a), width=5)
    caption(d, ["Столбцы идут в том же порядке, что и на схеме сверху вниз: шина → автомат → аппараты → модуль и канал → "
                "клемма → кабель → линия. Служебные столбцы справа можно не открывать."], t, 5.2)
    return im


def s_breaker(t, dur):
    im, d = frame_base(3, "Автомат: «авто» или свой текст")
    cols = [("B", "№\nместа", 80, "src"), ("E", "Автомат\n«авто» / текст", 190, "in"), ("F", "30 мА", 90, "in"),
            ("G", "QF", 110, "res"), ("H", "Автомат итог", 230, "res"), ("P", "Клемма", 110, "res"),
            ("R", "Линия", 140, "src"), ("S", "Наименование", 360, "src", "l"), ("U", "Ток, А", 100, "src")]
    rows = [dict(r) for r in ROWS_R]
    rows[1]["G"] = "QFD6"; rows[1]["H"] = ""
    g = Grid(60, 150, cols, rows, first_row=27)
    over = {}; typed = {}
    # 1) вписываем «авто» в первую строку
    typed[("E", 0)] = typing("авто", t, 1.2, 0.6)
    if t < 1.9:
        for i in range(4): over[("H", i)] = ""; over[("G", i)] = ""
    else:
        for i in (2, 3): over[("H", i)] = ""; over[("G", i)] = ""; over[("E", i)] = ""
    # 2) третья строка — свой текст
    typed[("E", 2)] = typing("2P С20А 30мА", t, 7.2, 1.2)
    if t > 8.5: over[("G", 2)] = "QFD8"; over[("H", 2)] = "2P С20А 30мА"
    typed[("E", 3)] = typing("авто", t, 10.0, 0.6)
    if t > 10.7: over[("G", 3)] = "QFD9"; over[("H", 3)] = "2P С16А 30мА"
    glow = {("H", 0): ease((t - 2.0) / 0.4) * (1 - ease((t - 3.4) / 0.6)),
            ("G", 1): ease((t - 4.0) / 0.4) * (1 - ease((t - 6.2) / 0.6))}
    g.draw(d, over, typed, glow)
    g.ring(d, "G", 0, ease((t - 4.0) / 0.4) * (1 - ease((t - 6.2) / 0.4)), color=ACC2)
    g.ring(d, "G", 1, ease((t - 4.0) / 0.4) * (1 - ease((t - 6.2) / 0.4)), color=ACC2)
    keys = [(0, 900, 800), (1.0, *g.center("E", 0)), (3.0, *g.center("E", 0)), (4.2, *g.center("E", 1)), (6.6, *g.center("E", 1)),
            (7.0, *g.center("E", 2)), (9.6, *g.center("E", 2)), (9.9, *g.center("E", 3)), (dur, *g.center("E", 3))]
    x, y = path_pos(keys, t)
    clk = max(ease(1 - abs(t - 1.1) / 0.25), ease(1 - abs(t - 7.1) / 0.25), ease(1 - abs(t - 9.95) / 0.25))
    cursor(d, x, y, clk if clk > 0 else 0)
    if t < 6.8:
        cap = ["«авто» в первой строке группы — программа подберёт автомат и полюса. "
               "Строка ниже с пустым «Автомат» — та же группа: она сидит на том же QFD6."]
    else:
        cap = ["Можно вписать свой автомат текстом, например «2P С20А 30мА» — он возьмётся как есть и тоже будет проверен."]
    caption(d, cap, t, 0.6)
    return im


def s_nominal(t, dur):
    im, d = frame_base(4, "Как считается номинал")
    cols = [("R", "Линия", 140, "src"), ("U", "Ток, А", 100, "src"), ("Z", "Ток\nгруппы", 110, "res"),
            ("AA", "по току", 110, "res"), ("BB", "мин. по\nкоду", 110, "res"), ("AB", "макс. по\nкабелю", 120, "res"),
            ("H", "Автомат итог", 230, "res"), ("W", "Проверка", 240, "res")]
    rows = [{"R": "R.1.05.1", "U": "1,91", "Z": "1,91", "AA": "6", "BB": "16", "AB": "25", "H": "2P С16А 30мА", "W": "ok"},
            {"R": "L.1.06.1", "U": "0,00", "Z": "0,00", "AA": "6", "BB": "10", "AB": "16", "H": "1P С10А", "W": "ok"},
            {"R": "ТП.1.11.1", "U": "6,23", "Z": "6,23", "AA": "10", "BB": "—", "AB": "25", "H": "2P С10А 30мА", "W": "ok"},
            {"R": "ПН.1.17.1", "U": "27,3", "Z": "27,3", "AA": "32", "BB": "—", "AB": "25", "H": "2P С32А 30мА", "W": "КАБЕЛЬ НЕ ЗАЩИЩЁН"}]
    g = Grid(80, 150, cols, rows, first_row=None)
    g.draw(d)
    a1 = ease((t - 1.0) / 0.5) * (1 - ease((t - 5.5) / 0.4))
    g.ring(d, "AA", 0, a1, col2="BB"); g.ring(d, "AA", 1, a1, col2="BB")
    g.ring(d, "H", 0, a1, color=ACC2); g.ring(d, "H", 1, a1, color=ACC2)
    a2 = ease((t - 6.5) / 0.5)
    g.ring(d, "AB", 3, a2, col2="W", color=BAD)
    if t < 6.2:
        cap = ["Номинал = наибольший из трёх: по току группы, минимум по коду потребителя (розетки R — 16 А, свет L — 10 А) "
               "и ручной автомат из «Исходных данных»."]
    else:
        cap = ["Проверка: если ток больше номинала — «ПЕРЕГРУЗ». Если автомат больше, чем выдерживает сечение кабеля "
               "(1,5 мм² → 16 А, 2,5 мм² → 25 А) — «КАБЕЛЬ НЕ ЗАЩИЩЁН»: увеличьте сечение."]
    caption(d, cap, t, 0.6)
    return im


def s_ref(t, dur):
    im, d = frame_base(5, "Справочник_схемы: ваши правила")
    cols = [("F", "Код", 140, "in"), ("G", "УГО", 170, "in"), ("H", "30 мА", 110, "in"), ("I", "Мин.\nавтомат, А", 150, "in")]
    rows = [{"F": "R", "G": "SOCKET", "H": "да", "I": "16"}, {"F": "L", "G": "LAMP", "H": "нет", "I": "10"},
            {"F": "LED", "G": "LAMP", "H": "нет", "I": "10"}, {"F": "ТП", "G": "HEAT", "H": "да", "I": ""},
            {"F": "SH", "G": "MOTOR", "H": "нет", "I": ""}, {"F": "ЭК", "G": "VALVE", "H": "нет", "I": ""}]
    g = Grid(80, 150, cols, rows, first_row=7, maxw=900)
    typed = {("I", 3): typing("16", t, 5.2, 0.4)} if t > 5.2 else {}
    g.draw(d, typed=typed)
    g.ring(d, "I", 0, ease((t - 1.0) / 0.4) * (1 - ease((t - 4.5) / 0.4))); g.ring(d, "I", 1, ease((t - 1.0) / 0.4) * (1 - ease((t - 4.5) / 0.4)))
    g.ring(d, "H", 0, ease((t - 2.4) / 0.4) * (1 - ease((t - 4.5) / 0.4)), color=ACC2)
    # модули
    cols2 = [("A", "Модель", 190, "in"), ("B", "Префикс", 110, "in"), ("C", "Каналов", 110, "in")]
    rows2 = [{"A": "ZIOMB24V2", "B": "R", "C": "24"}, {"A": "ZIOMB8V3", "B": "R", "C": "8"}, {"A": "ZDINDX4", "B": "Di", "C": "4"},
             {"A": "ZDID64X2", "B": "DA", "C": "2"}]
    g2 = Grid(1100, 150, cols2, rows2, first_row=4, maxw=740)
    g2.draw(d)
    d.text((1100, g2.y + g2.th + g2.hh + g2.rh * 4 + 30), "Модули умного дома: модель → префикс имени (R.1, Di.1…)\nи сколько каналов", font=DJ(26), fill=MUTED)
    caption(d, ["Код потребителя (R, L, ТП…) задаёт значок УГО на схеме, нужен ли УЗО 30 мА и минимальный автомат. "
                "Хотите тёплый пол всегда на 16 А — впишите 16 в строку ТП. Новые модели модулей добавляйте в таблицу справа."], t, 0.6)
    return im


def s_devices(t, dur):
    im, d = frame_base(6, "Аппараты после автомата")
    cols = [("E", "Автомат", 120, "in"), ("H", "Автомат итог", 190, "res"), ("I", "Аппарат 1", 190, "in"), ("J", "Аппарат 2", 170, "in"),
            ("K", "Аппарат 3", 150, "in"), ("L", "Модуль", 150, "in"), ("R", "Линия", 170, "src")]
    rows = [{"E": "авто", "H": "1P С10А", "R": "В.1.17.2"}, {"E": "авто", "H": "1P С10А", "R": "L.1.03.3"},
            {"E": "авто", "H": "1P С16А", "R": "ТП.1.13.1"}]
    g = Grid(60, 150, cols, rows, first_row=88, maxw=1330)
    typed = {("I", 0): typing("T 12V 54W", t, 1.0, 1.0), ("I", 1): typing("БП", t, 3.2, 0.4),
             ("I", 2): typing("МК-5-1", t, 5.0, 0.7), ("J", 2): typing("КМ", t, 6.2, 0.4)}
    g.draw(d, typed=typed)
    # мини-схема цепочки
    x = 1650; y = 190
    items = [("QF", 1.0), ("T 12V 54W", 2.0), ("канал модуля", 2.6), ("клемма XT", 3.0)]
    for k, (s, tt) in enumerate(items):
        a = ease((t - tt) / 0.4)
        d.rounded_rectangle([x - 150, y + 110 * k, x + 150, y + 110 * k + 64], 12, fill=mix(BG, (255, 255, 255), a), outline=mix(BG, ACC if k == 1 else (180, 186, 194), a), width=3)
        d.text((x, y + 110 * k + 32), s, font=DJB(26), fill=mix(BG, (153, 27, 30) if k == 1 else INK, a), anchor="mm")
        if k: d.line([x, y + 110 * k - 46, x, y + 110 * k], fill=mix(BG, INK, a), width=4)
    caption(d, ["До трёх аппаратов подряд после автомата: трансформатор «T 12V 54W», «БП», «МК-5-1», «КМ». "
                "С трансформатором или БП автомат будет не меньше 10 А, а на схеме аппарат встанет сбоку от автомата."], t, 0.6)
    return im


def s_module(t, dur):
    im, d = frame_base(7, "Модуль умного дома и каналы")
    cols = [("G", "QF", 100, "res"), ("L", "Модуль\n(модель)", 190, "in"), ("M", "Модуль", 110, "res"), ("N", "Канал\nввод", 110, "in"),
            ("O", "Канал", 100, "res"), ("P", "Клемма", 110, "res"), ("R", "Линия", 190, "src"), ("S", "Наименование", 330, "src", "l")]
    rows = [{"G": "QF45", "R": "В.1.03.1", "S": "Вентилятор Пом. №03"}, {"G": "QF45", "R": "В.1.15.1", "S": "Вентилятор Пом. №15"},
            {"G": "QF45", "R": "", "S": ""}, {"G": "QF45", "R": "В.1.17.2", "S": "Вентилятор Пом. №17"},
            {"G": "QF45", "R": "К.1.02.1", "S": "Кондиционер Пом. №02"}, {"G": "QF45", "R": "", "S": ""}]
    g = Grid(80, 150, cols, rows, first_row=78)
    typed = {("L", 0): typing("ZIOMB24V2", t, 1.0, 1.0), ("N", 2): typing("р", t, 6.0, 0.3), ("L", 5): typing("-", t, 8.4, 0.3)}
    over = {}
    if t > 2.2:
        for i in range(5):
            over[("M", i)] = "R.1"
        chans = [1, 2, 3, 4, 5]
        for i in range(5):
            over[("O", i)] = chans[i] if (i != 2 or t > 6.4) else ""
        xt = {0: "XT49", 1: "XT50", 3: "XT51", 4: "XT52"}
        for i, v in xt.items(): over[("P", i)] = v
    g.draw(d, over, typed)
    g.ring_rows(d, "M", 0, 4, ease((t - 2.4) / 0.4) * (1 - ease((t - 5.4) / 0.4)))
    g.ring_rows(d, "O", 0, 4, ease((t - 2.4) / 0.4) * (1 - ease((t - 5.4) / 0.4)), color=ACC2)
    if t < 5.8:
        cap = ["В первой строке модуля впишите модель (ZIOMB24V2). Имя R.1 и номера каналов 1, 2, 3… проставятся сами, "
               "клеммы XT — тоже."]
    elif t < 8.2:
        cap = ["Пустая строка без линии — впишите «р»: это резервный канал, номер пропускать не нужно."]
    else:
        cap = ["Модуль закончился — поставьте «-» в «Модуль». Следующая модель в столбце начнёт новый модуль (R.2, Di.1…)."]
    caption(d, cap, t, 0.4)
    return im


def s_codes(t, dur):
    im, d = frame_base(8, "Столбец «Канал»: особые случаи")
    items = [("пусто", "номер канала по порядку (обычный случай)"),
             ("7", "свой номер канала, если нужен не по порядку"),
             ("+", "линия в тот же канал и ту же клемму, что строка выше"),
             ("=", "параллельно предыдущей линии (перемычка на клеммах)"),
             ("р", "резервный канал"),
             ("б/к", "без канала и без клеммы: вводы, PE, КУП, шина DALI")]
    for k, (a_, b_) in enumerate(items):
        a = ease((t - 0.4 - 0.9 * k) / 0.4)
        y = 170 + 105 * k
        d.rounded_rectangle([140, y, 330, y + 76], 12, fill=mix(BG, C_IN, a), outline=mix(BG, (200, 190, 150), a), width=2)
        d.text((235, y + 38), a_, font=MONO(40), fill=mix(BG, INK, a), anchor="mm")
        d.text((380, y + 38), b_, font=DJ(34), fill=mix(BG, INK, a), anchor="lm")
    caption(d, ["Ошибки канала (два одинаковых номера, номер больше, чем каналов у модели) подсветятся в столбце проверки «Канал»."], t, 6.2)
    return im


def s_bus(t, dur):
    im, d = frame_base(9, "Шина и режим «вручную»")
    cols = [("D", "Шина\n(пусто = основная)", 230, "in"), ("E", "Автомат", 120, "in"), ("H", "Автомат итог", 200, "res"),
            ("Y", "Режим", 150, "in"), ("R", "Линия", 170, "src"), ("S", "Наименование", 360, "src", "l")]
    rows = [{"D": "", "E": "авто", "H": "2P С16А 30мА", "R": "R.1.07.1", "S": "Розеточная группа Пом. №07"},
            {"D": "", "E": "авто", "H": "2P С16А 30мА", "R": "R.1.07.2", "S": "Розеточная группа Пом. №07"},
            {"D": "", "E": "авто", "H": "1P С10А", "R": "LED.1.08.1", "S": "Светодиодное освещение"},
            {"D": "", "E": "", "H": "", "R": "Ввод.2", "S": "Дополнительный ввод"}]
    g = Grid(120, 150, cols, rows, first_row=25)
    typed = {("D", 1): typing("ИБП", t, 1.0, 0.5), ("D", 2): typing("КМ-1", t, 2.2, 0.5), ("Y", 3): typing("вручную", t, 6.0, 0.8)}
    g.draw(d, typed=typed)
    if t < 5.6:
        cap = ["«Шина»: пусто — основная шина L1,L2,L3. Впишите имя (ИБП, КМ-1) — автомат сядет на вторую шину с этим именем, "
               "она протянется до конца листа."]
    else:
        cap = ["«вручную» — строку программа не рисует и не подбирает (вводы, особые узлы). Её вы дочерчиваете сами."]
    caption(d, cap, t, 0.6)
    return im


def s_checks(t, dur):
    im, d = frame_base(10, "Перед построением: проверки")
    cols = [("J", "Итоги по однолинейке", 520, "src", "l"), ("K", "", 120, "res")]
    rows = [{"J": "Автоматов всего", "K": "51"}, {"J": "  из них с 30 мА (QFD)", "K": "28"}, {"J": "Клемм XT", "K": "52"},
            {"J": "Модулей всего", "K": "0"}, {"J": "Ошибок по автоматам", "K": "0"}, {"J": "Ошибок по каналам", "K": "0"}]
    g = Grid(80, 150, cols, rows, first_row=7, maxw=900)
    g.draw(d)
    g.ring(d, "J", 4, ease((t - 1.2) / 0.4), col2="K", color=OK); g.ring(d, "J", 5, ease((t - 1.2) / 0.4), col2="K", color=OK)
    d.text((1060, 260), "Что значат красные надписи", font=DJB(30), fill=INK)
    for k, s in enumerate(["ПЕРЕГРУЗ — ток группы больше номинала", "КАБЕЛЬ НЕ ЗАЩИЩЁН — увеличьте сечение",
                           "ЛИШНИЙ КАНАЛ — у модели меньше каналов", "ДУБЛЬ КАНАЛА — один канал дважды"]):
        a = ease((t - 2.0 - 0.6 * k) / 0.4)
        d.text((1060, 330 + 60 * k), s, font=DJ(28), fill=mix(BG, BAD, a))
    caption(d, ["Ошибок должно быть 0. Красные надписи на листе «Однолинейка» показывают, в какой строке проблема."], t, 4.6)
    return im


def s_acad(t, dur):
    im, d = frame_base(11, "Построение в AutoCAD")
    # командная строка
    d.rounded_rectangle([80, 150, 900, 470], 14, fill=(33, 40, 48))
    lines = [("Команда: ", "APPLOAD", 0.6), ("", "[SX_Schema v3.5.3]", 1.6), ("Команда: ", "SXDRAW", 2.4)]
    for k, (p, s, tt) in enumerate(lines):
        if t > tt:
            d.text((110, 190 + 60 * k), p + typing(s, t, tt, 0.5), font=MONO(30), fill=(210, 220, 230) if k != 1 else (120, 200, 140))
    # окно SXDRAW
    a = ease((t - 3.2) / 0.5)
    if a > 0:
        x0, y0 = 980, 150
        d.rounded_rectangle([x0, y0, x0 + 860, y0 + 680], 12, fill=mix(BG, (248, 248, 248), a), outline=mix(BG, (150, 156, 164), a), width=2)
        d.rectangle([x0, y0, x0 + 860, y0 + 50], fill=mix(BG, (60, 70, 82), a))
        d.text((x0 + 20, y0 + 25), "SX_Schema — однолинейная схема из Excel", font=DJ(24), fill=mix(BG, (255, 255, 255), a), anchor="lm")
        fields = [("Файл Excel", "…\\Однолинейка_v3.3.xlsx", "Обзор…"), ("Лист", "Однолинейка", ""), ("Щит", "ЩР", ""),
                  ("Точка вставки", "2497,2; 1522,5", "Указать точку"), ("Библиотека", "…\\02_Однолинейка.dwg", "Обзор…"),
                  ("Мест / шаг", "17 мест на лист, шаг 20 мм", ""), ("", "☑ удалить прежнее построение", "")]
        for k, (lab, val, btn) in enumerate(fields):
            yy = y0 + 80 + 72 * k
            hl = ease(1 - abs(t - (4.4 + 1.1 * k)) / 0.7) if k < 5 else 0
            d.text((x0 + 30, yy + 22), lab, font=DJ(24), fill=mix(BG, INK, a), anchor="lm")
            d.rounded_rectangle([x0 + 290, yy, x0 + 640 if btn else x0 + 830, yy + 44], 6,
                                fill=mix(BG, (255, 255, 255), a), outline=mix((170, 176, 184), ACC, hl), width=2 + int(2 * hl))
            d.text((x0 + 302, yy + 22), val, font=DJ(22), fill=mix(BG, INK, a), anchor="lm")
            if btn:
                d.rounded_rectangle([x0 + 655, yy, x0 + 830, yy + 44], 6, fill=mix(BG, (228, 232, 238), a), outline=mix(BG, (160, 166, 174), a))
                d.text((x0 + 742, yy + 22), btn, font=DJ(20), fill=mix(BG, INK, a), anchor="mm")
        okh = ease(1 - abs(t - 10.6) / 0.6)
        d.rounded_rectangle([x0 + 560, y0 + 610, x0 + 690, y0 + 660], 8, fill=mix(mix(BG, (228, 232, 238), a), ACC, okh))
        d.text((x0 + 625, y0 + 635), "OK", font=DJB(24), fill=mix(BG, INK, a), anchor="mm")
        d.rounded_rectangle([x0 + 705, y0 + 610, x0 + 835, y0 + 660], 8, fill=mix(BG, (228, 232, 238), a))
        d.text((x0 + 770, y0 + 635), "Отмена", font=DJ(22), fill=mix(BG, INK, a), anchor="mm")
    steps = ["1. APPLOAD → SX_Schema.lsp (один раз за сеанс)", "2. SXDRAW", "3. Файл Excel и щит", "4. «Указать точку» — левый нижний угол таблицы",
             "5. Библиотека — ваш 02_Однолинейка.dwg", "6. OK — схема и нижняя таблица готовы"]
    for k, s in enumerate(steps):
        a2 = ease((t - 1.0 - 1.6 * k) / 0.4)
        d.text((90, 520 + 50 * k), s, font=DJ(28), fill=mix(BG, INK, a2))
    return im


def s_tips(t, dur):
    im, d = frame_base(12, "Если что-то пошло не так")
    tips = [("SXCLEAR", "удалить всё построенное (слои SX_*) и построить заново"),
            ("SX_log.txt", "журнал последнего запуска в папке с программой — что и где сломалось"),
            ("Блоки не найдены", "проверьте путь к 02_Однолинейка.dwg — иначе рисуются простые заменители"),
            ("Поменяли таблицу", "сохраните Excel и снова SXDRAW — схема перестроится целиком")]
    for k, (a_, b_) in enumerate(tips):
        a = ease((t - 0.4 - 1.0 * k) / 0.4)
        y = 190 + 130 * k
        d.rounded_rectangle([120, y, 520, y + 90], 14, fill=mix(BG, (255, 255, 255), a), outline=mix(BG, ACC2, a), width=3)
        d.text((320, y + 45), a_, font=DJB(30), fill=mix(BG, INK, a), anchor="mm")
        d.text((570, y + 45), b_, font=DJ(30), fill=mix(BG, INK, a), anchor="lm")
    return im


def s_outro(t, dur):
    im = Image.new("RGB", (W, H), (26, 32, 40)); d = ImageDraw.Draw(im)
    a = ease(t / 0.6)
    d.text((W / 2, 440), "Заполнили таблицу → SXDRAW → готово", font=DJB(60), fill=mix((26, 32, 40), (255, 255, 255), a), anchor="mm")
    d.text((W / 2, 540), "Подробности — README.md в папке SX_Однолинейка", font=DJ(34), fill=mix((26, 32, 40), (170, 180, 192), a), anchor="mm")
    return im


SCENES = [(s_intro, 5), (s_flow, 11), (s_colors, 11), (s_breaker, 13), (s_nominal, 12), (s_ref, 10), (s_devices, 10),
          (s_module, 12), (s_codes, 11), (s_bus, 11), (s_checks, 9), (s_acad, 13), (s_tips, 8), (s_outro, 5)]

if PREVIEW:
    for k, (fn, dur) in enumerate(SCENES):
        fade(fn(dur * 0.8, dur), dur * 0.8, dur).save(f"prev_{k:02d}.png")
    sys.exit()

p = subprocess.Popen(["ffmpeg", "-y", "-loglevel", "error", "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{W}x{H}", "-r", str(FPS),
                      "-i", "-", "-c:v", "libx264", "-preset", "medium", "-crf", "20", "-pix_fmt", "yuv420p", "-movflags", "+faststart", OUT],
                     stdin=subprocess.PIPE)
for fn, dur in SCENES:
    n = int(dur * FPS)
    for i in range(n):
        t = i / FPS
        p.stdin.write(fade(fn(t, dur), t, dur).tobytes())
p.stdin.close(); p.wait()
print("done", OUT, sum(d for _, d in SCENES), "s")
