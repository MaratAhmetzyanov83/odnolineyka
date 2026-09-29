"""Видеогайд «Однолинейка v3.11: рабочий процесс» — кадры PIL → ffmpeg (H.264), без звука, с подписями.
Использует графические помощники из make.py (первая видеоинструкция) и реальные значения объекта «ЖД» (video_data.json).
python3 make_guide.py out.mp4 [--preview]
"""
import json, os, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = open(os.path.join(HERE, "make.py"), encoding="utf-8").read()
exec(SRC.split("# ------------------------------------------------------------------ данные примера")[0])   # W, H, Grid, caption, cursor…

OUT = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") else "guide.mp4"
PREVIEW = "--preview" in sys.argv
D = json.load(open(os.path.join(HERE, "video_data.json"), encoding="utf-8"))
DARK = (26, 32, 40)
TAB_BLUE, TAB_GREEN, TAB_ORANGE, TAB_GREY = (68, 114, 196), (112, 173, 71), (237, 125, 49), (166, 166, 166)


def n(v, k=2):
    if v in (None, ""): return ""
    if isinstance(v, (int, float)):
        s = f"{v:.{k}f}".rstrip("0").rstrip(".") if isinstance(v, float) else str(v)
        return s.replace(".", ",")
    return str(v)


def base(step, title):
    im, d = frame_base(step, title)
    d.rectangle([W - 560, 0, W, 92], fill=DARK)
    d.text((W - 40, 46), "Однолинейка v3.11 · рабочий процесс", font=DJ(20), fill=(150, 160, 172), anchor="rm")
    return im, d


def box(d, x, y, w, h, title, sub, a, color=(255, 255, 255), edge=(190, 196, 204), tsize=28):
    d.rounded_rectangle([x, y, x + w, y + h], 16, fill=mix(BG, color, a), outline=mix(BG, edge, a), width=3)
    d.text((x + w / 2, y + 34), title, font=DJB(tsize), fill=mix(BG, INK, a), anchor="mm")
    for j, s in enumerate(sub.split("\n") if sub else []):
        d.text((x + w / 2, y + 74 + 32 * j), s, font=DJ(22), fill=mix(BG, MUTED, a), anchor="mm")


def arrow(d, x0, y0, x1, y1, a, color=INK, w=4):
    import math
    d.line([x0, y0, x1, y1], fill=mix(BG, color, a), width=w)
    ang = math.atan2(y1 - y0, x1 - x0); L = 16
    p1 = (x1 - L * math.cos(ang - 0.45), y1 - L * math.sin(ang - 0.45)); p2 = (x1 - L * math.cos(ang + 0.45), y1 - L * math.sin(ang + 0.45))
    d.polygon([(x1, y1), p1, p2], fill=mix(BG, color, a))


# ================================================================== сцены
def s_intro(t, dur):
    im = Image.new("RGB", (W, H), DARK); d = ImageDraw.Draw(im)
    a = ease(t / 0.8)
    d.text((W / 2, 380), "Однолинейка v3.11", font=DJB(84), fill=mix(DARK, (255, 255, 255), a), anchor="mm")
    d.text((W / 2, 480), "рабочий процесс: от строки в таблице до ПЗ и схемы", font=DJ(40), fill=mix(DARK, (200, 208, 218), a), anchor="mm")
    b = ease((t - 0.8) / 0.8)
    d.rounded_rectangle([W / 2 - 360, 580, W / 2 + 360, 586], 3, fill=mix(DARK, ACC, b))
    d.text((W / 2, 660), "пример — объект «Резиденция, здание жилого дома»: ГРЩ-ЖД и 14 щитов", font=DJ(28), fill=mix(DARK, (150, 160, 172), b), anchor="mm")
    return im


def s_map(t, dur):
    im, d = base(1, "Объект = папка, щит = файл")
    a0 = ease((t - 0.3) / 0.6)
    box(d, W / 2 - 230, 150, 460, 120, "Однолинейка ГРЩ-ЖД.xlsx", "ГРЩ: фидеры к щитам", a0, (255, 246, 226), ACC)
    kids = ["ЩР-001", "ЩР-01", "ЩР-1", "ЩР-2", "ЩОВ-01", "ЩОВ-001", "ЩОВ-002", "ЩОВ-2", "ЩС-Б", "ЩС-ВК.001", "ЩС-ВК.002", "ЩС-ВК.003", "ЩУОВ", "ЩСПЗ-ЖД"]
    bw, bh, gap = 200, 70, 16
    x0 = (W - (7 * bw + 6 * gap)) / 2
    for k, nm in enumerate(kids):
        a = ease((t - 1.0 - 0.12 * k) / 0.4)
        row, col = divmod(k, 7)
        x = x0 + col * (bw + gap); y = 420 + row * (bh + 26)
        d.rounded_rectangle([x, y, x + bw, y + bh], 10, fill=mix(BG, (255, 255, 255), a), outline=mix(BG, (180, 186, 194), a), width=2)
        d.text((x + bw / 2, y + bh / 2), nm, font=DJB(24), fill=mix(BG, INK, a), anchor="mm")
        if row == 0: d.line([x + bw / 2, y, x + bw / 2, 330], fill=mix(BG, (170, 176, 184), a), width=2)
    a1 = ease((t - 3.2) / 0.5)
    d.line([x0 + bw / 2, 330, x0 + 6 * (bw + gap) + bw / 2, 330], fill=mix(BG, (170, 176, 184), a1), width=2)
    d.line([W / 2, 270, W / 2, 330], fill=mix(BG, (170, 176, 184), a1), width=2)
    a2 = ease((t - 4.0) / 0.5)
    arrow(d, W / 2 + 300, 300, W / 2 + 300, 160, a2, OK, 5)
    d.text((W / 2 + 325, 230), "вверх: Руст, Кс-группы, cos φ, Iр", font=DJ(26), fill=mix(BG, OK, a2), anchor="lm")
    arrow(d, W / 2 - 300, 160, W / 2 - 300, 300, a2, ACC2, 5)
    d.text((W / 2 - 325, 230), "вниз: ΔU до шин, Zпетли, Iкз, In фидера", font=DJ(26), fill=mix(BG, ACC2, a2), anchor="rm")
    caption(d, ["Один файл — один щит, все файлы объекта в одной папке. Щиты связаны с ГРЩ обычной «Вставить связь»: нагрузка "
                "поднимается вверх, потери напряжения и ток КЗ спускаются вниз. Файлы не переименовывайте."], t, 5.0)
    return im


def s_sheets(t, dur):
    im, d = base(2, "Листы книги: цвет ярлыка = роль")
    items = [(TAB_BLUE, "Справочная, Каталог", "коды потребителей, этажи, кабели, таблица Кс, оборудование"),
             (TAB_GREEN, "Исходные данные", "линии: мощность, фазы, длины, кабель, группа Кс, категория"),
             (TAB_GREEN, "Разбивка_по_щитам", "какая линия в каком щите и в каком порядке"),
             (TAB_GREEN, "Однолинейка", "автоматы, УЗО, модули, каналы, клеммы → AutoCAD"),
             (TAB_GREEN, "Питание_щита · Нижестоящие_щиты", "связи с другими однолинейками, ΔU и Iкз до шин"),
             (TAB_ORANGE, "Нагрузка_щитов", "фазы, Кс, Рр, Iр щита (читает «В Акад»)"),
             (TAB_ORANGE, "Расчёт_нагрузок", "таблица и основные показатели для ПЗ стадии П"),
             (TAB_ORANGE, "ТКЗ, Кабельный журнал, Спецификация, Бирки", "результаты для чертежей и закупки"),
             (TAB_GREY, "Спецификация_авто, В Акад, Диапазоны", "служебные — не удалять, не переименовывать")]
    for k, (c, nm, s) in enumerate(items):
        a = ease((t - 0.3 - 0.45 * k) / 0.4)
        y = 130 + 86 * k
        d.rounded_rectangle([90, y, 130, y + 58], 6, fill=mix(BG, c, a))
        d.text((160, y + 29), nm, font=DJB(30), fill=mix(BG, INK, a), anchor="lm")
        d.text((1010, y + 29), s, font=DJ(26), fill=mix(BG, MUTED, a), anchor="lm")
    caption(d, ["Заполняете зелёные листы, оранжевые считаются сами. Внутри листа: жёлтые ячейки — ввод, остальное — формулы."], t, 4.8)
    return im


def s_new(t, dur):
    im, d = base(3, "Новый щит")
    a = ease((t - 0.3) / 0.5)
    box(d, 180, 260, 560, 150, "Однолинейка_шаблон.xlsx", "чистый шаблон v3.11", a, (255, 255, 255))
    b = ease((t - 1.6) / 0.5)
    arrow(d, 760, 335, 1000, 335, b, ACC, 6)
    d.text((880, 300), "копия", font=DJ(26), fill=mix(BG, ACC, b), anchor="mm")
    c = ease((t - 2.4) / 0.5)
    name = typing("Однолинейка_ЩР-3.xlsx", t, 2.6, 1.6)
    box(d, 1020, 260, 640, 150, name or " ", "папка объекта, рядом с ГРЩ", c, (255, 246, 226), ACC)
    rules = ["имя: «Однолинейка_<щит>.xlsx», без номера версии", "не вставлять строки и столбцы, не вырезать (Ctrl+X)",
             "числа — с запятой: 0,7 (текст «0.7» подсветится красным)"]
    for k, s in enumerate(rules):
        e = ease((t - 4.3 - 0.6 * k) / 0.4)
        d.text((200, 500 + 60 * k), "•  " + s, font=DJ(30), fill=mix(BG, INK, e), anchor="lm")
    caption(d, ["Каждый щит начинается с копии шаблона. Структуру книги не меняем — всё новое уже есть в свободных столбцах."], t, 6.4)
    return im


def isx_rows():
    rows = []
    for r in D["isx"][1:5]:
        rows.append({"D": r["D"], "E": n(r["E"]), "F": n(r["F"]), "H": n(r["H"]), "I": n(r["I"]), "M": n(r["M"]),
                     "N": r["N"], "O": n(r["O"]), "Q": n(r["Q"]), "R": n(r["R"]), "T": n(r["T"]), "U": n(r["U"]),
                     "V": n(r["V"]), "AB": n(r["AB"]), "AZ": r["AZ"] or "", "BA": n(r["BA"]), "BB": n(r["BB"]),
                     "BD": r["BD"] or "", "BF": n(r["BF"]), "BG": n(r["BG"]), "BH": n(r["BH"]), "BK": r["BK"] or ""})
    return rows


def s_isx1(t, dur):
    im, d = base(4, "«Исходные данные»: строка = линия")
    cols = [("D", "Номер\nлинии", 150, "in"), ("E", "P, кВт", 90, "in"), ("F", "U, В", 80, "in"), ("H", "Фаз", 70, "in"),
            ("I", "Длина\nпо плану", 110, "in"), ("M", "Длина\nитог", 100, "res"), ("N", "Марка", 80, "in"), ("O", "Жил", 70, "in"),
            ("Q", "Сечение\nручн.", 100, "in"), ("R", "Автомат\nручн.", 100, "in"),
            ("T", "Ток, А", 90, "res"), ("U", "Автомат", 100, "res"), ("V", "Сечение", 100, "res"), ("AB", "ΔU, %", 90, "res")]
    rows = isx_rows()
    g = Grid(40, 150, cols, rows, first_row=11, maxw=W - 80)
    typed = {}
    if t > 6.5:
        rows.append({"D": "", "E": "", "F": "", "H": "", "I": "", "N": "", "O": ""})
        g.rows = rows
        typed = {("D", 4): typing("R.2.021А.1", t, 6.8, 1.0), ("E", 4): typing("0,36", t, 8.0, 0.4),
                 ("F", 4): typing("230", t, 8.6, 0.3), ("H", 4): typing("1", t, 9.0, 0.2), ("I", 4): typing("28", t, 9.3, 0.3),
                 ("N", 4): typing("мс", t, 9.7, 0.3), ("O", 4): typing("3", t, 10.1, 0.2)}
    over = {}
    if t > 10.6:
        over = {("M", 4): "43", ("T", 4): "1,64", ("U", 4): "16", ("V", 4): "2,5", ("AB", 4): "0,4"}
    g.draw(d, over, typed)
    g.ring_rows(d, "T", 0, 3, ease((t - 1.5) / 0.4) * (1 - ease((t - 5.8) / 0.4)))
    g.ring_rows(d, "AB", 0, 3, ease((t - 1.5) / 0.4) * (1 - ease((t - 5.8) / 0.4)), color=ACC2)
    if t < 6.3:
        cap = ["Жёлтые столбцы — ваши: номер линии (код.этаж.помещение.№), мощность, напряжение, фазы, длина по плану, марка кабеля, жилы. "
               "Ток — с cos φ при 230/400 В, автомат, сечение по ПУЭ табл. 1.3.6 и потери по ГОСТ Р 50571.5.52 считаются сами."]
    else:
        cap = ["Новая линия: заполнили жёлтое — итоговая длина (с опусками и этажами), ток, автомат, сечение и ΔU появились сразу. "
               "Сечение/автомат можно задать вручную в «Сечение ручн.» / «Автомат ручн.»."]
    caption(d, cap, t, 0.6)
    return im


def s_isx2(t, dur):
    im, d = base(5, "«Исходные данные»: группа Кс, категория, проверки")
    cols = [("D", "Номер\nлинии", 150, "src"), ("AZ", "Группа Кс", 120, "in"), ("BA", "Кс\n(Сил)", 80, "in"), ("BB", "Секция", 90, "in"),
            ("BD", "Категория для ПЗ", 330, "in", "l"), ("BH", "К прокл.", 90, "in"),
            ("BF", "Норма\nΔU, %", 90, "res"), ("BG", "ΔU от\nВРУ, %", 100, "res"), ("BK", "Отключение\n≤ 0,4 с", 160, "res")]
    rows = isx_rows()
    g = Grid(60, 150, cols, rows, first_row=11, maxw=W - 120)
    g.draw(d)
    g.ring_rows(d, "AZ", 0, 3, ease((t - 0.8) / 0.4) * (1 - ease((t - 5.4) / 0.4)))
    g.ring_rows(d, "BG", 0, 3, ease((t - 6.0) / 0.4), color=ACC2)
    grp = [("Быт", "бытовая: Кс по таблице СП 256 по сумме Руст быт щита"), ("Сил", "силовая: свой Кс в «Кс (Сил)», пусто = 1"),
           ("Пост", "постоянная: Кс = 1 (обогрев, слаботочка, СПЗ)"), ("Нет", "не считается: вводы, резерв")]
    for k, (gname, s) in enumerate(grp):
        a = ease((t - 1.2 - 0.5 * k) / 0.4) * (1 - ease((t - 5.6) / 0.4))
        y = 560 + 52 * k
        d.text((120, y), gname, font=DJB(30), fill=mix(BG, ACC, a), anchor="lm")
        d.text((230, y), s, font=DJ(28), fill=mix(BG, INK, a), anchor="lm")
    if t < 5.8:
        cap = ["Группа Кс: пусто — возьмётся по коду потребителя. Секция 1/2 — для ГРЩ с двумя вводами. Категория — строка таблицы ПЗ."]
    else:
        cap = ["ΔU от ВРУ = потери линии + потери до шин щита (из «Питание_щита»). Норма СП 256 п. 8.23: свет ≤ 3 %, прочие ≤ 4 %. "
               "Больше нормы — ячейка красная. «Отключение ≤ 0,4 с» появится, когда известен ток КЗ на шинах."]
    caption(d, cap, t, 0.6)
    return im


def s_rz(t, dur):
    im, d = base(6, "«Разбивка_по_щитам»: порядок = места на схеме")
    cols = [("C", "№\nместа", 90, "res"), ("F", "Щит", 150, "in"), ("H", "Номер линии", 190, "in"), ("J", "Наименование", 560, "res", "l"),
            ("L", "P, кВт", 100, "res"), ("M", "Фаза", 100, "res")]
    rows = [{"C": n(r["C"]), "F": r["F"], "H": r["H"], "J": r["J"], "L": n(r["L"]) if r["L"] != "кВт" else "", "M": r["M"]} for r in D["rz"][:6]]
    g = Grid(120, 150, cols, rows, first_row=None, maxw=W - 240)
    g.draw(d)
    g.ring_rows(d, "F", 0, 5, ease((t - 0.8) / 0.4) * (1 - ease((t - 4.4) / 0.4)))
    g.ring_rows(d, "C", 0, 5, ease((t - 4.8) / 0.4), color=ACC2)
    caption(d, ["Сюда вписываются только «Щит» и «Номер линии» — в том порядке, в каком автоматы стоят в щите. "
                "Номер места, наименование, мощность и фаза подтянутся сами; фазы 1ф линий раскладываются L1-L2-L3 по кругу."], t, 0.6)
    return im


def s_ol(t, dur):
    im, d = base(7, "«Однолинейка»: автомат «авто» и статус")
    cols = [("B", "№\nместа", 80, "src"), ("E", "Автомат", 110, "in"), ("F", "30 мА", 80, "in"), ("G", "QF", 100, "res"),
            ("H", "Автомат итог", 190, "res"), ("P", "Клемма", 100, "res"), ("Q", "Кабель", 380, "res", "l"), ("R", "Линия", 150, "src"),
            ("W", "Проверка", 120, "res"), ("AE", "Статус", 260, "res", "l"), ("Y", "Режим", 110, "in")]
    rows = []
    for r in D["ol"][:6]:
        rows.append({k: n(r.get(k)) for k in ("B", "E", "F", "G", "H", "P", "Q", "R", "W", "AE", "Y")})
    g = Grid(30, 150, cols, rows, first_row=None, maxw=W - 60)
    g.draw(d)
    g.ring(d, "Y", 0, ease((t - 0.8) / 0.4) * (1 - ease((t - 3.8) / 0.4)))
    g.ring_rows(d, "E", 1, 5, ease((t - 4.2) / 0.4) * (1 - ease((t - 7.6) / 0.4)))
    g.ring_rows(d, "H", 1, 5, ease((t - 4.6) / 0.4) * (1 - ease((t - 7.6) / 0.4)), color=ACC2)
    g.ring_rows(d, "AE", 0, 5, ease((t - 8.0) / 0.4), color=BAD)
    if t < 4.0:
        cap = ["Ввод помечен «вручную» — программа его не подбирает и не рисует."]
    elif t < 7.8:
        cap = ["«авто» — автомат подбирается сам: наибольший из «по току», «мин. по коду» (розетки 16 А, свет 10 А) и ручного. "
               "Розеткам сразу добавляется 30 мА. Модули, каналы и аппараты — как в первой видеоинструкции."]
    else:
        cap = ["Столбец «Статус»: всё, что надо поправить, — нет фазы/мощности, перегруз, кабель не защищён, потери выше нормы, Iкз мало. "
               "Сводка ошибок — в шапке листа."]
    caption(d, cap, t, 0.6)
    return im


def s_load(t, dur):
    im, d = base(8, "«Нагрузка_щитов»: фазы и Кс по СП 256")
    cols = [("A", "Группа", 90, "res"), ("B", "Кс", 70, "res"), ("D", "Линия", 150, "src"), ("E", "Фаза\nручн.", 90, "in"),
            ("G", "Фаза", 80, "res"), ("H", "P, кВт", 90, "src")]
    rows = [{"A": r["A"], "B": n(r["B"]), "D": r["D"], "E": n(r["E"]), "G": r["G"], "H": n(r["H"]) if r["H"] != "кВт" else ""} for r in D["ld"]["lines"][:7]]
    g = Grid(40, 150, cols, rows, first_row=3, maxw=760)
    g.draw(d)
    b = D["ld"]["blk"]
    labels = ["Руст, кВт", "Руст быт, кВт", "Кс быт (табл. СП 256)", "Рр сил = Σ P·Кс", "Рр пост (Кс = 1)", "Рр, кВт", "Кс итоговый", "Iр по фазе, А"]
    x0 = 900; y0 = 170
    for k, (lab, v) in enumerate(zip(labels, b)):
        a = ease((t - 1.5 - 0.35 * k) / 0.4)
        y = y0 + 58 * k
        hl = k in (5, 7)
        d.rounded_rectangle([x0, y, x0 + 900, y + 50], 8, fill=mix(BG, (255, 246, 226) if hl else (255, 255, 255), a), outline=mix(BG, (200, 205, 212), a), width=2)
        d.text((x0 + 20, y + 25), lab, font=DJ(26), fill=mix(BG, INK, a), anchor="lm")
        d.text((x0 + 880, y + 25), n(v, 3), font=DJB(28), fill=mix(BG, INK, a), anchor="rm")
    ph = D["ld"]["ph"]
    a = ease((t - 5.0) / 0.4)
    d.text((x0, y0 + 58 * 8 + 20), f"Фазы, кВт:  L1 {n(ph[0])}   L2 {n(ph[1])}   L3 {n(ph[2])}", font=DJB(28), fill=mix(BG, BAD, a), anchor="lm")
    g.ring_rows(d, "E", 0, 6, ease((t - 6.5) / 0.4), color=ACC)
    if t < 6.3:
        cap = [f"Щит ЩР-2: Руст быт {n(b[1])} кВт → Кс 0,8 по таблице; постоянная {n(b[4])} кВт с Кс 1. Рр = {n(b[5])} кВт. "
               f"Iр считается по самой нагруженной фазе — {n(b[7], 1)} А."]
    else:
        cap = ["Перекос фаз (L3 почти втрое больше L2) — переставьте фазы в «Фаза ручн.» (1, 2, 3 или 1,2,3), Iр сразу уменьшится. "
               "Этот блок попадает в «В Акад» и в таблицу ПЗ."]
    caption(d, cap, t, 0.6)
    return im


def s_supply(t, dur):
    im, d = base(9, "«Питание_щита»: откуда щит питается")
    ps = D["ps"]
    cols = [("B", "Щит", 140, "src"), ("C", "ΔU от ВРУ\nдо шин, %", 150, "in"), ("D", "Zпетли\nна шинах, Ом", 150, "in"),
            ("E", "Iкз max\nна шинах, кА", 150, "in"), ("F", "In аппарата\nфидера, А", 150, "in"), ("G", "Файл вышестоящего", 330, "in", "l")]
    imp = ps["6"]
    g = Grid(60, 140, cols, [{"B": imp[0], "C": n(imp[1]), "D": "", "E": "", "F": n(imp[4]), "G": imp[5]}], first_row=6, maxw=1200)
    g.draw(d)
    g.ring(d, "C", 0, ease((t - 0.8) / 0.4) * (1 - ease((t - 4.5) / 0.4)), col2="F")
    res = [("29", "Рр щита, кВт"), ("30", "Iр щита, А"), ("33", "ΔU от ВРУ до шин → в линии, %"), ("34", "Zпетли на шинах → в линии, Ом"),
           ("35", "Iкз min на шинах, А"), ("38", "Аппарат и кабель"), ("39", "Отключение КЗ ≤ 5 с"), ("40", "Термостойкость кабеля"),
           ("41", "Селективность")]
    for k, (row, lab) in enumerate(res):
        a = ease((t - 2.0 - 0.3 * k) / 0.4)
        v = ps.get(row, [None, None])[1]
        y = 400 + 50 * k
        d.text((80, y), lab, font=DJ(26), fill=mix(BG, INK, a), anchor="lm")
        s = n(v, 2) if v not in (None, "") else "— (нужны данные ТП/ввода)" if row in ("34", "35", "39", "40") else "—"
        col = BAD if isinstance(v, str) and v != "ok" else (OK if v == "ok" else INK)
        d.text((720, y), s[:70], font=DJB(26), fill=mix(BG, col, a), anchor="lm")
    if t < 6.5:
        cap = ["Раздел 1 заполняется связью из ГРЩ (ΔU до шин 2,94 %, фидер 16 А). Если вышестоящего файла нет — раздел 2: "
               "откуда питается, Iкз на шинах питающего или мощность ТП, кабель, длина, аппарат."]
    else:
        cap = ["Результат уходит во все линии щита: ΔU до шин и Zпетли. Здесь же проверки фидера: Iр ≤ In ≤ Iдоп, отключение, "
               "термостойкость и селективность. У ЩР-2 фидер 16 А и отходящие 16 А — нужна проверка селективности."]
    caption(d, cap, t, 0.6)
    return im


def s_link(t, dur):
    im, d = base(10, "Подключить щит к ГРЩ: две «Вставить связь»")
    cols = [("B", "Линия", 130, "in"), ("D", "Руст", 90, "in"), ("E", "Руст\nбыт", 90, "in"), ("G", "Рр\nсил", 90, "in"),
            ("H", "Руст\nпост", 90, "in"), ("J", "Iр, А", 90, "in"), ("K", "ΔU до\nшин, %", 100, "res"), ("N", "In\nфидера", 90, "res"),
            ("O", "Проверка фидера", 330, "res", "l")]
    rows = []
    for r in D["ns"][4:10]:
        rows.append({"B": r[1], "D": n(r[3]), "E": n(r[4]), "G": n(r[6]), "H": n(r[7]), "J": n(r[9], 1), "K": n(r[10]), "N": n(r[13]), "O": r[14]})
    g = Grid(40, 150, cols, rows, first_row=9, maxw=1320)
    over = {}
    if t < 4.0:
        for c in "DEGHJ": over[(c, 3)] = ""
    g.draw(d, over)
    # мини-окно нижестоящего
    a = ease((t - 0.6) / 0.4)
    x0, y0 = 1400, 170
    d.rounded_rectangle([x0, y0, x0 + 480, y0 + 250], 14, fill=mix(BG, (255, 255, 255), a), outline=mix(BG, (190, 196, 204), a), width=2)
    d.text((x0 + 20, y0 + 30), "Однолинейка_ЩР-2.xlsx", font=DJB(24), fill=mix(BG, INK, a), anchor="lm")
    d.text((x0 + 20, y0 + 70), "«Питание_щита», раздел 4", font=DJ(22), fill=mix(BG, MUTED, a), anchor="lm")
    d.rounded_rectangle([x0 + 20, y0 + 100, x0 + 460, y0 + 150], 8, fill=mix(BG, C_RES, a), outline=mix(BG, ACC, ease((t - 1.0) / 0.4)), width=3)
    d.text((x0 + 240, y0 + 125), "C:I  9,98 · 5,28 · 0 · 0 · 4,7 · 0,95 · 22,5", font=DJ(20), fill=mix(BG, INK, a), anchor="mm")
    d.text((x0 + 20, y0 + 195), "Ctrl+C →", font=DJB(24), fill=mix(BG, ACC, ease((t - 1.4) / 0.4)), anchor="lm")
    d.text((x0 + 20, y0 + 228), "Спец. вставка → «Вставить связь»", font=DJ(20), fill=mix(BG, ACC, ease((t - 2.4) / 0.4)), anchor="lm")
    g.ring(d, "D", 3, ease((t - 2.6) / 0.4) * (1 - ease((t - 6.0) / 0.4)), col2="J")
    g.ring(d, "K", 3, ease((t - 6.5) / 0.4), col2="N", color=ACC2)
    g.ring(d, "O", 3, ease((t - 9.5) / 0.4), color=BAD)
    if t < 6.3:
        cap = ["1) В ГРЩ, лист «Нижестоящие_щиты», впишите линию-фидер. 2) В файле щита скопируйте C:I раздела 4 «Питание_щита» → "
               "здесь выделите D:J → «Вставить связь». Линия становится группой «Щит», Кс пересчитывается по сумме."]
    elif t < 9.3:
        cap = ["3) Обратно: скопируйте K:N этой строки → в файле щита «Питание_щита», раздел 1, C:F → «Вставить связь». "
               "Щит получит ΔU до шин, Zпетли, Iкз и номинал фидера."]
    else:
        cap = ["Проверка фидера: у ЩР-2 расчётный ток 22,5 А больше автомата 16 А в ГРЩ. При открытии Excel спросит про связи — «Обновить»."]
    caption(d, cap, t, 0.6)
    return im


def s_checks(t, dur):
    im, d = base(11, "Где смотреть ошибки и что делать")
    items = [("Однолинейка · Статус", "ПЕРЕГРУЗ", "увеличить автомат или разделить группу"),
             ("", "КАБЕЛЬ НЕ ЗАЩИЩЁН", "увеличить сечение или уменьшить автомат"),
             ("", "потери 4,9 % > 4 %", "увеличить сечение / укоротить трассу / щит ближе"),
             ("", "Iкз мало — t > 0,4 с", "увеличить сечение, УЗО 30 мА или характеристика B"),
             ("Исходные данные", "красный «Кс (Сил)»", "Кс записан текстом — ввести числом с запятой"),
             ("Питание_щита", "In < 1,6·In отходящего", "проверить селективность по картам производителя"),
             ("Нижестоящие_щиты", "Iр > In", "фидер мал для расчётного тока щита"),
             ("Расчёт_нагрузок", "расхождение", "щит не в блоках 1–3 или не пересчитан (F9)")]
    for k, (where, what, todo) in enumerate(items):
        a = ease((t - 0.3 - 0.45 * k) / 0.4)
        y = 140 + 80 * k
        d.text((80, y + 25), where, font=DJB(26), fill=mix(BG, MUTED, a), anchor="lm")
        d.rounded_rectangle([470, y, 930, y + 52], 8, fill=mix(BG, (255, 199, 206), a))
        d.text((490, y + 26), what, font=DJB(26), fill=mix(BG, BAD, a), anchor="lm")
        d.text((960, y + 26), todo, font=DJ(26), fill=mix(BG, INK, a), anchor="lm")
    caption(d, ["Порядок работы: заполнили → посмотрели «Статус» и красные ячейки → поправили данные → только потом схема и ПЗ."], t, 4.8)
    return im


def s_pz(t, dur):
    im, d = base(12, "«Расчёт_нагрузок»: таблица для ПЗ")
    cols = [("A", "№", 50, "res"), ("B", "Наименование эл. приёмника", 470, "res", "l"), ("C", "Руст,\nкВт", 110, "res"),
            ("D", "Кс", 80, "res"), ("E", "cos φ", 80, "res"), ("F", "tg φ", 80, "res"), ("G", "Рр,\nкВт", 110, "res"),
            ("H", "Qр,\nквар", 110, "res"), ("I", "Sр,\nкВА", 110, "res"), ("J", "Iр, А", 110, "res")]
    rows = []
    for r in D["pz"][:8]:
        rows.append({"A": n(r[0]), "B": r[1], "C": n(r[2]), "D": n(r[3]), "E": n(r[4]), "F": n(r[5]), "G": n(r[6]), "H": n(r[7]), "I": n(r[8]), "J": n(r[9], 1)})
    tot = D["pz"][9]
    rows.append({"A": "", "B": "Расчётная нагрузка щита (Ко = 0,9)", "C": n(tot[2]), "D": n(tot[3]), "E": n(tot[4]), "F": n(tot[5]),
                 "G": n(tot[6]), "H": n(tot[7]), "I": n(tot[8]), "J": n(tot[9], 1)})
    g = Grid(60, 150, cols, rows, first_row=None, maxw=W - 120, smax=1.2)
    g.draw(d)
    g.ring(d, "A", 8, ease((t - 4.0) / 0.4), col2="J")
    if t < 6.5:
        cap = ["Вверху выбираете щит и режим (весь щит / секция 1 / 2). Строки — категории нагрузки, как в ПЗ; итог — с Ко = 0,9 "
               "для силовой. Для ГРЩ-ЖД: Рр 275,8 кВт, Sр 290,3 кВА, Iр 419 А."]
    else:
        cap = ["Ниже — «Основные показатели»: напряжение, система заземления, категория надёжности, питание, ΔU, Iкз, КРМ, учёт. "
               "В ПЗ: выделить A7:J64 → Копировать → Вставить как значения."]
    caption(d, cap, t, 0.6)
    return im


def s_acad(t, dur):
    im, d = base(13, "Схема в AutoCAD")
    steps = [("APPLOAD", "загрузить SX_Schema.lsp (один раз)"), ("SXDRAW", "окно: файл, лист «Однолинейка», щит, точка, библиотека блоков"),
             ("SXUPDATE", "перестроить после правок в Excel (без окна, файл можно не сохранять)"), ("SXCLEAR", "удалить построенное (слои SX_*)")]
    for k, (cmd, s) in enumerate(steps):
        a = ease((t - 0.3 - 0.8 * k) / 0.4)
        y = 200 + 120 * k
        d.rounded_rectangle([140, y, 460, y + 80], 12, fill=mix(BG, DARK, a))
        d.text((300, y + 40), cmd, font=MONO(36), fill=mix(BG, (255, 255, 255), a), anchor="mm")
        d.text((500, y + 40), s, font=DJ(30), fill=mix(BG, INK, a), anchor="lm")
    caption(d, ["Схему строит программа по листу «Однолинейка» из ваших блоков. Нижняя таблица схемы — тоже из таблицы."], t, 3.8)
    return im


def s_rules(t, dur):
    im, d = base(14, "Коротко: правила")
    rules = ["один щит — один файл, все файлы объекта в одной папке, не переименовывать",
             "не вставлять строки/столбцы, не вырезать; новое вписывать в жёлтые ячейки",
             "числа — с запятой; Кс, длины, мощности — числом",
             "при открытии ГРЩ — «Обновить связи», чтобы подтянулись щиты",
             "сначала «Статус» без красного, потом схема и таблица ПЗ",
             "ТП/ввод заполнить в «Питание_щита» ГРЩ — тогда посчитаются Iкз во всех щитах"]
    for k, s in enumerate(rules):
        a = ease((t - 0.3 - 0.6 * k) / 0.4)
        d.text((140, 190 + 95 * k), f"{k + 1}.  {s}", font=DJ(32), fill=mix(BG, INK, a), anchor="lm")
    return im


def s_outro(t, dur):
    im = Image.new("RGB", (W, H), DARK); d = ImageDraw.Draw(im)
    a = ease(t / 0.6)
    d.text((W / 2, 470), "Готово: таблица → проверки → ПЗ → схема", font=DJB(56), fill=mix(DARK, (255, 255, 255), a), anchor="mm")
    d.text((W / 2, 560), "шаблон и пример — в папке SX_Однолинейка; история — репозиторий odnolineyka", font=DJ(28), fill=mix(DARK, (150, 160, 172), a), anchor="mm")
    return im


SCENES = [(s_intro, 5), (s_map, 12), (s_sheets, 11), (s_new, 10), (s_isx1, 15), (s_isx2, 13), (s_rz, 10), (s_ol, 14),
          (s_load, 13), (s_supply, 14), (s_link, 15), (s_checks, 11), (s_pz, 13), (s_acad, 8), (s_rules, 9), (s_outro, 4)]

if PREVIEW:
    for k, (fn, dur) in enumerate(SCENES):
        fade(fn(dur * 0.8, dur), dur * 0.8, dur).save(os.path.join(HERE, f"gprev_{k:02d}.png"))
    sys.exit()

p = subprocess.Popen(["ffmpeg", "-y", "-loglevel", "error", "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{W}x{H}", "-r", str(FPS),
                      "-i", "-", "-c:v", "libx264", "-preset", "medium", "-crf", "21", "-pix_fmt", "yuv420p", "-movflags", "+faststart", OUT],
                     stdin=subprocess.PIPE)
for fn, dur in SCENES:
    for i in range(int(dur * FPS)):
        tt = i / FPS
        p.stdin.write(fade(fn(tt, dur), tt, dur).tobytes())
p.stdin.close(); p.wait()
print("done", OUT, sum(dd for _, dd in SCENES), "s")
