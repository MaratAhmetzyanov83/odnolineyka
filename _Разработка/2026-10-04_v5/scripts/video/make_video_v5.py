# -*- coding: utf-8 -*-
"""make_video_v5.py — видеогайд по книге v5 (пример «все случаи»): кадры из реальных листов, озвучка (голос Irina), субтитры, главы.

    py make_video_v5.py            -> 4_Инструкции/Видеогайд_v5.mp4 (+ .srt)
    py make_video_v5.py --frames   -> только кадры f##.png во временной папке (проверить рамки, без озвучки)

Книга копируется во временную папку; состояния (выдача, правки, упорядочивание, наклейки, выгрузка) создаются её же макросами.
Цвет шрифта условного форматирования (серое / синее) в PDF не печатается — перед снимком он задаётся напрямую (только в копии).
"""
import os
import shutil
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".."))
sys.path.insert(0, os.path.join(ROOT, "_Разработка", "2026-09-29_speedup-automation", "scripts"))
sys.path.insert(0, HERE)
from model_v312 import ExcelSession  # noqa: E402
import vidlib as V  # noqa: E402

SRC = os.path.join(ROOT, "_Разработка", "2026-10-04_v5", "out", "Однолинейка_v5_пример_все_случаи.xlsm")
OUT = os.path.join(ROOT, "4_Инструкции", "Видеогайд_v5.mp4")
TMP = os.path.join(os.environ.get("TEMP", "."), "sx_video_v5")
GREY, BLUE = 8421504, 12611584

STEPS = []          # (глава, подпись, текст озвучки, ключ снимка, рамки, кадрирование, надпись)
SHOTS = {}          # снимки: ключ → vidlib.Shot


def step(title, cap, text, shot=None, marks=(), crop=None, note=None):
    """адреса рамок и кадрирования пересчитываются в пиксели сразу (пока Excel открыт); кортеж — уже пиксели снимка."""
    if shot:
        sh = SHOTS[shot]
        marks = [a if isinstance(a, tuple) else sh.rect(a) for a in marks]
        crop = (crop if isinstance(crop, tuple) else sh.rect(crop)) if crop else None
    STEPS.append(dict(title=title, cap=cap, text=text, shot=shot, marks=list(marks), crop=crop, note=note))


def main():
    os.makedirs(TMP, exist_ok=True)
    work = os.path.join(TMP, "demo.xlsm")
    shutil.copy(SRC, work)
    exp = work[:-5] + "_AutoCAD.xlsx"
    if os.path.exists(exp):
        os.remove(exp)
    shots = SHOTS
    t0 = time.time()
    with ExcelSession() as xl:
        orig_af = xl.AutoCorrect.AutoFillFormulasInLists
        xl.AutomationSecurity = 1
        try:
            wb = xl.Workbooks.Open(work, 0, False)
            run = lambda m, *a: xl.Run("'" + wb.Name + "'!" + m, *a)
            wsl = wb.Worksheets("Линии")
            lo = wsl.ListObjects("Линии")
            heads = [c.Name for c in lo.ListColumns]
            colnum = {h: lo.ListColumns(h).Range.Column for h in heads}
            L = lambda h: V_letter(colnum[h])
            vis0 = {c: wsl.Columns(c).Hidden for c in range(1, len(heads) + 1)}
            wsl.Columns(colnum["Аппараты"]).ColumnWidth = 24

            def row_of(line, board=None):
                for r in range(lo.DataBodyRange.Row, lo.DataBodyRange.Row + lo.ListRows.Count):
                    if wsl.Cells(r, colnum["Номер линии"]).Value == line and (board is None or wsl.Cells(r, 1).Value == board):
                        return r
                raise KeyError(line)

            def rows_of(board):
                return [r for r in range(lo.DataBodyRange.Row, lo.DataBodyRange.Row + lo.ListRows.Count) if wsl.Cells(r, 1).Value == board]

            def view(cols, rows):
                """показать только заданные столбцы «Линий» (по заголовкам) и строки."""
                want = {colnum[h] for h in cols} | {1}
                for c in range(1, len(heads) + 1):
                    hide = c not in want
                    if wsl.Columns(c).Hidden != hide:
                        wsl.Columns(c).Hidden = hide
                r0, r1 = lo.DataBodyRange.Row, lo.DataBodyRange.Row + lo.ListRows.Count - 1
                rs = set(rows)
                for r in range(r0, r1 + 1):
                    h_ = r not in rs
                    if wsl.Rows(r).Hidden != h_:
                        wsl.Rows(r).Hidden = h_
                greyblue(cols, rows)
                return "A%d:%s%d" % (lo.HeaderRowRange.Row, V_letter(max(want)), max(rows))

            def greyblue(cols, rows):
                """серый курсив — формула, синий жирный — своё значение (в столбцах ввода с выбором таблицы)."""
                for h in cols:
                    c = colnum[h]
                    if wsl.Cells(lo.HeaderRowRange.Row, c).Font.Italic is not True:
                        continue                                  # курсив в шапке — столбец «серое / синее»
                    for r in rows:
                        cell = wsl.Cells(r, c)
                        if cell.HasFormula:
                            cell.Font.Color, cell.Font.Italic, cell.Font.Bold = GREY, True, False
                        elif cell.Value not in (None, ""):
                            cell.Font.Color, cell.Font.Italic, cell.Font.Bold = BLUE, False, True

            def snap(key, ws, area, **kw):
                shots[key] = V.shot(xl, ws, area, TMP, key, **kw)
                return shots[key]

            def find_row(ws, text, col_from=1, col_to=6, r_to=300):
                for r in range(1, r_to):
                    for c in range(col_from, col_to + 1):
                        if str(ws.Cells(r, c).Value or "").strip().startswith(text):
                            return r
                raise KeyError(text)

            # ---------- 1. титул и листы книги
            step("Однолинейка v5", "Видеогайд: что изменилось и как с ней работать",
                 "Это видеогайд по новой книге однолинейки, версия пять. За несколько минут посмотрим главное: как заполнять, "
                 "что книга считает сама и какие документы получаются.", note="Однолинейка v5 — видеогайд")
            wg = wb.Worksheets("Гайд")
            r1 = find_row(wg, "2. ЛИСТЫ КНИГИ")
            r2 = find_row(wg, "3. ПОРЯДОК РАБОТЫ")
            snap("sheets", wg, "A%d:C%d" % (r1, r2 - 1), gridlines=False)
            step("1. Листы книги", "Ввод → результаты → справка; расчётные листы — в конце",
                 "Одна книга — один объект. Сначала листы для ввода: Объект, Щиты, Линии и Коды. Потом результаты: Проверка, Щит, журнал, "
                 "спецификация, бирки, изменения и наклейки. В конце — справка и расчётные листы, которые нужны не всегда.", "sheets")
            # ---------- 2. «Объект»
            wo = wb.Worksheets("Объект")
            wo.Columns("D").ColumnWidth = 34
            snap("obj", wo, "B1:M16", gridlines=False)
            step("2. Лист «Объект»", "Название, методика, наименования; модули по умолчанию; настройки LED 24 В",
                 "Начинаем с листа «Объект»: название, методика расчёта и как собирать наименования линий. Справа — таблица модулей по умолчанию, "
                 "внизу — две настройки для линий двадцать четыре вольта: запас мощности блока питания и допустимые потери.",
                 "obj", marks=["B3:C3", "B8:C8", "B10:C10", "I2:K7", "B13:C14"])
            # ---------- 3. «Щиты»
            wsb = wb.Worksheets("Щиты")
            lb = wsb.ListObjects("Щиты")
            bh = [c.Name for c in lb.ListColumns]
            bc = {h: lb.ListColumns(h).Range.Column for h in bh}
            for c in range(1, len(bh) + 1):
                wsb.Columns(c).Hidden = bh[c - 1] not in ("Щит", "Вводов", "АВР", "ИБП от", "Мест под вводной узел", "Питается от",
                                                         "Замечания щита", "Руст, кВт", "Рр, кВт", "Iр, А", "ΔU до шин, %")
            rb0, rb1 = lb.HeaderRowRange.Row, lb.HeaderRowRange.Row + lb.ListRows.Count
            last = max(bc[h] for h in ("ΔU до шин, %", "Iр, А", "Замечания щита"))
            snap("boards", wsb, "A%d:%s%d" % (rb0, V_letter(last), rb1))
            step("3. Лист «Щиты»", "Строка на щит: вводы, АВР, ИБП, место под узел; «Питается от» — серым сам",
                 "Лист «Щиты» — строка на щит. Указываете число вводов, АВР, от какого ввода ИБП и сколько мест занимает вводной узел. "
                 "«Питается от» книга находит сама по линии-фидеру — это серое значение. Справа результаты щита: мощность, ток и потери до шин.",
                 "boards", marks=["%s%d:%s%d" % (V_letter(bc["Вводов"]), rb0 + 1, V_letter(bc["Мест под вводной узел"]), rb1),
                                  "%s%d:%s%d" % (V_letter(bc["Питается от"]), rb0 + 1, V_letter(bc["Питается от"]), rb1)])
            # ---------- 4–8. «Линии»: серое и синее, строка, замечания, «гр», УЗО
            R = rows_of("ЩР-1")
            main_cols = ["Номер линии", "Наименование", "Р, кВт", "Группа", "Автомат", "Аппараты", "Сечение, мм²", "Фаза", "Клемма", "Замечания"]
            area = view(main_cols, R)
            snap("lines", wsl, area)
            r_r101, r_r103 = row_of("R.1.01.1"), row_of("R.1.03.1")
            r_r102, r_r1022 = row_of("R.1.02.1"), row_of("R.1.02.2")
            r_uzo, r_uz3 = row_of("УЗО.1"), row_of("L.1.04.1")
            r_r104 = row_of("R.1.04.1")
            blue_cells = [a for a in ("%s%d" % (L("Р, кВт"), r_r104), "%s%d" % (L("Наименование"), r_r103), "%s%d" % (L("Фаза"), row_of("УВ.1.01.1")))]
            step("4. Серое и синее", "Серый курсив — выбрала таблица; синее — ваше значение (таблица проверяет)",
                 "Главное правило — серое и синее. Серый курсив — значение выбрала таблица: по коду, по расчёту или по чертежу. "
                 "Синее — вписали вы: таблица его не меняет, а проверяет. Если стереть синее, вернётся серое.",
                 "lines", marks=["%s%d:%s%d" % (L("Наименование"), r_r101, L("Автомат"), r_r101)] + blue_cells)
            step("5. Строка после DLN", "Подбор: наименование, мощность, автомат, сечение, фаза, клемма",
                 "После DLN в строке уже есть номер линии, длина и число потребителей. Остальное таблица подбирает сама: наименование по коду "
                 "и номеру помещения, мощность, автомат по току и коду, сечение по току, потерям и току короткого замыкания, фазу — "
                 "автобалансировкой, и номер клеммы.",
                 "lines", marks=["%s%d" % (L(h), r_r101) for h in ("Наименование", "Р, кВт", "Автомат", "Сечение, мм²", "Фаза", "Клемма")])
            step("6. Замечания", "«Замечания» — что не так в строке; сводка по объекту — лист «Проверка»",
                 "Столбец «Замечания» говорит, что не так в строке: перегруз, большие потери, медленное отключение, ток утечки. "
                 "Пустой столбец — строка в норме. Сводка по всему объекту — на листе «Проверка».",
                 "lines", marks=["%s%d" % (L("Замечания"), r_r103)])
            step("7. Группы «гр»", "«гр» — линия на автомате строки выше; автомат — по току группы, фаза общая",
                 "Несколько линий на одном автомате — «гр» в столбце «Группа»: строка идёт на автомат строки выше. "
                 "Автомат подбирается по току всей группы, а фаза у группы общая.",
                 "lines", marks=["%s%d:%s%d" % (L("Группа"), r_r102, L("Автомат"), r_r1022)])
            step("8. УЗО на группу", "Строка «УЗО.1» + линии «узо» со своими автоматами; проверки номинала и тока утечки",
                 "УЗО на группу автоматов: строка «УЗО.1» с УЗО в столбце «Аппараты», ниже — линии со своими автоматами и меткой «узо». "
                 "Автомат строки УЗО подбирается по току блока, у линий — обычные автоматы, а таблица проверяет номинал УЗО и ток утечки блока.",
                 "lines", marks=["%s%d:%s%d" % (L("Номер линии"), r_uzo, L("Номер линии"), r_uz3),
                                 "%s%d:%s%d" % (L("Группа"), r_uzo + 1, L("Группа"), r_uz3), "%s%d" % (L("Аппараты"), r_uzo)])
            # ---------- 9. кнопки над таблицей
            for c in range(1, len(heads) + 1):
                if wsl.Columns(c).Hidden != vis0[c]:
                    wsl.Columns(c).Hidden = vis0[c]
            for r in R[:6]:
                wsl.Rows(r).Hidden = False
            printable(wsl)
            sh_top = snap("top", wsl, "A1:%s%d" % (V_letter(colnum["Лист / место"]), R[0] + 3), gridlines=False)
            fx0 = shape_x(wsl, "Починить формулы")[0] * sh_top.kx
            fx1 = shape_x(wsl, "Зафиксировать выдачу")[1] * sh_top.kx
            y1 = sh_top.rect("A1")[3]
            step("9. Кнопки щитов", "Кнопка щита — видны только его строки; слева карточка щита (мощность, ток, замечания)",
                 "Над таблицей — кнопки щитов: нажмите кнопку, и видны только строки этого щита. Слева карточка щита: откуда питается, "
                 "мощность, ток, потери до шин и число замечаний.", "top", crop=(0, 12, fx0 - 18, y1 - 12))
            step("9. Кнопки макросов", "«Починить формулы», «Упорядочить щит», «Закрепить фазы», «Зафиксировать выдачу»",
                 "Рядом кнопки макросов: «Починить формулы» — после вставки строк, «Упорядочить щит», «Закрепить фазы» и «Зафиксировать выдачу». "
                 "Макросы при открытии книги нужно включить.", "top", crop=(fx0 + 8, 12, fx1 - 8, y1 - 12))
            # ---------- 10. модули (у линий ЩС — Relay и шторы)
            RS = rows_of("ЩС")
            k = 0
            for r in RS:
                ln = wsl.Cells(r, colnum["Номер линии"]).Value
                if ln and not str(ln).startswith("Ввод") and k < 5:
                    wsl.Cells(r, colnum["Управление"]).Value = "шторы" if k in (3, 4) else "Relay"
                    k += 1
            xl.Calculate()
            mod_cols = ["Номер линии", "Наименование", "Управление", "Модуль", "Канал (итог)"]
            area = view(mod_cols, RS)
            snap("mods", wsl, area)
            step("10. Модули", "Модуль по управлению — модель из «Объекта»; новый модуль, когда каналы кончились; шторы — 2 выхода",
                 "Модули подставляются сами. По управлению — реле, шторы, диммер, DALI или клапаны — таблица берёт модель из таблицы на листе "
                 "«Объект». Подряд идущие линии с одним управлением идут на каналы одного модуля, а когда каналы кончились, начинается новый. "
                 "Шторы занимают два выхода.", "mods",
                 marks=["%s%d:%s%d" % (L("Управление"), RS[1], L("Управление"), RS[-1]), "%s%d" % (L("Модуль"), RS[1]),
                        "%s%d:%s%d" % (L("Канал (итог)"), RS[1], L("Канал (итог)"), RS[-1])])
            # ---------- 11. LED 24 В
            RB = rows_of("ЩБП")
            led_cols = ["Номер линии", "Группа", "Автомат", "Аппараты", "Напряжение", "Iр, А", "Сечение, мм²", "ΔU от ВРУ, %", "БП: нужно, Вт", "Замечания"]
            area = view(led_cols, RB)
            snap("led", wsl, area)
            r_b1 = row_of("LD.1.01.1")
            step("11. LED 24 В", "Группа под автоматом БП; постоянный ток, потери ≤ 5 %; мощность БП × 1,3",
                 "Светодиодные линии двадцать четыре вольта — это группа под автоматом блока питания: в «Аппаратах» модуль защиты и «БП 24V 480W», "
                 "ниже линии с «гр». Таблица считает постоянный ток и потери от блока питания, не больше пяти процентов, и подбирает сечение. "
                 "Мощность блока питания проверяется с запасом 1,3 — здесь её не хватает, и это видно в замечаниях.", "led",
                 marks=["%s%d" % (L("Аппараты"), r_b1), "%s%d:%s%d" % (L("Сечение, мм²"), RB[1], L("Сечение, мм²"), RB[-1]),
                        "%s%d:%s%d" % (L("ΔU от ВРУ, %"), RB[1], L("ΔU от ВРУ, %"), RB[-1]),
                        "%s%d" % (L("БП: нужно, Вт"), r_b1), "%s%d" % (L("Замечания"), r_b1)])
            # ---------- 12–13. «Щит» и выгрузка
            wsh = wb.Worksheets("Щит")
            wsh.Range("C3").Value = "ВРУ"
            xl.Calculate()
            n_tab = wsh.Range("B7").SpillingToRange.Rows.Count if wsh.Range("B7").HasSpill else 1
            snap("board", wsh, "B1:P%d" % (6 + n_tab), gridlines=False)
            step("12. Лист «Щит»", "Щит в ячейке C3: таблица под схемой и итоги по каждому вводу",
                 "Лист «Щит» — то, что идёт на схему. Выберите щит в ячейке C3: слева таблица под схемой — строка на каждое место, "
                 "справа итоги по каждому вводу с категорией.", "board", marks=["C3", "L6:N15"])
            run("SX_ExportAutoCAD")
            wsh.Range("C3").Value = "ЩР-1"
            xl.Calculate()
            ex = xl.Workbooks.Open(exp, 0, True)
            wex = ex.Worksheets("ЩР-1 схема")
            nex = wex.UsedRange.Rows.Count
            snap("export", wex, "A1:I%d" % min(nex, 24), gridlines=True)
            names_ = ", ".join(w.Name for w in ex.Worksheets)[:200]
            step("13. Выгрузка для AutoCAD", "Лёгкий файл рядом с книгой: на щит — листы «схема», «итоги», «журнал», «спец»",
                 "Кнопка «Выгрузить для AutoCAD» на листе «Щит» создаёт рядом с книгой лёгкий файл. На каждый щит — схема, итоги, журнал "
                 "и спецификация, только значения. Имена листов постоянные, поэтому Data Link в AutoCAD после новой выгрузки достаточно обновить.",
                 "export")
            ex.Close(False)
            # ---------- 14–15. журнал
            wj = wb.Worksheets("Кабельный_журнал")
            wj.Range("E1").Value = "ЩР-1"
            xl.Calculate()
            nj = wj.Range("A5").SpillingToRange.Rows.Count if wj.Range("A5").HasSpill else 1
            snap("journal", wj, "A1:AN%d" % (4 + nj), gridlines=False)
            step("14. Кабельный журнал", "По щиту (E1) или по объекту; гофра отдельными числами — диаметр и метры",
                 "Кабельный журнал — по выбранному щиту или по всему объекту. Гофра теперь отдельными числами: диаметр и метры.", "journal",
                 crop="A1:T%d" % (4 + nj), marks=["E1", "F5:G%d" % (4 + nj), "O2:T%d" % (4 + nj)])
            step("15. Сводки журнала", "Кабели: запас 20 %, вверх до 10 м; гофра — до бухты; лоток; сводка ПТО",
                 "Справа сводки: кабели с запасом двадцать процентов и округлением вверх до десяти метров, гофра по диаметрам с округлением "
                 "до целой бухты, лоток и сводка ПТО с артикулами.", "journal", crop="V2:AN12", marks=["Y5:Y9", "AE5:AE6", "AL5:AL6"])
            # ---------- 16–17. спецификация
            sp = wb.Worksheets("Спецификация")
            sp.Range("D1").Value = "ЩР-1"
            xl.Calculate()
            ns = sp.Range("B3").SpillingToRange.Rows.Count if sp.Range("B3").HasSpill else 1
            for lo_ in sp.ListObjects:
                try:
                    lo_.ListColumns(next(c.Name for c in lo_.ListColumns if c.Name.startswith("Наименование"))).Range.ColumnWidth = 30
                except Exception:
                    pass
            snap("spec", sp, "B1:AK%d" % max(2 + ns, 16), gridlines=False)
            nocat = ["J%d" % r for r in range(3, 3 + ns) if "Каталог" in str(sp.Cells(r, 10).Value or "")]
            step("16. Спецификация", "По щиту (D1): аппараты из «Каталога»; чего нет — «впишите артикул»; кабели, гофра, крепёж",
                 "Спецификация тоже по щиту. Аппараты книга берёт из каталога по схеме, а то, чего в каталоге нет, показывает с пометкой "
                 "«впишите артикул». Ниже — кабели, гофра и крепёж.", "spec", crop="B1:J%d" % (2 + ns), marks=["D1"] + ([nocat[0] + ":" + nocat[-1]] if nocat else []))
            step("17. Корпус и комплект", "Корпус и комплект щита — вручную, один раз: строка на позицию, «Щит» — для какого щита",
                 "Корпус и комплект щита впишите один раз в таблицу справа от спецификации: строка на позицию и щит, к которому она относится.",
                 "spec", crop="M1:U5", marks=["M2:U3"])
            step("17. Нормы расхода и бухты", "Крепёж = база (метры гофры или кабеля) × норма на 1 м, вверх до кратности; гофра — до целой бухты",
                 "Рядом — нормы расхода крепежа: база, норма на метр и кратность, и таблица гофры с длиной бухты. По умолчанию — как в вашей "
                 "спецификации, их можно поправить под объект.", "spec", crop="W1:AK15", marks=["AB2:AD8", "AK2:AK14"])
            # ---------- 18. упорядочить щит (было / стало)
            R = rows_of("ЩР-1")
            ord_cols = ["Номер линии", "Наименование", "Группа", "Управление", "Автомат"]
            area = view(ord_cols, R)
            snap("ord1", wsl, area)
            step("18. Упорядочить щит", "Кнопка «Упорядочить щит»: так было…",
                 "Кнопка «Упорядочить щит» раскладывает строки щита по порядку. Так было…", "ord1")
            run("SX_SortBoard", "ЩР-1")
            xl.Calculate()
            R = rows_of("ЩР-1")
            area = view(ord_cols, R)
            snap("ord2", wsl, area)
            step("18. Упорядочить щит", "…а так стало: узел на месте, дальше по управлению, коду и помещению; группы и блоки УЗО целы",
                 "…а так стало: вводной узел на месте, дальше по управлению, коду и помещению. Группы и блоки УЗО не разрываются.", "ord2")
            # ---------- 19–20. выдача и изменения
            run("SX_Snapshot", "1")
            r_ = row_of("R.1.01.1")
            wsl.Cells(r_, colnum["Р, кВт"]).Value = 2.4
            r2_ = row_of("L.1.03.1")
            wsl.Cells(r2_, colnum["Фаза"]).Value = "L3"
            xl.Calculate()
            wc = wb.Worksheets("Изменения")
            nc = wc.Range("B5").SpillingToRange.Rows.Count if wc.Range("B5").HasSpill else 1
            snap("changes", wc, "B1:L%d" % (4 + min(nc, 14)), gridlines=False)
            step("19. Выдача и изменения", "«Зафиксировать выдачу» → правки → лист «Изменения»: было → стало и листы схемы",
                 "Перед выдачей нажмите «Зафиксировать выдачу»: книга запомнит состояние и предложит закрепить фазы. После правок лист "
                 "«Изменения» покажет, что добавлено, удалено и что было и стало, а справа — номера листов схемы для таблицы изменений.",
                 "changes", marks=["B4:G4", "J4:L6"])
            R = rows_of("ЩР-1")
            area = view(main_cols, R)
            snap("changed", wsl, area)
            step("20. Подсветка изменений", "В «Линиях» изменённые после выдачи ячейки — оранжевые, новые линии — зелёные",
                 "В самой таблице изменённые после выдачи ячейки подсвечены оранжевым, новые линии — зелёным.", "changed",
                 marks=["%s%d" % (L("Р, кВт"), row_of("R.1.01.1"))])
            # ---------- 21. наклейки
            run("SX_Labels", "ЩР-1")
            wlb = wb.Worksheets("Наклейки")
            vals = wlb.Range("A4:O80").Value
            lr = 3 + max([i + 1 for i, rw in enumerate(vals) if any(v not in (None, "") for v in rw)] or [4])
            snap("labels", wlb, "A1:O%d" % lr, gridlines=False)
            step("21. Наклейки", "Лента на A4: модуль DIN = 18 мм на бумаге; длинная рейка — несколькими полосами",
                 "Наклейки: кнопка на листе «Щит» строит ленту на А4. Ячейка — один модуль DIN, восемнадцать миллиметров на бумаге: ширина "
                 "подгоняется под ваш принтер. Под каждым аппаратом обозначение, номер линии и наименование. Длинная рейка печатается "
                 "несколькими полосами, над каждой — номер рейки. Печатайте в масштабе сто процентов.", "labels", crop="A4:O%d" % lr)
            step("Где подробнее", "Лист «Гайд» в книге, «Методика» — нормы и формулы. Замечания — пишите",
                 "Подробности — лист «Гайд» в книге и «Методика» с нормами. Если что-то неудобно или непонятно — пишите, поправим.",
                 note="Подробно — лист «Гайд» и «Методика»")
            wb.Close(False)
        finally:
            xl.AutoCorrect.AutoFillFormulasInLists = orig_af
    print("снимки: %d, %.0f с" % (len(shots), time.time() - t0))
    build_video(shots, "--frames" in sys.argv)


def printable(ws):
    """кнопки (элементы управления формы) по умолчанию не печатаются — для снимка включить печать."""
    for shp in ws.Shapes:
        try:
            if shp.Type == 8:
                shp.ControlFormat.PrintObject = True
        except Exception:
            pass


def shape_x(ws, text):
    """(левый, правый) край кнопки или фигуры с этой подписью, пт от левого края листа."""
    for shp in ws.Shapes:
        t = ""
        for get in (lambda: shp.TextFrame2.TextRange.Text, lambda: shp.OLEFormat.Object.Caption, lambda: shp.TextFrame.Characters().Text):
            try:
                t = get()
                break
            except Exception:
                pass
        if str(t).strip() == text:
            return shp.Left, shp.Left + shp.Width
    raise KeyError(text)


def V_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def build_video(shots, frames_only=False):
    clips, srt, meta, t = [], [], [";FFMETADATA1"], 0.0
    chapter = None
    for i, s in enumerate(STEPS, 1):
        png = os.path.join(TMP, "f%02d.png" % i)
        wav = os.path.join(TMP, "a%02d.wav" % i)
        mp4 = os.path.join(TMP, "c%02d.mp4" % i)
        V.frame(png, s["title"], s["cap"], shots.get(s["shot"]), s["marks"], note=s["note"], crop=s["crop"])
        if frames_only:
            continue
        dur = V.tts(V.say_text(s["text"]), wav, rate=2)
        dd = V.clip(png, wav, dur, mp4)
        clips.append(mp4)
        srt.append("%d\n%s --> %s\n%s\n" % (i, V.srt_time(t), V.srt_time(t + dur), s["text"]))
        ch = s["title"].split(". ", 1)[-1]
        if ch != chapter:
            meta.append("[CHAPTER]\nTIMEBASE=1/1000\nSTART=%d\nEND=%d\ntitle=%s" % (int(t * 1000), int((t + dd) * 1000), s["title"]))
            chapter = ch
        else:
            meta[-1] = meta[-1].rsplit("END=", 1)[0] + "END=%d\ntitle=%s" % (int((t + dd) * 1000), meta[-1].rsplit("title=", 1)[1])
        t += dd
        print("  шаг %2d: %5.1f с  %s" % (i, dd, s["title"]))
    if frames_only:
        print("кадры: %s" % TMP)
        return
    lst = os.path.join(TMP, "list.txt")
    with open(lst, "w", encoding="utf-8") as fh:
        fh.write("".join("file '%s'\n" % c.replace("\\", "/") for c in clips))
    joined = os.path.join(TMP, "joined.mp4")
    V.ffmpeg(["-f", "concat", "-safe", "0", "-i", lst, "-c", "copy", joined])
    srt_path = OUT[:-4] + ".srt"
    with open(srt_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(srt))
    meta_path = os.path.join(TMP, "chapters.txt")
    with open(meta_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(meta) + "\n")
    V.ffmpeg(["-i", joined, "-i", srt_path, "-i", meta_path, "-map", "0", "-map", "1", "-map_metadata", "2", "-c", "copy", "-c:s", "mov_text",
              "-metadata:s:s:0", "language=rus", "-metadata:s:a:0", "language=rus", "-metadata", "title=Однолинейка v5 — видеогайд", OUT])
    print("готово: %s (%.0f с, %d шагов)" % (OUT, t, len(STEPS)))


if __name__ == "__main__":
    main()
