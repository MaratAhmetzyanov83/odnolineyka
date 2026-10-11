# -*- coding: utf-8 -*-
"""Тестовая книга «Однолинейка_тест_сценарии.xlsx»: шаблон v3.12 + все сценарии заполнения.

    py make_scenarios.py [выход.xlsx]   (по умолчанию 1_Таблицы_Excel/Однолинейка_тест_сценарии.xlsx)

Щит ЩТ — правильные сценарии (памятка «Как_заполнять_строки.md» + сечение/потери/24 В),
щит ЩТ2 — намеренные ошибки (статус строки должен их показать).
Лист «Сценарии»: что проверяем, что ввели, что ждём, что получилось, OK.
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
TPL = os.path.join(ROOT, "1_Таблицы_Excel", "Однолинейка_шаблон.xlsx")
OUT = os.path.join(ROOT, "1_Таблицы_Excel", "Однолинейка_тест_сценарии.xlsx")
PWD = "sx"
SRC0 = 21          # первая строка линий в «Исходных данных»
ODN0 = 3           # первая строка мест в «Разбивке» / «Однолинейке»

# линия: исходные данные (столбец -> значение)
def ln(num, e, i=10, **kw):
    d = {"D": num, "E": e, "I": i}
    d.update(kw)
    return d

R230 = dict(F=230, H=1)
X = dict(N="мс", O=3, F=230, H=1)     # коды без марки / жил / напряжения в «Справочной» — вписываем сами

# место: (щит, линия или None, ввод в «Однолинейку», сценарий, проверка)
# проверка: функция(row) -> (получилось, ok) ; row — словарь значений по линии / месту
PLACES = []


def place(board, line, odn, what, check=None):
    PLACES.append((board, line, odn, what, check))


def v(r, k):
    return r.get(k)


def num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- щит ЩТ
for _ in range(5):
    place("ЩТ", None, {"Y": "вручную"}, "Пустое место вводного узла (вручную)")
place("ЩТ", ln("Ввод.1", 15, 20, F=400, H=3, N="мс", O=5), {"E": "авто", "P": "XT1"},
      "Автоматический ввод (Ввод.1): петля кабеля от питающего щита",
      lambda r: (f"{v(r,'H')}; {v(r,'Q')}", bool(v(r, "H"))))
place("ЩТ", ln("R.1.01.1", 0.3, 12, G=5), {"E": "авто", "P": "XT2"},
      "№2 Линия со своим автоматом",
      lambda r: (f"{v(r,'G')} {v(r,'H')}", bool(v(r, "H"))))
place("ЩТ", ln("R.1.02.1", 0.24, 15), {"E": "авто", "P": "XT3"}, "№3 Группа: первая строка — автомат",
      lambda r: (f"{v(r,'G')} {v(r,'H')}", bool(v(r, "H"))))
place("ЩТ", ln("R.1.02.2", 0.18, 18), {"P": "XT4"}, "№3/№4 Группа: линия на автомате выше",
      lambda r: (f"QF {v(r,'G')}", bool(v(r, "G")) and not v(r, "H")))
place("ЩТ", ln("R.1.02.3", 0.12, 20), {"P": "XT5"}, "№3 Группа: третья линия, перемычка клемм",
      lambda r: (f"QF {v(r,'G')}", bool(v(r, "G"))))
place("ЩТ", ln("R.1.03.1", 0.5, 10), {"E": "1P C16A", "F": "нет", "P": "XT6"},
      "Свой текст автомата, без 30 мА", lambda r: (str(v(r, "H")), v(r, "H") == "1P C16A"))
place("ЩТ", ln("R.1.03.2", 0.3, 10), {"E": "авто", "F": "да", "P": "XT7"}, "30 мА — да (явно)",
      lambda r: (str(v(r, "H")), "30мА" in str(v(r, "H"))))
place("ЩТ", ln("R.1.04.1", 7, 16, F=400, H=3, O=5, C="Варочная панель"), {"E": "авто", "P": "XT8"},
      "3 фазы, 7 кВт: 4P, фазы L1,L2,L3, 5 жил",
      lambda r: (f"{v(r,'H')}, фаза {v(r,'V')}, {v(r,'Q')}", str(v(r, "V")) == "L1,L2,L3" and "4P" in str(v(r, "H"))))
place("ЩТ", ln("R.1.04.2", 7, 16, C="Духовой шкаф"), {"E": "авто", "P": "XT26"},
      "1 фаза, 7 кВт: сечение по току больше минимума по коду (2,5)",
      lambda r: (f"Iр {num(v(r,'sT')) and round(num(v(r,'sT')),1)} А, In {v(r,'sU')}, S {v(r,'sV')}",
                 (num(v(r, "sV")) or 0) > 2.5))
place("ЩТ", ln("R.1.05.1", 2, 150), {"E": "авто", "P": "XT9"},
      "Длинная линия 150 м, 2 кВт: сечение растёт по потерям (≤ 4,5 %)",
      lambda r: (f"S {v(r,'sV')}, ΔU {v(r,'sAB')} %", (num(v(r, "sV")) or 0) > 2.5 and (num(v(r, "sAB")) or 99) <= 4.5))
place("ЩТ", ln("R.1.05.2", 0.3, 10, Q=4), {"E": "авто", "P": "XT10"}, "Ручное сечение Q = 4",
      lambda r: (f"S {v(r,'sV')}", num(v(r, "sV")) == 4))
place("ЩТ", ln("R.1.05.3", 0.3, 10, R=25), {"E": "авто", "P": "XT11"}, "Ручной автомат R = 25 А",
      lambda r: (f"In {v(r,'sU')}, {v(r,'H')}", num(v(r, "sU")) == 25))
place("ЩТ", ln("R.1.06.1", 0.2, 10), {"Y": "вручную"}, "№1 Только в таблицу (вручную): на схеме нет",
      lambda r: (f"линия {v(r,'R')}, автомат «{v(r,'H') or ''}»", v(r, "R") == "R.1.06.1" and not v(r, "H")))
# лист 2 (места 18–34)
place("ЩТ", ln("R.1.07.1", 0.4, 12), {"D": "ИБП", "E": "авто", "P": "XT12"}, "№5 Вторая шина ИБП",
      lambda r: (f"шина {v(r,'D')}, {v(r,'H')}", bool(v(r, "H"))))
place("ЩТ", ln("R.1.07.2", 0.3, 12), {"D": "ИБП", "E": "авто", "P": "XT13"}, "№5 Вторая шина ИБП, вторая линия",
      lambda r: (f"шина {v(r,'D')}", bool(v(r, "H"))))
place("ЩТ", ln("L.1.01.1", 0.1, 10, G=4), {"E": "авто", "L": "ZIOMB24V2", "P": "XT14"},
      "№6 Модуль: первая строка, канал по порядку",
      lambda r: (f"модуль {v(r,'M')} канал {v(r,'O')}", bool(v(r, "M"))))
place("ЩТ", ln("L.1.01.2", 0.1, 12), {"P": "XT15"}, "№6 Модуль: следующий канал",
      lambda r: (f"модуль {v(r,'M')} канал {v(r,'O')}", num(v(r, "O")) == 2))
place("ЩТ", ln("L.1.01.3", 0.05, 14), {"N": "=", "P": "XT16"}, "№7 Параллельная линия на тот же канал (=)",
      lambda r: (f"канал {v(r,'O')}", num(v(r, "O")) == 2))
place("ЩТ", None, {"N": "р"}, "№8 Резерв канала (р)",
      lambda r: (f"канал {v(r,'O')}", v(r, "O") not in (None, "")))
place("ЩТ", ln("L.1.01.4", 0.1, 10), {"N": 6, "P": "XT17"}, "Канал вручную (N = 6)",
      lambda r: (f"канал {v(r,'O')}", num(v(r, "O")) == 6))
place("ЩТ", None, {"N": "+", "P": "XT18"}, "№9 Ещё одна клемма в тот же кабель (+)",
      lambda r: (f"канал {v(r,'O')}, клемма {v(r,'P')}", v(r, "O") not in (None, "")))
place("ЩТ", ln("L.1.01.5", 0.1, 10), {"E": "авто", "L": "-", "P": "XT19"}, "№6 Конец модуля (L = -): строка уже без модуля",
      lambda r: (f"модуль «{v(r,'M') or ''}»", not v(r, "M")))
place("ЩТ", ln("LED.1.01.1", 0.1, 10, N="мс", O=2, F=24, H=1, C="Подсветка 24 В"),
      {"E": "авто", "I": "БП 24V 150W", "P": "XT20"},
      "№11 Линия 24 В через блок питания: потери от 24 В (≤ 5 %)",
      lambda r: (f"U {v(r,'sAC')} В, S {v(r,'sV')}, ΔU {v(r,'sAB')} %",
                 num(v(r, "sAC")) == 24 and (num(v(r, "sAB")) or 99) <= 5))
place("ЩТ", ln("DALI.1.01.1", 0, 30), {"E": "авто", "I": "БП DALI", "N": "б/к"},
      "Слаботочная: шина DALI от БП DALI (б/к) — без тока и нагрузки, статус чистый",
      lambda r: (f"ток {v(r,'sT')}, слаб. {v(r,'BZ')}, статус «{v(r,'AE') or ''}»",
                 num(v(r, "sT")) == 0 and num(v(r, "BZ")) == 1 and not v(r, "AE")))
place("ЩТ", ln("KNX.1.01.1", 0, 40), {"E": "авто", "I": "PS-KNX 640мА", "P": "XT41"},
      "Слаботочная: шина KNX от PS-KNX — кабель по коду, клемма без PE",
      lambda r: (f"{v(r,'Q')}, ток {v(r,'sT')}, статус «{v(r,'AE') or ''}»",
                 "J-Y" in str(v(r, "Q")) and num(v(r, "sT")) == 0 and not v(r, "AE")))
place("ЩТ", ln("В.1.01.1", 0.06, 8, C="Вытяжной вентилятор с/у", **X), {"E": "авто", "I": "КМ", "P": "XT21"},
      "Двигатель (cos φ по справочнику) + контактор КМ после автомата",
      lambda r: (f"cos φ {v(r,'sP')}, {v(r,'I')}", bool(v(r, "sP"))))
place("ЩТ", ln("ТП.1.01.1", 1.2, 12, **X), {"E": "авто", "P": "XT22"},
      "Тёплый пол: УГО «Нагреватель», 30 мА по коду",
      lambda r: (str(v(r, "H")), "30мА" in str(v(r, "H"))))
place("ЩТ", ln("K.1.01.1", 0.8, 14, **X), {"E": "авто", "I": "МК-5-1", "P": "XT23"},
      "Аппарат МК под автоматом (кондиционер)", lambda r: (f"{v(r,'I')}", bool(v(r, "H"))))
place("ЩТ", ln("АСУ.1.01.1", 0.5, 6, C="Щит автоматизации ЩА", **X), {"E": "авто", "P": "XT24"},
      "№12 «Прочее»: рамка с названием (нижестоящий щит)", lambda r: (str(v(r, "S")), bool(v(r, "H"))))
place("ЩТ", ln("LD.1.01.1", 0.4, 12, **X), {"E": "авто", "I": "Диммер 4к 4x150W", "P": "XT42"},
      "Диммер 4-канальный под автоматом (значок SX_DIM4)",
      lambda r: (f"аппарат «{v(r,'I')}», статус «{v(r,'AE') or ''}»", bool(v(r, "H"))))
place("ЩТ", ln("XYZ.1.01.1", 0.1, 10, **X), {"E": "авто", "P": "XT25"},
      "№12 Неизвестный код: красный кабельный вывод", lambda r: (f"статус «{v(r,'AE') or ''}»", True))
place("ЩТ", None, {}, "Пустое место между группами")
# лист 3+: ещё линии, чтобы в щите было больше 27 (ошибка фаз v3.12 в строках 29+)
for k in range(1, 15):
    extra = {"S": 2} if k == 3 else {}
    odn = {"E": "авто", "P": f"XT{26 + k}"}
    what = f"Линия № {k} сверх 27-й (фаза — от своей строки)"
    if k == 3:
        what = "Прокладка по полу (S = 2): опуск в гофру ПНД"
    if k == 5:
        what = "Ручная фаза L2 на «Нагрузка_щитов»"
    place("ЩТ", ln(f"R.2.{k:02d}.1", 0.2 + 0.05 * (k % 4), 10 + k, **extra), odn, what,
          (lambda r: (f"фаза {v(r,'V')}", v(r, "V") == "L2")) if k == 5 else
          (lambda r: (f"ПВХ {v(r,'sK')}, ПНД {v(r,'sL')}", (num(v(r, "sL")) or 0) > 0)) if k == 3 else
          (lambda r: (f"фаза {v(r,'V')}", bool(v(r, "V")))))

# ---------------------------------------------------------------- щит ЩТ2: ошибки
place("ЩТ2", ln("R.3.01.1", 0.2, 10), {"P": "XT9"}, "ОШИБКА: клемма без автомата выше — SXDRAW предупредит до рисования",
      lambda r: ("проверить в SXDRAW", None))
place("ЩТ2", ln("R.3.01.2", 0.2, 10), {"E": "авто", "P": "XT1"}, "ОШИБКА: дубль клеммы XT1",
      lambda r: (f"статус «{v(r,'AE') or ''}»", "дубль" in str(v(r, "AE"))))
place("ЩТ2", ln("KNX.1.01.2", 0, 20), {"E": "авто", "P": "XT5"},
      "ОШИБКА: слаботочная линия с автоматом без промежуточного устройства",
      lambda r: (f"статус «{v(r,'AE') or ''}»", "слаботочная" in str(v(r, "AE"))))
place("ЩТ2", ln("R.3.01.3", 0, 10), {"E": "авто", "P": "XT1"}, "ОШИБКА: нет мощности (и та же XT1)",
      lambda r: (f"статус «{v(r,'AE') or ''}»", "мощност" in str(v(r, "AE"))))

# линия только в «Исходных данных» — нет в «Разбивке»
ORPHAN = ln("R.9.01.1", 0.2, 10)

ODN_COLS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
            "U", "V", "W", "X", "Y", "AE"]
BZ = 77   # индекс «слаботочн.» (BZ) в прочитанном диапазоне A:BZ
SRC_COLS = {"sT": "T", "sU": "U", "sV": "V", "sAB": "AB", "sAC": "AC", "sP": "P", "sK": "K", "sL": "L", "sAD": "AD"}


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def unprot(ws):
    if ws.ProtectContents:
        ws.Unprotect(PWD)


def build(out):
    shutil.copyfile(TPL, out)
    app, pid = excel()
    try:
        wb = app.Workbooks.Open(out, 0, False)
        app.Calculation = -4135
        src, rz, odn, nag = (wb.Worksheets(n) for n in ("Исходные данные", "Разбивка_по_щитам", "Однолинейка",
                                                         "Нагрузка_щитов"))
        for ws in (src, rz, odn, nag):
            unprot(ws)
        rz.Range("F3:F1000").ClearContents()
        rz.Range("H3:H1000").ClearContents()
        for c in ("D", "E", "F", "I", "J", "K", "L", "N", "P", "Y"):
            odn.Range(f"{c}3:{c}1000").ClearContents()
        lines = [p[1] for p in PLACES if p[1]] + [ORPHAN]
        for i, d in enumerate(lines):
            r = SRC0 + i
            for c, val in d.items():
                src.Range(f"{c}{r}").Value = val
        for i, (board, line, inp, _, _) in enumerate(PLACES):
            r = ODN0 + i
            rz.Range(f"F{r}").Value = board
            if line:
                rz.Range(f"H{r}").Value = line["D"]
            for c, val in inp.items():
                odn.Range(f"{c}{r}").Value = val
        app.Calculation = -4105
        app.CalculateFull()
        # ручная фаза L2 на «Нагрузка_щитов» для R.2.05.1
        dcol = [x[0] for x in nag.Range("D3:D1000").Value]
        if "R.2.05.1" in dcol:
            nag.Range(f"E{3 + dcol.index('R.2.05.1')}").Value = 2
        app.CalculateFull()
        # прочитать результаты
        oval = odn.Range(f"A{ODN0}:BZ{ODN0 + len(PLACES) - 1}").Value
        sval = {n: src.Range(f"{c}{SRC0}:{c}{SRC0 + len(lines) - 1}").Value for n, c in SRC_COLS.items()}
        sidx = {d["D"]: i for i, d in enumerate(lines)}
        rows, nok = [], 0
        for i, (board, line, inp, what, check) in enumerate(PLACES):
            rec = {c: oval[i][j] for j, c in enumerate(ODN_COLS[:-1])}
            rec["AE"] = oval[i][30]
            rec["BZ"] = oval[i][BZ]
            if line:
                for n in SRC_COLS:
                    rec[n] = sval[n][sidx[line["D"]]][0]
            got, ok = ("", None)
            if check:
                try:
                    got, ok = check(rec)
                except Exception as e:  # noqa: BLE001
                    got, ok = f"ошибка проверки: {e}", False
            if ok is False:
                nok += 1
            entered = "; ".join(f"{k}={val}" for k, val in inp.items())
            if line:
                entered = f"линия {line['D']}: " + ", ".join(f"{k}={val}" for k, val in line.items() if k != "D") + \
                          (" | " + entered if entered else "")
            rows.append([i + 1, board, rec.get("B"), what, entered, got, "" if ok is None else ("OK" if ok else "НЕТ"),
                         rec.get("AE") or ""])
        # лист «Сценарии»
        try:
            ws = wb.Worksheets("Сценарии")
            ws.Cells.Clear()
        except Exception:  # noqa: BLE001
            ws = wb.Worksheets.Add(Before=wb.Worksheets(1))
            ws.Name = "Сценарии"
        ws.Range("A1").Value = ("ТЕСТОВЫЕ СЦЕНАРИИ. Щит ЩТ — правильное заполнение, ЩТ2 — намеренные ошибки. "
                                "Линия R.9.01.1 есть только в «Исходных данных» (её нет в «Разбивке»). "
                                "Схему проверить: SXDRAW → эта книга → щит ЩТ, потом ЩТ2.")
        hdr = ["№", "Щит", "Место", "Что проверяем", "Что ввели", "Получилось", "OK", "Статус строки (Однолинейка AE)"]
        ws.Range("A3:H3").Value = hdr
        ws.Range("A3:H3").Font.Bold = True
        ws.Range("A4").Resize(len(rows), 8).Value = rows
        for c, w in zip("ABCDEFGH", (5, 6, 7, 55, 60, 40, 6, 40)):
            ws.Columns(c).ColumnWidth = w
        ws.Range(f"A3:H{3 + len(rows)}").WrapText = True
        ws.Range(f"A3:H{3 + len(rows)}").VerticalAlignment = -4160
        for i, r in enumerate(rows):
            if r[6] == "НЕТ":
                ws.Range(f"G{4 + i}").Interior.Color = 0x9C9CFF
            elif r[6] == "OK":
                ws.Range(f"G{4 + i}").Interior.Color = 0xCEEFC6
        ws.Activate()
        wb.Save()
        wb.Close(False)
        print(f"книга: {out}\nмест: {len(PLACES)}, линий: {len(lines)}, проверок НЕТ: {nok}")
        for r in rows:
            if r[6] == "НЕТ" or r[7]:
                print(" ", r[0], r[1], r[3][:60], "|", r[5], "|", r[6], "|", r[7])
    finally:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    build(os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else OUT)
