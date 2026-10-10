# -*- coding: utf-8 -*-
"""1.05 -> 1.06: исправления расчётов без изменения вида таблицы.

Запуск:  py patch_106.py <вход 1.05.xlsx> <выход 1.06.xlsx>
Работает через свой скрытый экземпляр Excel (не трогает открытые окна пользователя).
Функция patch(wb) используется и в сравнении (compare_106.py).
"""
import os
import re
import sys

ROWS = (3, 1000)            # строки данных «Исходных данных»
ELV = "AND(ISNUMBER(F{r}),F{r}>0,F{r}<=50,F{r}<>16)"   # линия до 50 В — постоянный ток (24/12 В); 16 В — шина DALI, считается отдельно (0)

# ПУЭ табл. 1.3.6, медь, в воздухе; 1ж / 2ж (1ф: L+N) / 3ж (3ф). 240 — только одножильные.
PUE = [
    (1.5, 23, 19, 19), (2.5, 30, 27, 25), (4, 41, 38, 35), (6, 50, 50, 42),
    (10, 80, 70, 55), (16, 100, 90, 75), (25, 140, 115, 95), (35, 170, 140, 120),
    (50, 215, 175, 145), (70, 270, 215, 180), (95, 325, 260, 220), (120, 385, 300, 260),
    (150, 440, 350, 305), (185, 510, 405, 350), (240, 605, None, None),
]
RHO_OLD, RHO_NEW = 0.0175, 0.0225   # ρ меди 20 °C -> ρ1 по ГОСТ Р 50571.5.52 прил. G

# Диаметры (мм) для гофры: 5х10 — каталог ВВГнг(А)-LS 0,66 кВ (ок); 5х16 — значение
# из таблиц «Жуковки»; 5х25, 5х95 — каталог АВВГнг-LS той же конструкции (уточнить).
DIAMS = [
    ("ВВГнг(А)-LS", 5, 10, 17.4, "каталог ВВГнг(А)-LS 0,66 кВ"),
    ("ВВГнг(А)-LS", 5, 16, 21.26, "из таблиц «Жуковки»"),
    ("ВВГнг(А)-LS", 5, 25, 23.9, "по каталогу АВВГнг-LS"),
    ("ВВГнг(А)-LS", 5, 95, 41.9, "по каталогу АВВГнг-LS"),
]

LOG_106 = ("1. Ток линии с cos φ, напряжение 230/400 В; линии до 50 В (24/12 В) — на своём напряжении; у шины DALI (16 В) ток, потери и мощность не считаются и не идут в нагрузку щита.\n"
           "2. Сечение по ПУЭ табл. 1.3.6 по числу нагруженных жил (1ф — двухжильные, 3ф — трёхжильные, "
           "4–5 жил ×0,92), In ≤ Iдоп («Настройки»).\n"
           "3. Потери по ГОСТ Р 50571.5.52 прил. G: итоговая длина, cos φ, ρ = 0,0225; % от 230 В "
           "(до 50 В — от своего напряжения, допустимо 5 %).\n"
           "4. Фаза в «Разбивке» берётся по строке, а не по номеру линии (резервы больше не на одной фазе).\n"
           "5. Iр щита через 0,4 кВ; Кс и cos φ в «В Акад» — 2 знака.\n"
           "6. Расчет_БП: итоговая длина и итоговое сечение; мощность БП подписана в Вт (считалась в Вт, подпись была кВт).\n"
           "7. Заголовки K/L «Гофра ПВХ»/«Гофра ПНД»; в справочник диаметров добавлены 5х10, 5х16, 5х25, 5х95.\n"
           "8.Столбцы для DLN (AW:BA): длины ПВХ/ПНД с плана, хэндлы, примечание, текст выноски. Команды DLN (заполнить) и DLN_UPD (обновить).")


def col_formulas(rng):
    """Формулы столбца как список строк."""
    v = rng.Formula
    return [row[0] for row in v] if isinstance(v, tuple) else [v]


def set_col(rng, formulas):
    rng.Formula = tuple((f,) for f in formulas)


def add_name(wb, name, ref):
    try:
        wb.Names(name).Delete()
    except Exception:
        pass
    wb.Names.Add(Name=name, RefersTo=ref)


def patch(wb, log=print):
    r0, r1 = ROWS
    sp = wb.Worksheets("Справочная")
    src = wb.Worksheets("Исходные данные")

    # --- Справочная: таблица ПУЭ и настройки (AL:AP — свободно в 1.05) ---
    if sp.Range("AL2:AP24").Cells.Count != 0 and sp.Application.WorksheetFunction.CountA(sp.Range("AL2:AP24")) > 0:
        if sp.Range("AL3").Value != "Сечение, мм²":
            raise RuntimeError("«Настройки» занята — патч остановлен")
    sp.Range("AL2").Value = "ДОПУСТИМЫЙ ДЛИТЕЛЬНЫЙ ТОК, А (ПУЭ табл. 1.3.6, медь, в воздухе) — для подбора сечения, 1.06"
    sp.Range("AL3:AO3").Value = (("Сечение, мм²", "Одножильные", "Двухжильные (1ф: L+N)", "Трёхжильные (3ф); 4–5 жил ×0,92"),)
    for i, row in enumerate(PUE):
        # пусто нельзя: XLOOKUP считает пустую ячейку больше любого числа; 0 совпал бы при токе 0 -> −1, показ «—»
        sp.Range(f"AL{4+i}:AO{4+i}").Value = (tuple(-1 if x is None else x for x in row),)
    sp.Range("AM4:AO18").NumberFormatLocal = '0;"—";"—"'
    sp.Range("AL20").Value = ("Сечение = наименьшее, у которого Iдоп·k ≥ номинала автомата (ГОСТ Р 50571.4.43: "
                              "Ib ≤ In ≤ Iz); k = 0,92 для 4–5-жильных в 3ф (ПУЭ 1.3.10). Сечение (ручное) — главнее.")
    sp.Range("AL22").Value = "Uн линейное, В (ГОСТ 29322: 230/400)"
    sp.Range("AM22").Value = 400
    sp.Range("AL23").Value = "Допустимые потери до 50 В (LED 24/12 В), %"
    sp.Range("AM23").Value = 5
    sp.Range("AL24").Value = "ρ меди для потерь — столбец D (0,0225 по ГОСТ Р 50571.5.52 прил. G)"
    sp.Range("AL3:AO3").Font.Bold = True
    sp.Range("AL2").Font.Bold = True
    add_name(wb, "Iдоп_S", "=Справочная!$AL$4:$AL$18")
    add_name(wb, "Iдоп_1ж", "=Справочная!$AM$4:$AM$18")
    add_name(wb, "Iдоп_2ж", "=Справочная!$AN$4:$AN$18")
    add_name(wb, "Iдоп_3ж", "=Справочная!$AO$4:$AO$18")
    add_name(wb, "Uн_В", "=Справочная!$AM$22")
    add_name(wb, "Потери_до50В", "=Справочная!$AM$23")
    # ρ: 0,0175 -> 0,0225 (столбец D «Кабель_сопротивление»)
    n = 0
    for r in range(4, 19):
        c = sp.Range(f"D{r}")
        if c.Value == RHO_OLD:
            c.Value = RHO_NEW
            n += 1
    log(f"Справочная: таблица ПУЭ, настройки, ρ заменено в {n} строках")

    # --- Исходные данные ---
    if str(src.Range("K2").Value).strip() == "ПНД" and "ПВХ" in str(src.Range("L2").Value):
        src.Range("K2").Value = "Гофра ПВХ"
        src.Range("L2").Value = "Гофра ПНД"
        log("Исходные данные: заголовки K/L исправлены")

    # столбцы для DLN (он ищет их по заголовку; без них создавал бы справа от последнего — в AU)
    hdr = [str(x or "") for x in src.Range("A2:DZ2").Value[0]]
    if not any(h.startswith("DWG:") for h in hdr):
        if src.Application.WorksheetFunction.CountA(src.Range("AW2:BA1000")) == 0:
            src.Range("AW2:BA2").Value = (("DWG: гофра ПВХ по плану, м", "DWG: гофра ПНД по плану, м",
                                           "DWG: хэндлы полилиний", "DWG: примечание",
                                           "Доп. информация из выноски"),)
            src.Range("AW2:BA2").Font.Bold = True
            src.Range("AW2:BA2").WrapText = True
            log("Исходные данные: столбцы DLN AW:BA")

    elv = ELV.format(r=r0)
    src.Range(f"T{r0}:T{r1}").Formula2 = (
        f"=LET(cosf,IF(ISNUMBER(P{r0}),IF(P{r0}>0,P{r0},0.95),0.95),"
        f"IF(F{r0}=16,0,IF({elv},E{r0}*1000/F{r0},"
        f"IFS(H{r0}=1,E{r0}*1000/(Uн_В/SQRT(3)*cosf),H{r0}=3,E{r0}*1000/(SQRT(3)*Uн_В*cosf)))))")

    # V: в 1.05 формула в строках разная (4 строки с проверкой марки, остальные без) —
    # в каждой строке меняем только подбор по старой таблице «Кабель_номиналы».
    old_pick = "VLOOKUP(RC[-1],Кабель_номиналы,2,TRUE)"
    pick = ('IFERROR(XLOOKUP(RC[-1]/IF(AND(RC[-14]=3,RC[-7]>=4),0.92,1),'
            'IF(RC[-7]=1,Iдоп_1ж,IF(RC[-14]=3,Iдоп_3ж,Iдоп_2ж)),Iдоп_S,"Не исп",1),"Не исп")')
    rng = src.Range(f"V{r0}:V{r1}")
    fl = [x[0] for x in rng.FormulaR1C1]
    n = sum(1 for f in fl if isinstance(f, str) and old_pick in f)
    rng.Formula2R1C1 = tuple(((f.replace(old_pick, pick) if isinstance(f, str) else f),) for f in fl)
    log(f"Исходные данные: подбор сечения заменён в {n} строках")

    src.Range(f"AA{r0}:AA{r1}").Formula2 = (
        f"=IF(F{r0}=16,0,IF({elv},2*W{r0}*M{r0}/V{r0}*T{r0},"
        f"Y{r0}*(W{r0}*M{r0}/V{r0}*P{r0}+X{r0}*M{r0}*Z{r0}/1000)*T{r0}))")
    # мощность для «Разбивки»/«В Акад»: у шины DALI не считается
    src.Range(f"AL{r0}:AL{r1}").Formula = f"=IF(F{r0}=16,0,E{r0})"
    src.Range(f"AB{r0}:AB{r1}").Formula2 = f"=ROUNDUP(AA{r0}/IF({elv},F{r0},Uн_В/SQRT(3))*100,2)"
    src.Range(f"AF{r0}:AF{r1}").Formula2 = f"=IFERROR(IF(AB{r0}>IF({elv},Потери_до50В,4.5),1,2),0)"

    notes = {
        "T2": "1.06: Iр = P/(Uф·cos φ) для 1ф, P/(√3·Uн·cos φ) для 3ф; Uн = 400 В («Настройки»). "
              "Линии до 50 В (24/12 В в «Напряжении») — постоянный ток: P/U. 16 В — шина DALI: ток, потери и мощность не считаются.",
        "V2": "1.06: наименьшее сечение с Iдоп·k ≥ номинала автомата (ПУЭ табл. 1.3.6 — «Настройки»). "
              "1ф — двухжильные, 3ф — трёхжильные, 4–5 жил ×0,92. «Сечение (ручное)» главнее.",
        "AA2": "1.06: ГОСТ Р 50571.5.52 прил. G: ΔU = b·(ρ1·L/S·cos φ + x·L·sin φ)·I; b = 2 (1ф), 1 (3ф); "
               "L — итоговая длина (M); ρ1 = 0,0225. До 50 В: ΔU = 2·ρ1·L·I/S.",
        "AB2": "1.06: % от фазного 230 В (до 50 В — от своего напряжения).",
    }
    for addr, text in notes.items():
        c = src.Range(addr)
        try:
            if c.Comment is not None:
                c.Comment.Delete()
        except Exception:
            pass
        c.AddComment(text)
    log("Исходные данные: T, V, AA, AB, AF, AL переписаны, примечания к заголовкам")

    # --- Разбивка: фаза по строке (номер E), а не по номеру линии ---
    rz = wb.Worksheets("Разбивка_по_щитам")
    rz.Range("M3:M1000").Formula = '=IF(E3=0,"",IFERROR(INDEX(Нагрузка_щитов!$G$3:$G$1000,E3),""))'
    log("Разбивка: фаза по строке")

    # --- Нагрузка_щитов: мощность шины DALI (16 В) не идёт в нагрузку щита ---
    nz = wb.Worksheets("Нагрузка_щитов")
    old_p = "VLOOKUP(RC[-4],Исходные_данные,5,FALSE)"
    new_p = "IF(VLOOKUP(RC[-4],Исходные_данные,6,FALSE)=16,0,VLOOKUP(RC[-4],Исходные_данные,5,FALSE))"
    rng = nz.Range("H3:H1000")
    fl = [x[0] for x in rng.FormulaR1C1]
    n = sum(1 for f in fl if isinstance(f, str) and f == "=" + old_p)
    rng.FormulaR1C1 = tuple((("=" + new_p if f == "=" + old_p else f),) for f in fl)
    log(f"Нагрузка_щитов: мощность DALI исключена (формула заменена в {n} строках)")

    # --- Нагрузка_щитов: Iр через 0,4 кВ ---
    used = nz.Range("O1:AC1000")
    fs = used.Formula
    k = 0
    new = []
    for row in fs:
        nr = []
        for f in row:
            if isinstance(f, str):
                g = re.sub(r"^=V(\d+)/1\.73/0\.38/V(\d+)$", r"=V\1/SQRT(3)/0.4/V\2", f)
                # блоки 2–5: «…/V(cos)*V(Кс)» — Кс второй раз (Рр уже с Кс)
                g = re.sub(r"^=V(\d+)/0\.38/1\.732/V(\d+)\*V\d+$", r"=V\1/SQRT(3)/0.4/V\2", g)
                if g != f:
                    f = g
                    k += 1
            nr.append(f)
        new.append(tuple(nr))
    if k:
        used.Formula = tuple(new)
    log(f"Нагрузка_щитов: Iр = Рр/(√3·0,4·cos φ) в {k} блоках (в блоках 2–5 убран повторный Кс)")

    # --- В Акад: Кс, cos φ и остальные итоги — 2 знака ---
    va = wb.Worksheets("В Акад")
    k = 0
    for top in (2, 9, 16, 23, 30):
        for col in ("Q", "T"):
            for r in range(top, top + 5):
                c = va.Range(f"{col}{r}")
                f = str(c.Formula)
                if f.startswith("=") and "ROUND" not in f and "!" in f:
                    x = f[1:]
                    c.Formula = f"=IF(ISNUMBER({x}),ROUND({x},2),{x})"
                    try:
                        c.NumberFormatLocal = "0,00"
                    except Exception:
                        pass
                    k += 1
    log(f"В Акад: округление в {k} ячейках")

    # --- Расчет_БП: итоговая длина (M) и итоговое сечение (V); ρ1 ---
    bp = wb.Worksheets("Расчет_БП")
    for col, a, b in (("C", ",9,FALSE)", ",13,FALSE)"), ("D", ",17,FALSE)", ",22,FALSE)")):
        rng = bp.Range(f"{col}3:{col}107")
        fl = col_formulas(rng)
        set_col(rng, [f.replace(a, b) if isinstance(f, str) else f for f in fl])
    if str(bp.Range("G2").Value).strip().endswith("кВт"):   # J·K·1,3 — это Вт
        bp.Range("G2").Value = "Расчетная мощность БП, Вт (с запасом 1,3)"
    for r in range(3, 24):
        c = bp.Range(f"R{r}")
        if c.Value == RHO_OLD:
            c.Value = RHO_NEW
    log("Расчет_БП: длина M, сечение V, ρ1")

    # --- Справочник диаметров кабелей ---
    spp = wb.Worksheets("Справочная_потребители")
    names = [str(x[0]) if x[0] is not None else "" for x in spp.Range("T4:T503").Value]
    last = max(i for i, s in enumerate(names) if s) + 4
    added = []
    for mark, cores, sec, d, note in DIAMS:
        title = f"{mark} ({cores}х{str(sec).replace('.', ',')})"
        if title in names:
            continue
        last += 1
        spp.Range(f"Q{last}:U{last}").Value = ((mark, cores, sec, title, d),)
        spp.Range(f"V{last - 1}:X{last - 1}").Copy(spp.Range(f"V{last}:X{last}"))
        spp.Range(f"Y{last}").Value = note
        added.append(title)
    log(f"Диаметры: добавлены {added}")

    # --- Краткое руководство: версия и ведомость ---
    kr = wb.Worksheets("Краткое руководство")
    kr.Range("C1").Value = "Версия таблицы 1.06"
    r = 2
    while kr.Range(f"E{r}").Value not in (None, ""):
        r += 1
    if kr.Range(f"E{r-1}").Value != "1.06":
        kr.Range(f"E{r}").Value = "1.06"
        kr.Range(f"F{r}").Value = LOG_106
        kr.Range(f"F{r}").WrapText = True
    h2 = str(kr.Range("H2").Value or "")
    if h2 and "1.06" not in h2:
        kr.Range("H2").Value = h2 + " — исправлено в 1.06"
    log("Краткое руководство: версия 1.06")


PIDS = {}


def excel():
    import win32com.client
    app = win32com.client.DispatchEx("Excel.Application")
    try:
        import win32process
        PIDS[id(app)] = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    except Exception:
        pass
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app


def main(src, dst):
    import shutil
    src, dst = os.path.abspath(src), os.path.abspath(dst)
    shutil.copyfile(src, dst)
    app = excel()
    try:
        wb = app.Workbooks.Open(dst, UpdateLinks=0)
        app.Calculation = -4135
        patch(wb)
        app.Calculation = -4105
        app.CalculateFull()
        wb.Save()
        wb.Close(False)
    finally:
        quit_excel(app)


def quit_excel(app):
    pid = PIDS.pop(id(app), None)
    try:
        app.Quit()
    except Exception:
        pass
    if pid:
        import subprocess, time
        time.sleep(2)
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
