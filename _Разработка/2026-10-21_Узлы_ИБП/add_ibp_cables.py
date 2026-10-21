# Кабели ИБП в книгу ЩГП: коды марок (Справочная), строки «Исходные данные», места в конце «Разбивки» (→ КЖ),
# строки спецификации (кол-во — из КЖ формулой), QS2 в п. 1.8. Правит локальную копию в своём скрытом Excel.
import os, sys
import win32com.client

p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "xl", "Однолинейка_ЩГП.xlsx")

# код → (марка с сечением, имя линии, наименование)
CODES = [
    ("ибпдс", "ПуГВнг(А)-LS 3х(1х50)"),
    ("ибппе", "ПуГВнг(А)-LS 1х25"),
    ("кв4", "КВВГнг(А)-LS 4х1,0"),
    ("кипэ", "КИПЭВнг(А)-LS 1х2х0,60"),
    ("utp5e", "U/UTP кат.5e нг(А)-LS 4х2х0,52"),
]
LINES = [  # имя линии, наименование, код
    ("ИБП-АКБ.1", "Шкаф АКБ (QFB) – ИБП, клеммы АКБ +, N, –", "ибпдс"),
    ("ИБП-АКБ.2", "Шкаф АКБ – ИБП, PE", "ибппе"),
    ("ИБП-АКБ.4", "ИБП (порт п. 10) – доп. контакт QFB шкафа АКБ", "кв4"),
    ("ИБП-EPO", "ИБП (EPO) – кнопка аварийного отключения SB1", "кв4"),
    ("ИБП-АСУ.1", "ИБП (RS485) – АСУ", "кипэ"),
    ("ИБП-АСУ.2", "ИБП (SNMP) – АСУ", "utp5e"),
]
RAZB_ROW0 = 101          # «подальше» в Разбивке: места ниже схемы
SPEC = [  # (наименование, марка для SUMIF в КЖ, множитель жил)
    ("Кабель ПуГВнг(А)-LS 1х50 (ИБП-АКБ.1: +, N, –; 3 жилы)", "ПуГВнг(А)-LS 3х(1х50)", 3),
    ("Кабель ПуГВнг(А)-LS 1х25 жёлто-зелёный (ИБП-АКБ.2, PE)", "ПуГВнг(А)-LS 1х25", 1),
    ("Кабель КВВГнг(А)-LS 4х1,0 (ИБП-АКБ.4, ИБП-EPO)", "КВВГнг(А)-LS 4х1,0", 1),
    ("Кабель КИПЭВнг(А)-LS 1х2х0,60 (ИБП-АСУ.1, RS485)", "КИПЭВнг(А)-LS 1х2х0,60", 1),
    ("Кабель U/UTP кат.5e нг(А)-LS 4х2х0,52 (ИБП-АСУ.2, SNMP)", "U/UTP кат.5e нг(А)-LS 4х2х0,52", 1),
]

xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False
xl.EnableEvents = False
try:
    wb = xl.Workbooks.Open(p, UpdateLinks=0, ReadOnly=False)
    # 1. коды марок
    sp = wb.Worksheets("Справочная")
    have = {str(sp.Cells(r, 12).Value or "").strip() for r in range(4, 52)}
    r = 4
    for code, mark in CODES:
        if code in have:
            continue
        while sp.Cells(r, 12).Value not in (None, ""):
            r += 1
        if r > 51:
            raise SystemExit("Кабель_обозначение: нет места")
        sp.Cells(r, 12).Value, sp.Cells(r, 13).Value, sp.Cells(r, 14).Value = code, mark, 2
        print("код", r, code, mark)
    # 2. Исходные данные: первые свободные строки после последней заполненной D
    ws = wb.Worksheets("Исходные данные")
    names = {str(ws.Cells(i, 4).Value or "") for i in range(3, 1001)}
    last = max(i for i in range(3, 1001) if ws.Cells(i, 4).Value not in (None, ""))
    r = last + 2
    for name, desc, code in LINES:
        if name in names:
            print("уже есть", name)
            continue
        if any(str(ws.Cells(r, c).Formula) != "" for c in (3, 4, 14, 15)):
            raise SystemExit("строка Исходных данных занята: %d" % r)
        ws.Cells(r, 3).Value = desc        # C наименование (ручное)
        ws.Cells(r, 4).Value = name        # D номер линии
        ws.Cells(r, 6).Value = 400         # F напряжение
        ws.Cells(r, 7).Value = 1           # G кол-во потребителей
        ws.Cells(r, 8).Value = 3           # H фазы (без мощности — на нагрузку не влияет)
        ws.Cells(r, 9).ClearContents()     # I длина по плану — пусто (в шаблоне заглушка =9+8+7+9+8)
        ws.Cells(r, 10).Value = 0          # J лоток
        ws.Cells(r, 11).Value = 0          # K гофра ПВХ (в шаблоне =M-5-J)
        ws.Cells(r, 12).Value = 0          # L гофра ПНД
        ws.Cells(r, 14).Value = code       # N марка кабеля (код)
        ws.Cells(r, 15).Value = 1          # O жил (марка задана целиком)
        ws.Cells(r, 52).Value = "Нет"      # AZ группа Кс
        print("исходные", r, name)
        r += 1
    # 3. Разбивка: места в конце
    rz = wb.Worksheets("Разбивка_по_щитам")
    for k, (name, _, _) in enumerate(LINES):
        rr = RAZB_ROW0 + k
        if rz.Cells(rr, 8).Value not in (None, ""):
            raise SystemExit("Разбивка H%d занята" % rr)
        if rz.Cells(rr, 6).Value != "ЩГП":
            raise SystemExit("Разбивка F%d не ЩГП" % rr)
        rz.Cells(rr, 8).Value = name
    # 4. спецификация: QS2 в п. 1.8, строки кабелей после п. 2.7
    sc = wb.Worksheets("Спецификация")
    c11 = sc.Range("C11")
    if "QS2" not in str(c11.Value):
        c11.Value = "Переключатель реверсивный TwinBlock 4P 80 А I-0-II (QS2, сервисный байпас Сеть — 0 — ИБП)"
    r27 = next(i for i in range(40, 80) if str(sc.Cells(i, 2).Value or "") == "2.7")
    sc.Rows("%d:%d" % (r27 + 1, r27 + len(SPEC))).Insert()
    for k, (nm, mark, mult) in enumerate(SPEC):
        rr = r27 + 1 + k
        sc.Cells(rr, 2).Value = "2.%d" % (8 + k)
        sc.Cells(rr, 3).Value = nm
        sc.Cells(rr, 5).Value = mark
        sc.Cells(rr, 7).Value = "м"
        sc.Cells(rr, 8).Formula = ('=ROUNDUP(SUMIF(Кабельный_журнал!$S$5:$S$1000,"%s",Кабельный_журнал!$W$5:$W$1000)'
                                   '*(1+Кабельный_журнал!$D$3),0)*%d' % (mark, mult))
        sc.Cells(rr, 10).Value = "длина — из КЖ (заполнить «Исходные данные», ст. I)"
    xl.CalculateFull()
    # проверка: строки КЖ
    kj = wb.Worksheets("Кабельный_журнал")
    for i in range(5, 40):
        pv = kj.Cells(i, 16).Value
        if pv and str(pv).startswith("ИБП"):
            print("КЖ", i, pv, "|", kj.Cells(i, 17).Value, "|", kj.Cells(i, 18).Value, "|", kj.Cells(i, 19).Value, "| L", kj.Cells(i, 23).Value)
    for i in range(last + 2, last + 2 + len(LINES)):
        print("исх", i, ws.Cells(i, 4).Value, "| AE", ws.Cells(i, 31).Text, "| M", ws.Cells(i, 13).Text, "| V", ws.Cells(i, 22).Text, "| AO", ws.Cells(i, 41).Text.replace(chr(10), " "))
    for i in range(r27 + 1, r27 + 1 + len(SPEC)):
        print("спец", sc.Cells(i, 2).Value, sc.Cells(i, 3).Value, "=", sc.Cells(i, 8).Value)
    print("C11:", sc.Range("C11").Value)
    wb.Save()
    wb.Close(False)
finally:
    xl.Quit()
