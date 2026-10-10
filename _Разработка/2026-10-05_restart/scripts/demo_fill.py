"""Заполнить демо-щит в книге 1.08 (для тестов и как пример)."""
# Исходные: D, E, F, G, H, I, (BA — текст выноски)
SRC_LINES = [
    ("Ввод.1", 15.0, 400, 1, 3, 10, None),
    ("R.1.01.1", 2.0, 230, 4, 1, 20, None),
    ("R.1.01.2", 1.5, 230, 3, 1, 18, None),
    ("R.1.02.1", 1.8, 230, 5, 1, 25, None),
    ("L.1.01.1", 0.3, 230, 8, 1, 22, "Relay"),
    ("L.1.01.2", 0.2, 230, 6, 1, 17, "Relay"),
    ("L.1.02.1", 0.25, 230, 5, 1, 19, "DIM"),
    ("LD.1.01.1", 0.1, 24, 1, 1, 12, None),
    ("LD.1.01.2", 0.1, 24, 1, 1, 14, None),
    ("SH.1.01.1", 0.1, 230, 1, 1, 15, None),
    ("П.1.01.1", 5.0, 400, 1, 3, 30, None),
    ("ТП.1.01.1", 1.2, 230, 1, 1, 12, None),
]
# Разбивка: щит, линия, V группа, W автомат, X дифф, Y аппараты, Z модуль, AA канал, AB клемма
RZ_ROWS = [
    ("ЩР", "Ввод.1", None, "3P C63A", None, "Реле напряжения; УЗИП", "-", None, None),
    ("ЩР", None), ("ЩР", None),
    ("ЩР", "R.1.01.1"), ("ЩР", "R.1.01.2", "гр"), ("ЩР", "R.1.02.1"),
    ("ЩР", "L.1.01.1"), ("ЩР", "L.1.01.2"), ("ЩР", "L.1.02.1"),
    ("ЩР", "LD.1.01.1", None, None, None, "БП 24V 240W"), ("ЩР", "LD.1.01.2", "гр"),
    ("ЩР", "SH.1.01.1"), ("ЩР", "П.1.01.1"), ("ЩР", "ТП.1.01.1"), ("ЩР", None),
]


def fill(wb, node=3):
    ws = wb.Worksheets("Исходные данные")
    for i, (d, e, f, g, h, l, ba) in enumerate(SRC_LINES):
        r = 21 + i
        for col, v in zip("DEFGHI", (d, e, f, g, h, l)):
            ws.Range(f"{col}{r}").Value = v
        if ba:
            ws.Range(f"BA{r}").Value = ba
    rz = wb.Worksheets("Разбивка_по_щитам")
    hdr = [("" if x is None else str(x)) for x in rz.Range("A2:EZ2").Value[0]]
    col = lambda name: hdr.index(name) + 1          # столбцы — по заголовкам (порядок может меняться)
    heads = ("Щит", "Номер линии", "Группа", "Автомат", "Дифзащита, мА", "Аппараты", "Модуль", "Канал", "Клемма")
    for i, row in enumerate(RZ_ROWS):
        r = 3 + i
        row = tuple(row) + (None,) * (9 - len(row))
        for name, v in zip(heads, row):
            c = rz.Cells(r, col(name))
            if v is not None:
                c.Value = v
            elif name == "Номер линии":
                c.ClearContents()
    for r in range(3 + len(RZ_ROWS), 60):   # заготовки номеров шаблона 1.05 ниже демо-щита
        for name in ("Щит", "Начало линии", "Номер линии"):
            rz.Cells(r, col(name)).ClearContents()
    wb.Worksheets("Диапазоны").Range("H4").Value = node
    wb.Worksheets("Диапазоны").Range("I4").Value = 0.12
