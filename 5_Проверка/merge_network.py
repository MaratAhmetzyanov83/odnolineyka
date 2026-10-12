# -*- coding: utf-8 -*-
"""Один лист «Сеть_щита» вместо «Питание_щита», «Нижестоящие_щиты» и «ТКЗ» (книги v3.12).

    py merge_network.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит

Раздел А (строки 1–82) — бывший «Питание_щита», на месте: откуда питается щит, питающая линия, ΔU и Iкз на шинах.
Раздел Б (с строки 90) — бывший «Нижестоящие_щиты», перенесён вырезанием: Excel сам правит формулы и имена
  (НЩ_Руст, НЩ_линии…). Служебные столбцы сборки ГРЩ (P:BG) переносятся вместе.
Раздел В (с строки 130) — токи КЗ по линиям для печати (как отчёт ТКЗ): из «Исходных данных» BI–BM.
Лист «Питание_щита» переименован в «Сеть_щита»; «Нижестоящие_щиты» и «ТКЗ» удалены (пустые / отчёт).
Тексты «Питание_щита» / «Нижестоящие_щиты» в подписях заменены; ссылки «Пульта» исправлены.
Связи с ДРУГИМИ книгами объекта, которые ссылались на эти листы, придётся вставить заново
(кнопка надстройки «Связать с вышестоящим»).
Свой скрытый экземпляр Excel; закрывается только он. COM — только позиционные аргументы.
"""
import os
import signal
import sys

PWD = "sx"
POW, SUB, TKZ, NET, PULT = "Питание_щита", "Нижестоящие_щиты", "ТКЗ", "Сеть_щита", "Пульт"
SUB_ROW = 90                 # раздел Б: сюда встаёт A1 бывших «Нижестоящих»
SUB_ROWS = 38                # место под раздел Б (в v3.12 — 34 строки)
TKZ_ROW = 130                # раздел В
LINES = 998                  # строк линий, как в «Исходных данных» (3…1000)
HELP_COL = "BJ"              # служебный столбец раздела В (за P:BG сборки ГРЩ)
DONE_MARK = "СЕТЬ ЩИТА"
HINT_SUB = ("Связи ставит надстройка: в книге НИЖЕСТОЯЩЕГО щита — лента «Однолинейка» → «Связать с вышестоящим» "
            "(выбрать этот файл и линию-фидер). Строка заполнится сама: D:J — нагрузка из его файла, K:N — что он получит. "
            "«Обновить связи» — подтянуть свежие данные, «Дерево щитов» — все щиты папки объекта.")


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def sheet(wb, name):
    try:
        return wb.Worksheets(name)
    except Exception:
        return None


def rgb(t):
    return t[0] + t[1] * 256 + t[2] * 65536


def section(ws, row, text):
    c = ws.Range(f"A{row}")
    c.Value = text
    c.Font.Bold = True
    c.Font.Size = 12
    ws.Range(f"A{row}:S{row}").Interior.Color = rgb((221, 235, 247))


def move_sub(wb, net, log):
    sub = sheet(wb, SUB)
    if sub is None:
        return
    used = sub.UsedRange
    last_r = used.Row + used.Rows.Count - 1
    last_c = used.Column + used.Columns.Count - 1
    if last_r > SUB_ROWS:
        raise RuntimeError(f"«{SUB}»: {last_r} строк, места под раздел Б — {SUB_ROWS}")
    if wb.Application.WorksheetFunction.CountA(net.Range(f"A{SUB_ROW - 2}:BJ{SUB_ROW + SUB_ROWS}")) > 0:
        raise RuntimeError(f"«{POW}»: строки {SUB_ROW - 2}–{SUB_ROW + SUB_ROWS} заняты")
    src = sub.Range(sub.Cells(1, 1), sub.Cells(last_r, last_c))
    addr = src.Address.replace("$", "")
    src.Cut(net.Range(f"A{SUB_ROW}"))
    sub.Visible = -1
    sub.Delete()
    section(net, SUB_ROW - 2, "РАЗДЕЛ Б. НИЖЕСТОЯЩИЕ ЩИТЫ — кого питает этот щит")
    net.Range(f"A{SUB_ROW + 1}").Value = HINT_SUB
    log.append(f"раздел Б: «{SUB}» {addr} -> строка {SUB_ROW}, лист удалён")


def build_tkz(wb, net, log):
    if net.Range(f"A{TKZ_ROW}").Value:
        return
    t = sheet(wb, TKZ)
    if t is not None:
        t.Visible = -1
        t.Delete()
    section(net, TKZ_ROW, "РАЗДЕЛ В. ТОКИ КЗ ПО ЛИНИЯМ — для печати, ввода нет "
                          "(итог на шинах — раздел А, п. 3; линии — «Исходные данные» BI–BM)")
    h = TKZ_ROW + 1
    r0, r1 = h + 1, h + LINES
    hdr = {"A": "№", "B": "Линия", "C": "Щит", "D": "L, м", "E": "In, А", "G": "Наименование", "H": "Кабель",
           "I": "Zп в конце, Ом", "J": "Iкз min, А", "K": "Iкз min / In", "L": "Iкз max, кА", "M": "Отключение ≤ 0,4 с"}
    for c, v in hdr.items():
        net.Range(f"{c}{h}").Value = v
    hr = net.Range(f"A{h}:M{h}")
    hr.Font.Bold = True
    hr.WrapText = True
    hr.Interior.Color = rgb((242, 242, 242))
    S = "'Исходные данные'!"
    k = HELP_COL
    net.Range(f"{k}{h}").Value = "строка ИД"
    # номер строки «Исходных данных» со следующей линией — поиск от предыдущей, без летучих функций
    net.Range(f"{k}{r0}").Formula = f'=IFERROR(MATCH(TRUE,INDEX(LEN({S}$D$3:$D$1000)>0,0),0)+2,"")'
    net.Range(f"{k}{r0 + 1}:{k}{r1}").Formula = (
        f'=IF({k}{r0}="","",IFERROR({k}{r0}+MATCH(TRUE,INDEX(LEN(INDEX({S}$D$1:$D$1000,{k}{r0}+1):{S}$D$1000)>0,0),0),""))')

    def pick(col, rnd=None):
        v = f"INDEX({S}${col}$1:${col}$1000,${k}{r0})"
        if rnd is not None:
            v = f'IF(ISNUMBER({v}),ROUND({v},{rnd}),{v}&"")'
        return f'=IF(${k}{r0}="","",{v})'
    cols = {"A": f'=IF(${k}{r0}="","",ROW()-{r0 - 1})', "B": pick("D"), "C": pick("BZ"), "D": pick("M"),
            "E": pick("U"), "G": pick("AJ"), "H": pick("AE"), "I": pick("BL", 4), "J": pick("BI", 0),
            "K": pick("BJ", 2), "L": pick("BM", 2), "M": pick("BK")}
    for c, f in cols.items():
        net.Range(f"{c}{r0}:{c}{r1}").Formula = f
    net.Range(f"A{h}:M{r1}").Borders.LineStyle = 1
    net.Range(f"A{h}:M{r1}").Borders.Color = rgb((191, 191, 191))
    net.Columns(k).Hidden = True
    if net.Columns("A").ColumnWidth < 4:
        net.Columns("A").ColumnWidth = 4         # № до 998
    if net.Columns("H").ColumnWidth < 20:
        net.Columns("H").ColumnWidth = 20        # кабель
    log.append(f"раздел В: токи КЗ, строки {r0}–{r1}" + (f"; лист «{TKZ}» удалён" if t is not None else ""))


def rename_texts(wb, log):
    n = 0
    for ws in wb.Worksheets:
        for what, repl in (("«Питание_щита»", "«Сеть_щита»"),
                           ("«Нижестоящие_щиты»", "«Сеть_щита», раздел Б"),
                           ("листа «ТКЗ»", "раздела В «Сети_щита»")):
            try:
                if ws.Cells.Find(what, ws.Cells(1, 1), -4123, 2) is not None:
                    ws.Cells.Replace(what, repl, 2)
                    n += 1
            except Exception:
                pass
    if n:
        log.append(f"подписи с именами листов поправлены: {n}")


def fix_pult(wb, log):
    ws = sheet(wb, PULT)
    if ws is None:
        return
    gone = {POW, SUB, TKZ}
    first = None
    for r in range(1, 60):
        v = ws.Cells(r, 2).Value
        if v in gone:
            if first is None:
                first = r
                ws.Cells(r, 2).Hyperlinks.Delete()
                ws.Cells(r, 2).Value = None
                ws.Hyperlinks.Add(ws.Cells(r, 2), "", f"'{NET}'!A1", NET, NET)
                ws.Cells(r, 3).Value = "питание щита (сверху), нижестоящие щиты, токи КЗ по линиям"
            else:
                ws.Cells(r, 2).Hyperlinks.Delete()
                ws.Range(f"B{r}:C{r}").Value = None
    # пустые строки в списке — сдвинуть вверх (список — последний блок «Пульта»)
    if first is not None:
        rows = [(ws.Cells(r, 2).Value, ws.Cells(r, 3).Value, ws.Cells(r, 2).Hyperlinks.Count > 0)
                for r in range(first, 60)]
        items = [x for x in rows if x[0]]
        for r in range(first, 60):
            ws.Cells(r, 2).Hyperlinks.Delete()
            ws.Range(f"B{r}:C{r}").Value = None
        for i, (name, text, link) in enumerate(items):
            r = first + i
            if link:
                ws.Hyperlinks.Add(ws.Cells(r, 2), "", f"'{name}'!A1", name, name)
            else:
                ws.Cells(r, 2).Value = name
            ws.Cells(r, 3).Value = text
        log.append("«Пульт»: ссылка «Сеть_щита» вместо трёх листов")


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        app.Calculation = -4135
        net = sheet(wb, NET) or sheet(wb, POW)
        if net is None:
            raise RuntimeError(f"нет листа «{POW}»")
        if net.Name == POW:
            if net.ProtectContents:
                net.Unprotect(PWD)
            move_sub(wb, net, log)
            build_tkz(wb, net, log)
            title = net.Range("B1")
            if isinstance(title.Value, str) and DONE_MARK not in title.Value:
                title.Value = (f"{DONE_MARK}. РАЗДЕЛ А — ПИТАНИЕ: " + title.Value.split(":", 1)[-1].strip() +
                               f"  ·  раздел Б (нижестоящие) — строка {SUB_ROW - 2}, раздел В (токи КЗ) — строка {TKZ_ROW}")
            net.Range("B4").Value = "1. ОТ ВЫШЕСТОЯЩЕГО ЩИТА — ставит надстройка: «Связать с вышестоящим» (выбрать его файл и фидер)"
            net.Name = NET
            log.append(f"«{POW}» -> «{NET}»")
            rename_texts(wb, log)
            fix_pult(wb, log)
        else:
            log.append("уже объединено — пропуск")
        app.Calculation = -4105
        app.CalculateFull()
        wb.Save()
        wb.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        try:
            os.kill(pid, signal.SIGTERM)
        except Exception:
            pass
    print(os.path.basename(path))
    for s in log:
        print("  " + s)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for p in sys.argv[1:]:
        run(p)
