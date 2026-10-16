# -*- coding: utf-8 -*-
"""Новые УГО потребителей (SXDRAW v3.24): кондиционеры, вентилятор, насос, приточная установка, щит.

    py patch_ugo.py <книга.xlsx> [...]      — книга закрыта; повторный запуск ничего не портит

«Справочная»:
  * BA:BD («УГО_блоки») — новые значки: «Кондиционер (внутр. блок)», «Кондиционер (наруж. блок, ККБ)»,
    «Вентилятор», «Насос», «Приточная установка», «Щит». SXDRAW ставит ваш динамический блок «Дом.1» с нужным
    состоянием видимости (голубой, как в ЩР-2); нет «Дом.1» — свой значок программы (SX_UGO_*). Щит — значок программы.
  * AU:AV («УГО_спр») — коды на новые значки, только если там значение по умолчанию (свой выбор не трогаем):
    K, К, КВ — внутр. блок; ККБ, НБ — наруж. блок; В — вентилятор; НС — насос; ПВ, ПУ — приточная установка
    (было «Привод» или пусто); Щ — щит (было «Прочее»).
    Кода ККБ нет — добавляется в «Доп. коды».
«Однолинейка» CF (узел КМ): двигатели — все коды «MOTOR*» (новые значки кондиционеров, вентилятора и насоса тоже).
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import signal
import sys

PWD = "sx"
SPR = "Справочная"
ODN = "Однолинейка"
PROG = "(условный значок программы)"
UGO = [   # название (AV), блок (BB, для справки), код программы (BC)
    ("Кондиционер (внутр. блок)", "Дом.1 «Внутренний блок»", "MOTOR_AC"),
    ("Кондиционер (наруж. блок, ККБ)", "Дом.1 «Внешний блок»", "MOTOR_ACO"),
    ("Вентилятор", "Дом.1 «Вентилятор»", "MOTOR_FAN"),
    ("Насос", "Дом.1 «Насос»", "MOTOR_PUMP"),
    ("Щит", PROG, "PANEL"),
    ("Приточная установка", "Дом.1 «Приточка»", "MOTOR_AHU"),
]
MOTOR, OTHER = "Привод (двигатель)", "Прочее (прямоуг. с подписью)"
# код линии -> (новое УГО, какие старые значения заменяем)
CODES = {
    "K": (UGO[0][0], (MOTOR, None)), "К": (UGO[0][0], (MOTOR, None)), "КВ": (UGO[0][0], (MOTOR, None)),
    "ККБ": (UGO[1][0], (MOTOR, None)), "НБ": (UGO[1][0], (MOTOR, None)),
    "В": (UGO[2][0], (MOTOR,)), "НС": (UGO[3][0], (MOTOR,)), "Щ": (UGO[4][0], (OTHER,)),
    "ПВ": (UGO[5][0], (MOTOR, None)), "ПУ": (UGO[5][0], (MOTOR, None)),
}
ADD = ["ККБ"]                                    # добавить в «Доп. коды», если кода нет
# узел КМ (patch_km.py): двигатели — BK «MOTOR*»
CF_OLD = ('"MOTOR"', '"<>MOTOR"')
CF_NEW = ('"MOTOR*"', '"<>MOTOR*"')


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def unprotect(ws):
    if not ws.ProtectContents:
        return None
    p = ws.Protection
    opts = (True, True, True, False, p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
            p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks, p.AllowDeletingColumns,
            p.AllowDeletingRows, p.AllowSorting, p.AllowFiltering, p.AllowUsingPivotTables)
    ws.Unprotect(PWD)
    return opts


def spr(wb, log):
    ws = wb.Worksheets(SPR)
    if str(ws.Range("BA3").Value or "").strip()[:3] != "УГО":
        raise SystemExit(f"«{SPR}» BA3 — не таблица значков УГО — книга не изменена")
    opts = unprotect(ws)
    # значки
    rows = [(r, str(ws.Range(f"BA{r}").Value or "").strip()) for r in range(4, 31)]
    have = {v for _, v in rows}
    free = [r for r, v in rows if not v]
    added = []
    for name, blk, code in UGO:
        if name in have:
            r = next(r for r, v in rows if v == name)
            if ws.Range(f"BB{r}").Value != blk and ws.Range(f"BC{r}").Value == code:
                ws.Range(f"BB{r}").Value = blk
                added.append(f"{name} (BB)")
            continue
        if not free:
            log.append(f"«{SPR}» BA4:BA30 — нет места для «{name}»")
            continue
        r = free.pop(0)
        ws.Range(f"BA{r}:BC{r}").Value = ((name, blk, code),)
        if not ws.Range(f"BD{r}").Value:
            ws.Range(f"BD{r}").Value = "нет"
        added.append(name)
    if added:
        log.append(f"«{SPR}» BA — новые УГО: {', '.join(added)}")
    # коды -> значки
    last = 72
    codes = {}
    for r in range(4, last + 1):
        c = str(ws.Range(f"AU{r}").Value or "").strip()
        if c:
            codes.setdefault(c, []).append(r)
    for c in ADD:
        if c in codes:
            continue
        hdr = next((r for r in range(4, last + 1) if str(ws.Range(f"AU{r}").Value or "").startswith("Доп. коды")), None)
        r = next((r for r in range((hdr or 53) + 1, last + 1) if not ws.Range(f"AU{r}").Value), None)
        if r is None:
            log.append(f"«{SPR}» AU — нет места для кода {c}")
            continue
        ws.Range(f"AU{r}").Value = c
        ws.Range(f"AW{r}").Value = "нет"
        codes[c] = [r]
        log.append(f"«{SPR}» AU{r} — новый код {c}")
    ch = []
    for c, (ugo, old) in CODES.items():
        for r in codes.get(c, []):
            v = ws.Range(f"AV{r}").Value
            v = str(v).strip() if v not in (None, "") else None
            if v != ugo and v in old:
                ws.Range(f"AV{r}").Value = ugo
                if not ws.Range(f"AW{r}").Value:
                    ws.Range(f"AW{r}").Value = "нет"
                ch.append(f"{c} → {ugo}")
    if ch:
        log.append(f"«{SPR}» AV — {'; '.join(ch)}")
    if opts:
        ws.Protect(PWD, *opts)


def odn(wb, log):
    ws = wb.Worksheets(ODN)
    if str(ws.Range("CF2").Value or "").find("двигател") < 0:
        return                                        # узла КМ нет (patch_km.py не применён)
    opts = unprotect(ws)
    last = max(ws.Cells(ws.Rows.Count, c).End(-4162).Row for c in ("A", "H", "AF", "BH"))
    f = ws.Range("CF3").Formula2
    if CF_OLD[0] in f:
        g = f
        for a, b in zip(CF_OLD, CF_NEW):
            g = g.replace(a, b)
        ws.Range(f"CF3:CF{last}").Formula2 = g
        log.append(f"«{ODN}» CF — двигатели узла КМ: все коды MOTOR* (кондиционеры, вентиляторы, насосы)")
    if opts:
        ws.Protect(PWD, *opts)


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        if wb.ReadOnly:                              # открыта у другого — не трогаем
            wb.Close(False)
            raise SystemExit(f"{os.path.basename(path)}: открыта у другого пользователя — пропущена")
        app.Calculation = -4135
        spr(wb, log)
        odn(wb, log)
        app.Calculation = -4105
        app.CalculateFull()
        if log:
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
    for s in log or ["изменений нет"]:
        print("  " + s)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for p in sys.argv[1:]:
        run(p)
