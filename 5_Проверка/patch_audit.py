# -*- coding: utf-8 -*-
"""Исправления по аудиту расчётов 11.10.2026 (ошибки 1–4) для книг v3.12.

    py patch_audit.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит

1. ΔU: подбор и проверка по одной норме.
   «Исходные данные» V — сечение по потерям подбирается по остатку: ΔUдоп (BF) − ΔU до шин щита,
   но не больше ΔUдоп линии (4,5 %, «Справочная» BQ32). Тогда подобранное сечение проходит проверку BG ≤ BF.
   Линии ≤ 50 В: ΔU от своего источника (БП), допуск 5 % (BQ33) — в BF; к BG не прибавляется ΔU до шин.
2. Ко = 0,9 только если есть бытовая нагрузка (СП 256: Рр = Рр.быт + 0,9·Рр.сил — для жилого дома).
   «Нагрузка_щитов» AF:AH 10/19/28, «Режимы» (по источникам и категориям), «Расчёт_нагрузок» U11:U40 и H4.
3. Код кабеля «мфр» был дважды («Справочная» L5 ВВГнг(А)-FRLS и L11 ППГнг(А)-FRHF) — второй -> «мфрх».
4. КЗ без данных питания: «Исходные данные» BK = «не проверено: заполните «Питание_щита»» (было пусто);
   на «Пульте» строка «КЗ» — сколько линий не проверено / мало Iкз.
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import re
import signal
import sys

PWD = "sx"
SRC, NAG, MOD, PZ, REF, PULT = "Исходные данные", "Нагрузка_щитов", "Режимы", "Расчёт_нагрузок", "Справочная", "Пульт"
LV = "AND(ISNUMBER(RC6),RC6>0,RC6<=50)"                       # линия ≤ 50 В (F — напряжение)
BUS_DU = ("IFERROR(INDEX(ПЩ_dU,MATCH(INDEX(Разбивка_по_щитам!R1C6:R1000C6,MATCH(RC4,Разбивка_по_щитам!R1C8:R1000C8,0)),"
          "ПЩ_щиты,0))*1,0)")                                   # ΔU до шин щита этой линии
NOT_CHECKED = "не проверено: заполните «Питание_щита»"
KO_LINE = re.compile(r"(\$?[A-Z]{1,3}\$?\d+)\*(\$?[A-Z]{1,3}\$?\d+)\+Ко_сил\*")
KO_PZ = "IF(SUM($M$11:$M$40)>0,Ко_сил,1)"


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


def unprotect(ws):
    if not ws.ProtectContents:
        return None
    p = ws.Protection
    opts = (True, True, True, False,
            p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
            p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks,
            p.AllowDeletingColumns, p.AllowDeletingRows, p.AllowSorting,
            p.AllowFiltering, p.AllowUsingPivotTables)
    ws.Unprotect(PWD)
    return opts


def runs(ws, col, r0=3, r1=1000):
    """Подряд идущие строки с одинаковой формулой R1C1: [(первая, последняя, формула)]."""
    fs = [x[0] for x in ws.Range(f"{col}{r0}:{col}{r1}").Formula2R1C1]
    out, start = [], r0
    for i in range(1, len(fs) + 1):
        if i == len(fs) or fs[i] != fs[i - 1]:
            out.append((start, r0 + i - 1, fs[i - 1]))
            start = r0 + i
    return out


def call_end(f, start):
    """индекс за закрывающей скобкой вызова, открывающая скобка — первая '(' начиная со start"""
    i = f.index("(", start)
    depth, q = 0, False
    for j in range(i, len(f)):
        ch = f[j]
        if ch == '"':
            q = not q
        elif not q:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    return j + 1
    raise ValueError("нет закрывающей скобки")


# ---- правила для столбцов «Исходных данных» (R1C1); None — не трогать
def rule_v(f):
    old = "lim_,IF(lv_,dU_НН_доп,dU_лин_доп)"
    if old not in f or "rem_" in f:
        return None
    new = ("lim_,IF(lv_,dU_НН_доп,LET(rem_,N(RC58)-" + BUS_DU +
           ",IF(rem_<=0,dU_лин_доп,MIN(dU_лин_доп,rem_))))")
    return f.replace(old, new, 1)


def rule_bf(f):
    key = ",IF(IFERROR(VLOOKUP("
    if key not in f or "dU_НН_доп" in f:
        return None
    return f.replace(key, f",IF({LV},dU_НН_доп,IF(IFERROR(VLOOKUP(", 1) + ")"


def rule_bg(f):
    key = "+IFERROR(INDEX(ПЩ_dU,"
    i = f.find(key)
    if i < 0 or LV in f:
        return None
    j = call_end(f, i + 1)                    # конец IFERROR(...)
    return f[:i] + f"+IF({LV},0," + f[i + 1:j] + ")" + f[j:]


def rule_bk(f):
    key = '=IF(NOT(ISNUMBER(RC[-1])),"",'
    if not f.startswith(key) or "не проверено" in f:
        return None
    return ('=IF(OR(RC4="",RC80=1,NOT(ISNUMBER(RC22))),"",IF(NOT(ISNUMBER(RC[-1])),"' + NOT_CHECKED + '",' +
            f[len(key):] + ")")


def patch_src(wb, log):
    ws = wb.Worksheets(SRC)
    opts = unprotect(ws)
    for col, rule in (("V", rule_v), ("BF", rule_bf), ("BG", rule_bg), ("BK", rule_bk)):
        n = 0
        for a, b, f in runs(ws, col):
            nf = rule(f) if isinstance(f, str) and f.startswith("=") else None
            if nf:
                ws.Range(f"{col}{a}:{col}{b}").Formula2R1C1 = nf
                n += b - a + 1
        if n:
            log.append(f"«{SRC}» {col}: {n} строк")
    h = ws.Range("BF2")
    if isinstance(h.Value, str) and "50 В" not in h.Value:
        h.Value = h.Value + "; ≤ 50 В — 5 % от своего источника"
    if opts:
        ws.Protect(PWD, *opts)


def patch_ko(wb, log):
    n = 0
    for name in (NAG, MOD):
        ws = sheet(wb, name)
        if ws is None:
            continue
        opts = unprotect(ws)
        rng = ws.UsedRange
        r0, c0 = rng.Row, rng.Column
        for i, row in enumerate(rng.Formula):        # одно чтение, запись — только изменённых ячеек
            for j, f in enumerate(row):
                if isinstance(f, str) and "+Ко_сил*" in f:
                    g, k = KO_LINE.subn(lambda m: f"{m.group(1)}*{m.group(2)}+IF(N({m.group(1)})>0,Ко_сил,1)*", f)
                    if k:
                        ws.Cells(r0 + i, c0 + j).Formula = g
                        n += 1
        if name == NAG:
            for a in ("AE10", "AE19", "AE28"):
                v = ws.Range(a).Value
                if isinstance(v, str) and "если есть быт" not in v:
                    ws.Range(a).Value = v + " (Ко — если есть быт)"
        if opts:
            ws.Protect(PWD, *opts)
    ws = sheet(wb, PZ)
    if ws is not None:
        opts = unprotect(ws)
        for r in range(11, 41):
            c = ws.Range(f"U{r}")
            f = c.Formula
            if isinstance(f, str) and "+Ко_сил*" in f:
                c.Formula = f.replace("+Ко_сил*", f"+{KO_PZ}*", 1)
                n += 1
        if ws.Range("H4").Formula == "=Ко_сил":
            ws.Range("H4").Formula = "=" + KO_PZ
            n += 1
        if opts:
            ws.Protect(PWD, *opts)
    ref = wb.Worksheets(REF)
    note = ref.Range("BR14")
    if isinstance(note.Value, str) and "только если" not in note.Value:
        note.Value = note.Value + ". Ко применяется только если в щите есть бытовая нагрузка; без быта Ко = 1."
    if n:
        log.append(f"Ко только при бытовой нагрузке: {n} формул")


def patch_codes(wb, log):
    ws = wb.Worksheets(REF)
    seen = {}
    for r in range(4, 52):
        code = ws.Range(f"L{r}").Value
        if not code:
            continue
        if code in seen and code == "мфр" and "FRHF" in str(ws.Range(f"M{r}").Value):
            ws.Range(f"L{r}").Value = "мфрх"
            log.append(f"«{REF}» L{r}: код «мфр» (ППГнг(А)-FRHF) -> «мфрх»")
        elif code in seen:
            log.append(f"ВНИМАНИЕ: код кабеля «{code}» дважды — «{REF}» L{seen[code]} и L{r}")
        seen.setdefault(code, r)
    note = ws.Range("BR32")
    if isinstance(note.Value, str) and "остат" not in note.Value:
        note.Value = ("ΔUдоп линии: сечение растёт, пока ΔU линии ≤ ΔUдоп и ΔU до шин + ΔU линии ≤ допуска от ВРУ "
                      "(«Исходные данные» BF: свет 3 %, прочие 4 %) — подбор по остатку (ГОСТ Р 50571.5.52, прил. G)")


PULT_KZ = ('=IF(COUNTIF(\'Исходные данные\'!$BK$3:$BK$1000,"не проверено*")>0,'
           '"не проверено на "&COUNTIF(\'Исходные данные\'!$BK$3:$BK$1000,"не проверено*")&" линиях — заполните «Питание_щита»",'
           'IF(COUNTIF(\'Исходные данные\'!$BK$3:$BK$1000,"мало*")>0,'
           '"мало Iкз на "&COUNTIF(\'Исходные данные\'!$BK$3:$BK$1000,"мало*")&" линиях — см. «Исходные данные» BK",'
           'IF(COUNTIF(\'Исходные данные\'!$BK$3:$BK$1000,"ok*")>0,"проверено","")))')


def patch_pult(wb, log):
    ws = sheet(wb, PULT)
    if ws is None:
        return
    # строка «Ошибки» — в B; под ней свободная строка: туда «КЗ»
    if any(ws.Cells(r, 2).Value == "КЗ" for r in range(4, 20)):
        return
    for r in range(4, 20):
        if ws.Cells(r, 2).Value == "Ошибки":
            if ws.Cells(r + 1, 2).Value not in (None, ""):
                log.append("«Пульт»: под «Ошибки» занято — строка «КЗ» не добавлена")
                return
            ws.Cells(r + 1, 2).Value = "КЗ"
            ws.Cells(r + 1, 3).Formula = PULT_KZ
            ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 3)).Font.Bold = True
            log.append("«Пульт»: строка «КЗ»")
            return


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        app.Calculation = -4135
        patch_src(wb, log)
        patch_ko(wb, log)
        patch_codes(wb, log)
        patch_pult(wb, log)
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
    for s in log or ["изменений нет"]:
        print("  " + s)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for p in sys.argv[1:]:
        run(p)
