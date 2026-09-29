# -*- coding: utf-8 -*-
"""Загрузка значений книги как «сетки» листов после полного пересчёта.

Бэкенды (по порядку предпочтения):
  excel        - Excel COM, CalculateFull, ссылки не обновляются, всегда Quit;
  libreoffice  - soffice --headless --convert-to xlsx (пересчёт при загрузке), если установлен;
  cache        - openpyxl data_only (кэш, сохранённый Excel; БЕЗ пересчёта) - с предупреждением.
"""
import os
import shutil
import subprocess
import tempfile
import time
import warnings
import glob

ERR_CODES = {
    -2146826281: "#ДЕЛ/0!", -2146826246: "#Н/Д", -2146826259: "#ИМЯ?",
    -2146826288: "#ПУСТО!", -2146826252: "#ЧИСЛО!", -2146826265: "#ССЫЛКА!",
    -2146826273: "#ЗНАЧ!", -2146826226: "#ДАННЫЕ!", -2146826233: "#ВЫЧИСЛ!",
}
ERR_TEXT = {  # то, что отдаёт openpyxl/LibreOffice -> единый вид
    "#N/A": "#Н/Д", "#REF!": "#ССЫЛКА!", "#VALUE!": "#ЗНАЧ!", "#NAME?": "#ИМЯ?",
    "#DIV/0!": "#ДЕЛ/0!", "#NUM!": "#ЧИСЛО!", "#NULL!": "#ПУСТО!",
    "#Н/Д": "#Н/Д", "#ССЫЛКА!": "#ССЫЛКА!", "#ЗНАЧ!": "#ЗНАЧ!", "#ИМЯ?": "#ИМЯ?",
    "#ДЕЛ/0!": "#ДЕЛ/0!", "#ЧИСЛО!": "#ЧИСЛО!", "#ПУСТО!": "#ПУСТО!",
    "#GETTING_DATA": "#ДАННЫЕ!", "#SPILL!": "#ВЫЧИСЛ!", "#CALC!": "#ВЫЧИСЛ!",
}


class Grid:
    """Значения одного листа: data[i][j], левый верхний угол = (r0, c0), нумерация с 1."""

    def __init__(self, name, r0, c0, data):
        self.name, self.r0, self.c0, self.data = name, r0, c0, data
        self.nr = len(data)
        self.nc = len(data[0]) if data else 0

    def get(self, r, c):
        i, j = r - self.r0, c - self.c0
        if 0 <= i < self.nr and 0 <= j < self.nc:
            return self.data[i][j]
        return None

    def rows(self):
        return range(self.r0, self.r0 + self.nr)


def _norm(v):
    if v is None:
        return None
    if isinstance(v, str):
        if v == "":
            return None
        return ERR_TEXT.get(v, v)
    if isinstance(v, bool):
        return v
    if isinstance(v, int) and v in ERR_CODES:
        return ERR_CODES[v]
    if isinstance(v, (int, float)):
        return float(v)
    return v


def _excel_pids():
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                             capture_output=True).stdout.decode("latin-1")
    except Exception:
        return set()
    pids = set()
    for ln in out.splitlines():
        p = [x.strip('"') for x in ln.split('","')]
        if len(p) > 1 and p[0].strip('"').upper() == "EXCEL.EXE":
            try:
                pids.add(int(p[1]))
            except ValueError:
                pass
    return pids


def load_excel(path, tmpdir, log=print):
    """Пересчитать КОПИЮ книги в отдельном экземпляре Excel и прочитать значения."""
    import pythoncom
    import win32com.client
    import win32process
    os.makedirs(tmpdir, exist_ok=True)
    copy = os.path.join(tmpdir, "regress_copy.xlsx")
    shutil.copyfile(path, copy)
    before = _excel_pids()
    pythoncom.CoInitialize()
    xl = wb = None
    pid = None
    grids = {}
    try:
        xl = win32com.client.DispatchEx("Excel.Application")   # свой экземпляр, не чужой Excel
        try:
            pid = win32process.GetWindowThreadProcessId(xl.Hwnd)[1]
        except Exception:
            pid = None
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AskToUpdateLinks = False
        xl.EnableEvents = False
        try:
            xl.AutomationSecurity = 3      # без макросов
        except Exception:
            pass
        wb = xl.Workbooks.Open(os.path.abspath(copy), 0, True)   # UpdateLinks=0, ReadOnly
        xl.CalculateFull()
        for ws in wb.Worksheets:
            ur = ws.UsedRange
            r0, c0 = ur.Row, ur.Column
            v = ur.Value2
            if v is None:
                grids[ws.Name] = Grid(ws.Name, r0, c0, [])
                continue
            if not isinstance(v, tuple):          # одна ячейка
                v = ((v,),)
            grids[ws.Name] = Grid(ws.Name, r0, c0, [[_norm(x) for x in row] for row in v])
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        xl = wb = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        if pid:                                    # ждём выхода НАШЕГО процесса (до 6 с), иначе убиваем его
            for _ in range(12):
                if pid not in _excel_pids():
                    break
                time.sleep(0.5)
            if pid in _excel_pids():
                subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
                time.sleep(0.5)
                if pid in _excel_pids():
                    log("ПРЕДУПРЕЖДЕНИЕ: не удалось закрыть EXCEL.EXE (PID %s)" % pid)
        else:
            time.sleep(1.5)
            left = _excel_pids() - before
            if left:
                log("ПРЕДУПРЕЖДЕНИЕ: PID Excel не определён; появились процессы EXCEL.EXE: %s" % sorted(left))
        try:
            os.remove(copy)
        except Exception:
            pass
    return grids


def load_cache(path):
    """Кэш openpyxl (значения, сохранённые Excel при последнем сохранении). БЕЗ пересчёта."""
    import openpyxl
    warnings.simplefilter("ignore")
    wb = openpyxl.load_workbook(path, data_only=True)
    grids = {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            grids[ws.title] = Grid(ws.title, 1, 1, [])
            continue
        w = max(len(r) for r in rows)
        data = [[_norm(x) for x in list(r) + [None] * (w - len(r))] for r in rows]
        grids[ws.title] = Grid(ws.title, 1, 1, data)
    return grids


def load_libreoffice(path, tmpdir, log=print):
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        for p in glob.glob(r"C:\Program Files*\LibreOffice*\program\soffice.exe"):
            exe = p
    if not exe:
        raise RuntimeError("LibreOffice не найден")
    os.makedirs(tmpdir, exist_ok=True)
    src = os.path.join(tmpdir, "lo_in.xlsx")
    shutil.copyfile(path, src)
    outdir = os.path.join(tmpdir, "lo_out")
    os.makedirs(outdir, exist_ok=True)
    subprocess.run([exe, "--headless", "--convert-to", "xlsx", "--outdir", outdir, src],
                   capture_output=True, timeout=600)
    return load_cache(os.path.join(outdir, "lo_in.xlsx"))


def load(path, backend="auto", tmpdir=None, log=print):
    """Возвращает (grids, backend_used, warnings)."""
    tmpdir = tmpdir or tempfile.mkdtemp(prefix="sx_regress_")
    warns = []
    order = ["excel", "libreoffice", "cache"] if backend == "auto" else [backend]
    for b in order:
        try:
            if b == "excel":
                return load_excel(path, tmpdir, log), "excel", warns
            if b == "libreoffice":
                g = load_libreoffice(path, tmpdir, log)
                warns.append("Пересчёт выполнен LibreOffice (возможны мелкие расхождения с Excel).")
                return g, "libreoffice", warns
            if b == "cache":
                warns.append("ВНИМАНИЕ: значения взяты из КЭША файла без пересчёта; сравнение достоверно, "
                             "только если книга сохранена Excel после последней правки.")
                return load_cache(path), "cache", warns
        except Exception as e:
            log("бэкенд %s недоступен: %s" % (b, e))
            warns.append("бэкенд %s недоступен (%s)" % (b, e))
    raise RuntimeError("Не удалось прочитать книгу ни одним способом")
