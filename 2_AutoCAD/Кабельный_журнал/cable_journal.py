"""Excel cable journal -> native DWG tables in the supplied A3 template.

Windows, desktop Excel and AutoCAD. Excel and template inputs are read-only.
Only explicitly selected generated output files are replaced, with a backup.
"""
from __future__ import annotations

import argparse
from collections import Counter
from datetime import datetime
import gc
import json
import math
import os
from pathlib import Path
import queue
import re
import shutil
import sys
import tempfile
import threading
import time
import traceback
import uuid

import pythoncom
import pywintypes
import win32com.client as com

VERSION = '1.4'
APP_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
DEFAULT_TEMPLATE = APP_DIR / 'КЖ_шаблон.dwg'
ERROR_CODES = {-2146826281, -2146826246, -2146826259, -2146826288,
               -2146826252, -2146826265, -2146826273, -2146826226, -2146826233}
ERROR_TEXT = {'#N/A', '#REF!', '#VALUE!', '#NAME?', '#DIV/0!', '#NUM!', '#NULL!',
              '#Н/Д', '#ССЫЛКА!', '#ЗНАЧ!', '#ИМЯ?', '#ДЕЛ/0!', '#ЧИСЛО!', '#ПУСТО!'}
MANIFEST_SUFFIX = '.kj.json'
ERROR_DISPLAY = {-2146826281:'#ДЕЛ/0!',-2146826246:'#Н/Д',-2146826259:'#ИМЯ?',
                 -2146826288:'#ПУСТО!',-2146826252:'#ЧИСЛО!',-2146826265:'#ССЫЛКА!',
                 -2146826273:'#ЗНАЧ!',-2146826226:'#ДАННЫЕ!',-2146826233:'#ВЫЧИСЛ!'}


class RetryCOM:
    """Transparent COM proxy: AutoCAD can reject any call while regenerating."""
    def __init__(self,native):
        object.__setattr__(self,'_native',native)

    def __getattr__(self,name):
        value=retry(lambda:getattr(self._native,name))
        if isinstance(value,com.dynamic.CDispatch):
            return wrapped(value)
        if callable(value):
            def invoke(*args,**kwargs):
                return wrapped(retry(lambda:value(*(unwrapped(a) for a in args),
                                                   **{k:unwrapped(v) for k,v in kwargs.items()})))
            return invoke
        return wrapped(value)

    def __setattr__(self,name,value):
        retry(lambda:setattr(self._native,name,unwrapped(value)))

    def __iter__(self):
        return iter(items(self))


def unwrapped(v):
    if isinstance(v,RetryCOM):return v._native
    if isinstance(v,tuple):return tuple(unwrapped(x) for x in v)
    return v


def wrapped(v):
    if isinstance(v,com.dynamic.CDispatch):return RetryCOM(v)
    if isinstance(v,tuple):return tuple(wrapped(x) for x in v)
    return v


def retry(fn, timeout=25):
    """Wait only for COM busy/not-yet-ready responses; never hide data failures."""
    deadline = time.monotonic() + timeout
    while True:
        try:
            return fn()
        except (pywintypes.com_error, AttributeError, StopIteration, TypeError) as exc:
            busy = isinstance(exc,(AttributeError,StopIteration)) or (isinstance(exc,TypeError) and 'enumeration' in str(exc)) or (isinstance(exc,pywintypes.com_error) and
                    (exc.hresult in (-2147418111,-2147417846) or
                     (exc.excepinfo and len(exc.excepinfo)>5 and exc.excepinfo[5]==-2145386296)))
            if not busy or time.monotonic() >= deadline:
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(.15)


def is_error(v):
    return (isinstance(v, int) and v in ERROR_CODES) or (isinstance(v, str) and v in ERROR_TEXT)


def is_summary_row(row):
    """User-inserted totals have a label in P and no route in Q/R."""
    label = ' '.join(text(row[0]).casefold().split())
    return bool(re.match(r'^(?:всего|итого)(?:\s|:|$)',label)) and all(v in (None,'',0) or is_error(v) for v in row[1:3])


def summary_header(row):
    labels = [' '.join(text(v).casefold().split()) if not is_error(v) else '' for v in row]
    return any(re.match(r'^(?:всего|итого)(?:\s|:|$)',v) for v in labels) or (
        labels[0]=='марка' and any(v.startswith('длина') for v in labels[3:]))


def lookup_formula(value):
    return bool(re.search(r'\b(?:VLOOKUP|HLOOKUP|XLOOKUP|ВПР|ГПР)\s*\(',text(value),re.I))


def parse_rows(values, formulas, filename, panel='', skip_first=False):
    """Separate source cable records, handwritten totals and unused formula tail."""
    rows, warnings, cable_markings = [], [], []
    summaries = 0
    summary_block = False
    excluded_first = None
    for row_number, (row, formula) in enumerate(zip(values, formulas), 5):
        missing_route = all(v in (None,'',0) or is_error(v) for v in row[1:3])
        header = summary_header(row) and missing_route
        if header:
            summary_block = True
        elif not missing_route:
            summary_block = False
        meaningful = any(v not in (None,'',0) and not is_error(v) for v in row)
        cable_lookup = row[0] not in (None,'',0) and not is_error(row[0]) and lookup_formula(formula[0]) and lookup_formula(formula[3])
        if summary_block and meaningful and (header or missing_route) and not cable_lookup:
            summaries += 1
            cells = []
            total_error = False
            for v, f in zip(row, formula):
                # Old line lookup formulas under a hand-written totals block
                # have no cable to look up. Do not export their #N/A placeholders.
                if is_error(v) and lookup_formula(f):
                    cells.append('')
                elif is_error(v):
                    cells.append(ERROR_DISPLAY.get(v,text(v)) if isinstance(v,int) else text(v))
                    total_error = True
                else:
                    cells.append(text(v))
            rows.append(cells)
            if total_error:
                warnings.append(f'Итоговая строка {row_number}: ошибка формулы перенесена в DWG.')
            continue
        if is_error(row[0]) and any(v not in (None,'',0) and not is_error(v) for v in row[1:4]):
            raise ValueError(f'{filename}, строка {row_number}: ошибка маркировки при заполненных данных трассы.')
        if row[0] in (None,'',0) or is_error(row[0]):
            continue
        # Exclude the first cable in the workbook before the optional panel filter.
        # Summary rows and empty formula placeholders do not consume this choice.
        if skip_first and excluded_first is None:
            excluded_first = {'mark':text(row[0]),'excel_row':row_number}
            continue
        if panel and text(row[1]).casefold()!=panel.casefold():
            continue
        if any(is_error(v) for v in row):
            columns=['P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB']
            bad=', '.join(columns[j]+str(row_number) for j,v in enumerate(row) if is_error(v))
            raise ValueError(f'{filename}, кабель {text(row[0])}: ошибка Excel в {bad}.')
        r=[text(v) for v in row]
        if not r[1] or not r[2] or not r[3]:
            warnings.append(f'Строка {row_number}, {r[0]}: не заполнено начало/конец/кабель.')
        if not r[7] or r[7]=='0':
            warnings.append(f'Строка {row_number}, {r[0]}: длина трассы {r[7] or "не задана"}.')
        rows.append(r)
        cable_markings.append(r[0])
    duplicates=[k for k,n in Counter(cable_markings).items() if n>1]
    if duplicates:
        raise ValueError(f'{filename}: повторяется маркировка кабеля: {", ".join(duplicates[:8])}. '
                         'Для безопасного обновления маркировка в одной книге должна быть уникальной.')
    return {'rows':rows,'warnings':warnings,'cable_count':len(cable_markings),'summary_count':summaries,
            'skip_first':bool(skip_first),'excluded_first':excluded_first}


def text(v):
    if v is None:
        return ''
    if isinstance(v, float):
        return format(v, '.10g').replace('.', ',')
    return str(v).strip()


def clean_mtext(v):
    s = text(v)
    s = re.sub(r'\\f[^;]*;', '', s)
    s = re.sub(r'\\[ACHWQT][^;]*;', '', s)
    return s.replace('\\P', '\n').replace('{', '').replace('}', '').strip()


def literal(v):
    # Escape AutoCAD MTEXT controls so Excel labels stay literal.
    return text(v).replace('\\', '\\\\').replace('{', '\\{').replace('}', '\\}').replace('\r', '').replace('\n', '\\P')


def point(x, y, z=0.):
    return com.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def source_files(source):
    p = Path(source).resolve()
    if p.is_file() and p.suffix.lower() in ('.xlsx', '.xlsm', '.xls'):
        return [p]
    if p.is_dir():
        books = sorted((q for q in p.iterdir() if q.suffix.lower() in ('.xlsx', '.xlsm', '.xls')
                        and not q.name.startswith('~$')), key=lambda q: q.name.casefold())
        if books:
            return books
    raise ValueError('Выберите файл Excel или папку с книгами Excel.')


def read_books(files, panel='', log=print, skip_first=False):
    """One private Excel instance, a single bulk read per book, no file writes."""
    xl = None
    result = []
    try:
        xl = com.DispatchEx('Excel.Application')
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        xl.AskToUpdateLinks = False
        xl.AutomationSecurity = 3
        for path in files:
            log(f'Читаю Excel: {path.name}')
            wb = None
            try:
                wb = xl.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True,
                                       Password='', WriteResPassword='', Notify=False,
                                       IgnoreReadOnlyRecommended=True, AddToMru=False)
                try:
                    sheet = wb.Worksheets.Item('Кабельный_журнал')
                except pywintypes.com_error:
                    log(f'Пропущено: в {path.name} нет листа «Кабельный_журнал».')
                    continue
                xl.CalculateFull()
                end = sheet.UsedRange.Row + sheet.UsedRange.Rows.Count - 1
                if end > 100000:
                    raise ValueError(f'{path.name}: слишком большой UsedRange, очистите лишние строки Excel.')
                headers = sheet.Range('P2:AB4').Value2
                if text(headers[0][0]) != 'Маркировка кабеля' or text(headers[2][7]) != 'Длина трассы, м':
                    raise ValueError(f'{path.name}: структура P2:AB4 отличается от принятого шаблона.')
                vals = sheet.Range(f'P5:AB{max(5,end)}').Value2
                formulas = sheet.Range(f'P5:AB{max(5,end)}').Formula
                parsed = parse_rows(vals,formulas,path.name,panel,skip_first)
                rows,warnings,summaries = parsed['rows'],parsed['warnings'],parsed['summary_count']
                if parsed['excluded_first']:
                    excluded = parsed['excluded_first']
                    log(f'Исключена первая кабельная строка Excel {excluded["excel_row"]}: {excluded["mark"]}. '
                        'Итоговые суммы переносятся как в Excel.')
                if rows:
                    result.append({'source':str(path),**parsed})
                    log(f'Кабельных линий: {parsed["cable_count"]}. Замечаний: {len(warnings)}.')
                    if summaries:
                        log(f'Включены итоговые строки Excel: {summaries}.')
                    for warning in warnings[:8]:
                        log(warning)
                else:
                    log(f'Пропущено: {path.name}, подходящих линий нет.')
            finally:
                if wb is not None:
                    wb.Close(SaveChanges=False)
        return result
    finally:
        if xl is not None:
            xl.Quit()
        xl = None
        gc.collect()


def autocad():
    try:
        app = com.GetActiveObject('AutoCAD.Application')
        retry(lambda: app.Version)
        return wrapped(app)
    except (pywintypes.com_error, AttributeError):
        app = com.Dispatch('AutoCAD.Application')
        app.Visible = True
        retry(lambda: app.Documents.Count, 60)
        return wrapped(app)


def dbx_open(app, path):
    major = text(retry(lambda: app.Version)).split('.')[0]
    db = retry(lambda: app.GetInterfaceObject(f'ObjectDBX.AxDbDocument.{major}'))
    db.Open(str(Path(path).resolve()))
    return db


def saved_actual(app, path):
    """Read saved as-built fields without activating/altering the user's drawing."""
    if not Path(path).exists():
        return {}, []
    db = dbx_open(app, path)
    actual = {}
    for tb in db.ModelSpace:
        if tb.ObjectName != 'AcDbTable' or tb.Columns != 13:
            continue
        for i in range(tb.Rows):
            mark = clean_mtext(tb.GetText(i, 0))
            if not mark or mark == 'Маркировка кабеля':
                continue
            if is_summary_row([mark,clean_mtext(tb.GetText(i,1)),clean_mtext(tb.GetText(i,2))]):
                continue
            fields = [tb.GetText(i, j) for j in range(8, 13)]
            if any(clean_mtext(v) for v in fields):
                if mark in actual and actual[mark] != fields:
                    raise ValueError(f'{path}: неоднозначные фактические данные кабеля {mark}.')
                actual[mark] = fields
    db = None
    gc.collect()
    return actual, []


def open_document(app, path):
    p = str(Path(path).resolve()).casefold()
    existing = retry(lambda: [d for d in documents(app) if text(d.FullName).casefold() == p])
    if not existing:
        def request_open():
            if not any(text(d.FullName).casefold()==p for d in documents(app)):
                app.Documents.Open(str(Path(path).resolve()))
        retry(request_open)
    # Documents.Open returns before all COM properties become accessible.
    def ready():
        d = next(d for d in documents(app) if text(d.FullName).casefold() == p)
        _ = d.ModelSpace.Count
        return d
    return retry(ready, 60)


def documents(app):
    collection = app.Documents
    return items(collection)


def items(collection):
    # COM enumerators become stale after objects are inserted/deleted.
    count=retry(lambda:collection.Count)
    return [retry(lambda i=i:collection.Item(i)) for i in range(count)]


def copy_objects(doc, objects, owner):
    objects = tuple(objects)
    if not objects:
        return []
    raw = doc.CopyObjects(com.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, objects), owner)
    # AutoCAD dynamic dispatch returns (copied array, IDPairs array).
    if isinstance(raw, tuple) and raw and isinstance(raw[0], tuple):
        raw = raw[0]
    return [wrapped(com.Dispatch(v)) for v in raw]


def page_chunks(rows, heights, max_rows, max_height):
    pages, page, used = [], [], 0.
    for r, h in zip(rows, heights):
        if h > max_height + .01:
            raise ValueError(f'Кабель {r[0]}: текст строки не помещается на листе А3.')
        if page and (len(page) >= max_rows or used + h > max_height + .01):
            pages.append(page)
            page, used = [], 0.
        page.append((r, h))
        used += h
    if page:
        pages.append(page)
    return pages


def fresh_table(model, prototype, rows):
    """Copy the appearance, without the old Excel data link's write locks."""
    widths=[prototype.GetColumnWidth(j) for j in range(13)]
    tb=model.AddTable(point(*prototype.InsertionPoint),rows,13,8.,30.)
    tb.RegenerateTableSuppressed=True
    tb.StyleName=prototype.StyleName
    tb.Layer=prototype.Layer
    tb.TitleSuppressed=True
    tb.HeaderSuppressed=True
    tb.UnmergeCells(0,0,0,12)
    tb.HorzCellMargin=prototype.HorzCellMargin
    tb.VertCellMargin=prototype.VertCellMargin
    for j,width in enumerate(widths):tb.SetColumnWidth(j,width)
    for i in range(rows):
        for j in range(13):
            tb.SetCellTextStyle(i,j,prototype.GetCellTextStyle(0,j))
            tb.SetCellTextHeight(i,j,2.5)
            tb.SetCellAlignment(i,j,prototype.GetCellAlignment(0,j))
    return tb


def prepare_drawing(doc, rows, title, max_rows, actual, log=print):
    """Regenerate a private copy, retaining template tables and printed layouts."""
    model = doc.ModelSpace
    tables = [e for e in items(model) if e.ObjectName == 'AcDbTable' and e.Columns == 13]
    headers = [e for e in tables if e.Rows == 3 and clean_mtext(e.GetText(0,0)) == 'Маркировка кабеля']
    bodies = [e for e in tables if e.Rows > 3 or (e.Rows >= 1 and clean_mtext(e.GetText(0,0)) != 'Маркировка кабеля')]
    if not headers or not bodies:
        raise ValueError('В шаблоне нужны две таблицы в модели: шапка 3×13 и строки кабелей 13 столбцов.')
    header = min(headers, key=lambda e: e.InsertionPoint[0])
    body = min(bodies, key=lambda e: e.InsertionPoint[0])
    layouts = sorted((l for l in items(doc.Layouts) if not l.ModelType), key=lambda l:l.TabOrder)
    if not layouts:
        raise ValueError('В шаблоне отсутствует настроенный лист для печати.')
    base = layouts[0]
    template_viewports = [e for e in items(base.Block) if e.ObjectName == 'AcDbViewport' and e.DisplayLocked]
    if len(template_viewports) != 1:
        raise ValueError('На первом листе шаблона нужен один заблокированный видовой экран таблицы.')
    base_vp = template_viewports[0]
    if abs(base_vp.CustomScale - 1.) > .001 or tuple(base_vp.Direction) != (0.,0.,1.):
        raise ValueError('Шаблон должен иметь верхний вид и масштаб видового экрана 1:1.')
    if abs(sum(body.GetColumnWidth(j) for j in range(13))-395.) > .1:
        raise ValueError('Этот профиль рассчитан на таблицу шириной 395 мм из вашего образца.')
    x0, top, z0 = body.InsertionPoint
    hx, hy, hz = header.InsertionPoint
    # In this A3 profile, the frame's upper title block edge is at y=52 mm.
    max_height = top - 52.
    if max_height <= 0:
        raise ValueError('Неверная точка вставки строк в шаблоне.')
    base_target = tuple(base_vp.Target)
    doc.ActiveLayout = base
    doc.SetVariable('TILEMODE', 0)
    for layer in doc.Layers:
        if layer.Name in {e.Layer for e in tables}:
            layer.Lock = False
    # First fill a temporary continuous table to obtain native text row heights.
    old_body=body
    body=fresh_table(model,old_body,len(rows))
    old_body.Delete()
    body.BreaksEnabled = False
    body.RegenerateTableSuppressed = True
    if body.Rows > len(rows):
        body.DeleteRows(len(rows), body.Rows-len(rows))
    elif body.Rows < len(rows):
        body.InsertRows(body.Rows, 8., len(rows)-body.Rows)
    log(f'Заполняю {body.Rows} строк таблицы.')
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            try:
                body.SetText(i, j, actual.get(row[0], [None]*5)[j-8]
                             if j >= 8 and row[0] in actual else ('{\\fGOST 2.304 type A|b0|i0|c0;'+literal(value)+'}' if value else ''))
            except Exception as exc:
                raise RuntimeError(f'Ячейка DWG {i+1}:{j+1}, строк таблицы {body.Rows}: {exc}') from exc
        body.SetRowHeight(i, 8.)
    body.RegenerateTableSuppressed = False
    body.Update()
    heights = [max(8., body.GetMinimumRowHeight(i), body.GetRowHeight(i)) for i in range(body.Rows)]
    pages = page_chunks(rows, heights, max_rows, max_height)
    # Snapshot old annotations, then replace only the tables and pagination guides.
    for e in items(model):
        if e.Handle not in (header.Handle, body.Handle) and (e.ObjectName=='AcDbTable' or e.Layer=='Defpoints'):
            e.Delete()
    for e in items(base.Block):
        if e.ObjectName == 'AcDbTable':
            e.Delete()  # old manually pasted header from later template pages
    for layout in layouts[1:]:
        layout.Delete()
    base.Name = 'КЖ-01'
    for page_index, page in enumerate(pages):
        offset = 420. * page_index
        hcopy = copy_objects(doc, [header], model)[0]
        bcopy = copy_objects(doc, [body], model)[0]
        hcopy.InsertionPoint = point(hx+offset, hy, hz)
        bcopy.InsertionPoint = point(x0+offset, top, z0)
        bcopy.RegenerateTableSuppressed = True
        if bcopy.Rows > len(page):
            bcopy.DeleteRows(len(page), bcopy.Rows-len(page))
        for i, (row, height) in enumerate(page):
            for j, value in enumerate(row):
                bcopy.SetText(i,j,actual[row[0]][j-8] if j>=8 and row[0] in actual else ('{\\fGOST 2.304 type A|b0|i0|c0;'+literal(value)+'}' if value else ''))
            bcopy.SetRowHeight(i, height)
        bcopy.RegenerateTableSuppressed = False
        bcopy.Update()
        if bcopy.Height > max_height + .05:
            raise ValueError(f'Лист {page_index+1}: таблица пересекает штамп. Уменьшите число строк на листе.')
        if page_index == 0:
            layout = base
        else:
            layout = doc.Layouts.Add(f'КЖ-{page_index+1:02d}')
            layout.CopyFrom(base)
            entities = [e for e in items(base.Block) if e.ObjectName != 'AcDbViewport' or e.DisplayLocked]
            copies = copy_objects(doc, entities, layout.Block)
            for e in copies:
                if e.ObjectName == 'AcDbViewport':
                    e.DisplayLocked = False
                    e.Target = point(base_target[0]+offset, base_target[1], base_target[2])
                    e.DisplayLocked = True
        layout.TabOrder = page_index + 1
        for e in items(layout.Block):
            if e.ObjectName == 'AcDbBlockReference' and e.HasAttributes:
                for att in e.GetAttributes():
                    tag = att.TagString.upper()
                    if tag == 'ОПИСАНИЕ_ЛИСТА':
                        att.TextString = title
                    elif tag in ('ЛИСТ_№','ЛИСТ№'):
                        att.TextString = str(page_index+1)
                    elif tag == 'КОЛ-ВО_ЛИСТОВ':
                        att.TextString = str(len(pages))
                    elif tag == 'СКВОЗНАЯ_НУМЕРЦИЯ':
                        att.TextString = ''
                    att.Update()
        log(f'Лист {page_index+1}/{len(pages)}: {len(page)} строк.')
    header.Delete()
    body.Delete()
    doc.ActiveLayout = base
    retry(lambda:doc.Regen(1))
    # Verify every exported project cell and retained manual field before saving.
    saved = []
    for tb in sorted((e for e in items(model) if e.ObjectName=='AcDbTable'
                      and clean_mtext(e.GetText(0,0))!='Маркировка кабеля'),
                     key=lambda e:e.InsertionPoint[0]):
        saved.extend([[clean_mtext(retry(lambda i=i,j=j:tb.GetText(i,j))) for j in range(13)] for i in range(tb.Rows)])
    expected = [[text(v) for v in row[:8]] + [clean_mtext(v) for v in actual.get(row[0],row[8:])] for row in rows]
    # MTEXT literal escapes are normalized here, without altering cable identifiers.
    saved = [[v.replace('\\\\','\\').replace('\\{','{').replace('\\}','}') for v in r] for r in saved]
    if saved != expected:
        raise RuntimeError('Проверка переноса данных Excel → DWG не прошла. Готовый DWG не заменён.')
    return len(pages)


def manifest_path(path):
    return Path(str(path)+MANIFEST_SUFFIX)


def write_json(path, value):
    path = Path(path)
    tmp = path.with_name(path.name+'.tmp')
    tmp.write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf-8')
    os.replace(tmp,path)


def publish_files(work, report_file, output, backup=None):
    """Stage ordinary files directly in the destination for same-volume renames.

    Windows private tempfile directory ACLs can be incompatible with SMB shares.
    Copy only contents, without transferring local permissions or file metadata.
    """
    token = uuid.uuid4().hex
    staged_dwg = output.parent/f'.cable_journal_{token}.dwg.tmp'
    staged_report = output.parent/f'.cable_journal_{token}.json.tmp'
    staged = (staged_dwg, staged_report)
    try:
        for source, destination in zip((work, report_file), staged):
            with Path(source).open('rb') as src, destination.open('xb') as dst:
                shutil.copyfileobj(src, dst)
                dst.flush()
                os.fsync(dst.fileno())
        os.replace(staged_dwg, output)
        try:
            os.replace(staged_report, manifest_path(output))
        except OSError:
            # Restore the old drawing if its accompanying report cannot be published.
            if backup:
                shutil.copyfile(backup, staged_dwg)
                os.replace(staged_dwg, output)
            else:
                output.unlink()
            raise
    finally:
        for path in staged:
            try: path.unlink(missing_ok=True)
            except OSError: pass


def record_error():
    """Keep the complete traceback without masking the original failure."""
    path = Path(os.environ.get('APPDATA',str(APP_DIR)))/'CableJournal'/'errors.log'
    try:
        path.parent.mkdir(parents=True,exist_ok=True)
        with path.open('a',encoding='utf-8') as f:
            f.write(f'\n{datetime.now().isoformat()} — {VERSION}\n{traceback.format_exc()}')
    except OSError:
        pass


def export_book(app, book, template, output, panel='', max_rows=25, log=print):
    started = time.monotonic()
    output = Path(output).resolve()
    template = Path(template).resolve()
    if output == template or output.suffix.lower() != '.dwg':
        raise ValueError('Готовый DWG должен отличаться от файла шаблона.')
    output.parent.mkdir(parents=True,exist_ok=True)
    current = retry(lambda: [d for d in documents(app) if text(d.FullName).casefold()==str(output).casefold()])
    if current:
        raise ValueError(f'Закройте готовый журнал {output.name} в AutoCAD перед обновлением. '
                         'Это нужно, чтобы сохранить ваши правки и заменить файл.')
    updating = output.exists()
    if updating and not manifest_path(output).exists():
        raise ValueError(f'{output.name} создан вне программы. Выберите другой выходной файл; '
                         'образец можно использовать в поле «Шаблон DWG».')
    actual, _ = saved_actual(app,output)
    removed = {k:v for k,v in actual.items() if k not in {r[0] for r in book['rows']}}
    manual_conflicts = []
    for row in book['rows']:
        if row[0] in actual and any(clean_mtext(a) and text(v) and clean_mtext(a)!=text(v)
                                  for a,v in zip(actual[row[0]],row[8:])):
            manual_conflicts.append(row[0])
    # Work exclusively in a private temporary copy. Existing output stays intact on failure.
    scratch = Path(tempfile.mkdtemp(prefix='cable_journal_'))
    work = scratch/'journal.dwg'
    doc = None
    previous = None
    keep_work = False
    try:
        shutil.copyfile(output if updating else template,work)
        previous = retry(lambda: app.ActiveDocument)
        doc = open_document(app,work)
        retry(lambda: doc.Activate())
        title = 'Кабельный журнал ' + (panel or ', '.join(dict.fromkeys(r[1] for r in book['rows'] if r[1])))
        pages = prepare_drawing(doc,book['rows'],title,max_rows,actual,log)
        retry(lambda:doc.Save())
        retry(lambda: doc.Close(False))
        doc = None
        backup = None
        if updating:
            stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            backup_dir = output.parent/'Резервные_копии'
            backup_dir.mkdir(exist_ok=True)
            backup = backup_dir/f'{output.stem}_{stamp}.dwg'
            shutil.copyfile(output,backup)
            shutil.copyfile(manifest_path(output),manifest_path(backup))
        report = {'version':VERSION,'source':book['source'],'template':str(template),
                  'output':str(output),'panel':panel,'max_rows':max_rows,
                  'skip_first':book.get('skip_first',False),'excluded_first':book.get('excluded_first'),
                  'lines':book.get('cable_count',len(book['rows'])),'summary_rows':book.get('summary_count',0),
                  'total_rows':len(book['rows']),'pages':pages,'updated_at':datetime.now().isoformat(),
                  'seconds':round(time.monotonic()-started,2),'warnings':book['warnings'],
                  'preserved_actual':list(actual),'actual_conflicts_dwg_kept':manual_conflicts,
                  'removed_cables_actual_archived':removed,'backup':str(backup) if backup else None}
        # Preserve the report before publishing the drawing; recover it if final rename fails.
        report_tmp = scratch/'report.json'
        write_json(report_tmp,report)
        try:
            publish_files(work,report_tmp,output,backup)
        except OSError as exc:
            keep_work = True
            raise RuntimeError(f'Не удалось сохранить журнал в {output.parent}: {exc}. '
                               f'Собранный DWG сохранён локально: {work}') from exc
        log(f'Готово: {output.name}; {report["lines"]} линий, {pages} листов, {report["seconds"]} с.')
        if actual:
            log(f'Сохранены ручные графы DWG: {len(actual)-len(removed)} кабелей.')
        if removed:
            log(f'Удалённые кабели с ручными данными сохранены в отчёте: {len(removed)}.')
        return report
    finally:
        if doc is not None:
            try: retry(lambda: doc.Close(False))
            except Exception: pass
        try: retry(lambda: previous.Activate())
        except Exception: pass
        # Delete only this verified tempfile directory, never template or user inputs.
        if not keep_work and scratch.parent == Path(tempfile.gettempdir()).resolve() and scratch.name.startswith('cable_journal_'):
            shutil.rmtree(scratch,ignore_errors=True)


def run_job(source, template, outdir, panel='', max_rows=25, target=None, log=print, skip_first=False):
    if not 1 <= max_rows <= 25:
        raise ValueError('Число строк на листе должно быть от 1 до 25.')
    if not Path(template).is_file():
        raise ValueError('Не найден шаблон DWG.')
    files = source_files(source)
    books = read_books(files,panel.strip(),log,skip_first)
    if not books:
        raise ValueError('В выбранных книгах нет подходящих строк кабельного журнала.')
    if target and len(books)!=1:
        raise ValueError('Для обновления одного DWG нужен один исходный файл Excel.')
    app = autocad()
    reports, failures = [], []
    for book in books:
        name = Path(book['source']).stem
        name = re.sub(r'^Однолинейка[_ ]*','',name,flags=re.I)
        filename = re.sub(r'[<>:"/\\|?*]', '_',f'КЖ_{name}'+(f'_{panel}' if panel else ''))+'.dwg'
        output = Path(target) if target else Path(outdir)/filename
        try:
            reports.append(export_book(app,book,template,output,panel,max_rows,log))
        except Exception as exc:
            record_error()
            failures.append(f'{name}: {exc}')
            log(f'Ошибка: {failures[-1]}')
    if failures:
        raise RuntimeError(f'Создано/обновлено DWG: {len(reports)}. Ошибки:\n'+'\n'.join(failures))
    return reports


def gui():
    import tkinter as tk
    from tkinter import ttk,filedialog,messagebox
    root=tk.Tk()
    root.title(f'Кабельный журнал — Excel → DWG {VERSION}')
    root.geometry('820x650')
    root.minsize(750,610)
    settings_path=Path(os.environ.get('APPDATA',str(APP_DIR)))/'CableJournal'/'settings.json'
    try:settings=json.loads(settings_path.read_text(encoding='utf-8'))
    except (OSError,ValueError):settings={}
    source=tk.StringVar(value=settings.get('source',''))
    template=tk.StringVar(value=settings.get('template',str(DEFAULT_TEMPLATE)))
    output=tk.StringVar(value=settings.get('output',str(APP_DIR/'Результат')))
    panel=tk.StringVar(value=settings.get('panel',''))
    count=tk.StringVar(value=str(settings.get('max_rows',25)))
    skip_first=tk.BooleanVar(value=settings.get('skip_first',False))
    frame=ttk.Frame(root,padding=18);frame.pack(fill='both',expand=True)
    ttk.Label(frame,text='Кабельный журнал из ваших Excel',font=('Segoe UI',17,'bold')).pack(anchor='w')
    ttk.Label(frame,text='Таблицы AutoCAD, рамки А3 и видовые экраны из вашего образца.').pack(anchor='w',pady=(4,16))
    def picker_row(label,var,kind):
        f=ttk.Frame(frame);f.pack(fill='x',pady=5)
        ttk.Label(f,text=label,width=18).pack(side='left')
        ttk.Entry(f,textvariable=var).pack(side='left',fill='x',expand=True)
        def pick_file():
            path=filedialog.askopenfilename(filetypes=[('Excel','*.xlsx *.xlsm *.xls')]) if kind=='excel' else filedialog.askopenfilename(filetypes=[('AutoCAD','*.dwg')])
            if path:var.set(path)
        if kind=='dir':
            def select_dir():
                path=filedialog.askdirectory()
                if path:var.set(path)
            ttk.Button(f,text='Папка…',command=select_dir,width=10).pack(side='left',padx=(6,0))
        else:
            ttk.Button(f,text='Файл…',command=pick_file,width=10).pack(side='left',padx=(6,0))
            if kind=='excel':
                def select_dir():
                    path=filedialog.askdirectory()
                    if path:var.set(path)
                ttk.Button(f,text='Папка…',command=select_dir,width=10).pack(side='left',padx=(6,0))
    picker_row('Excel / папка Excel',source,'excel')
    picker_row('Шаблон DWG',template,'dwg')
    picker_row('Папка результата',output,'dir')
    f=ttk.Frame(frame);f.pack(fill='x',pady=(10,4))
    ttk.Label(f,text='Начало трассы',width=18).pack(side='left')
    ttk.Entry(f,textvariable=panel,width=26).pack(side='left')
    ttk.Label(f,text='пусто — все строки книги').pack(side='left',padx=8)
    f=ttk.Frame(frame);f.pack(fill='x',pady=5)
    ttk.Label(f,text='Строк на листе',width=18).pack(side='left')
    ttk.Spinbox(f,from_=1,to=25,textvariable=count,width=5).pack(side='left')
    ttk.Label(f,text='длинные строки автоматически переносятся на следующий лист').pack(side='left',padx=8)
    ttk.Checkbutton(frame,text='Не включать первую кабельную строку Excel',variable=skip_first).pack(anchor='w',pady=(6,0))
    ttk.Label(frame,text='Из каждой книги; итоговые суммы остаются как в Excel.').pack(anchor='w',padx=20)
    ttk.Label(frame,text='Обновление сохраняет графы «Проложено», испытания и дату из DWG.\nПеред заменой создаётся резервная копия. Excel и исходный образец не меняются.',wraplength=760).pack(anchor='w',pady=(10,12))
    buttons=ttk.Frame(frame);buttons.pack(fill='x')
    status=tk.StringVar(value='Выберите исходный Excel или папку щитов.')
    ttk.Label(frame,textvariable=status).pack(anchor='w',pady=(12,5))
    log_box=tk.Text(frame,height=10,font=('Segoe UI',10),wrap='word',state='disabled')
    log_box.pack(fill='both',expand=True)
    q=queue.Queue();busy=False
    def add_log(s):
        log_box.configure(state='normal');log_box.insert('end',s+'\n');log_box.see('end');log_box.configure(state='disabled')
    def start(target=None,config=None):
        nonlocal busy
        if busy:return
        try:
            cfg=config or {'source':source.get().strip(),'template':template.get().strip(),
                           'outdir':output.get().strip(),'panel':panel.get().strip(),'max_rows':int(count.get()),
                           'skip_first':skip_first.get()}
            if not cfg['source'] or not cfg['outdir']:raise ValueError('Выберите Excel и папку результата.')
            settings_path.parent.mkdir(parents=True,exist_ok=True)
            write_json(settings_path,{'source':cfg['source'],'template':cfg['template'],'output':cfg['outdir'],
                                      'panel':cfg['panel'],'max_rows':cfg['max_rows'],'skip_first':cfg['skip_first']})
        except (ValueError,OSError) as exc:messagebox.showerror('Кабельный журнал',str(exc));return
        busy=True
        create_btn.configure(state='disabled');update_btn.configure(state='disabled')
        status.set('Формирую журнал…')
        def worker():
            pythoncom.CoInitialize()
            try:q.put(('done',run_job(**cfg,target=target,log=lambda s:q.put(('log',s)))))
            except Exception as exc:
                record_error()
                q.put(('error',str(exc)))
            finally:pythoncom.CoUninitialize()
        threading.Thread(target=worker,daemon=False).start()
    def update_existing():
        path=filedialog.askopenfilename(title='Обновить DWG, созданный этой программой',filetypes=[('AutoCAD','*.dwg')])
        if not path:return
        try:
            m=json.loads(manifest_path(path).read_text(encoding='utf-8'))
            cfg={'source':m['source'],'template':m['template'],'outdir':str(Path(path).parent),
                 'panel':m['panel'],'max_rows':m['max_rows'],'skip_first':skip_first.get()}
        except (OSError,ValueError,KeyError):
            messagebox.showerror('Нет связи с Excel','Рядом с DWG не найден его файл .kj.json. Создайте журнал через программу и держите эти два файла вместе.');return
        start(path,cfg)
    create_btn=ttk.Button(buttons,text='Создать / обновить по Excel',command=start);create_btn.pack(side='left')
    update_btn=ttk.Button(buttons,text='Обновить готовый DWG…',command=update_existing);update_btn.pack(side='left',padx=8)
    ttk.Button(buttons,text='Открыть результат',command=lambda:os.startfile(output.get()) if Path(output.get()).is_dir() else None).pack(side='right')
    def poll():
        nonlocal busy
        while not q.empty():
            typ,value=q.get()
            if typ=='log':add_log(value)
            else:
                busy=False;create_btn.configure(state='normal');update_btn.configure(state='normal')
                if typ=='done':status.set(f'Готово. DWG: {len(value)}. Все данные проверены.');add_log('Готовые DWG и связь с Excel сохранены в папке результата.')
                else:status.set('Операция не завершена.');add_log(value);messagebox.showerror('Кабельный журнал',value)
        root.after(150,poll)
    def close():
        if busy:messagebox.showinfo('Кабельный журнал','Дождитесь окончания текущего формирования.');return
        root.destroy()
    root.protocol('WM_DELETE_WINDOW',close)
    poll();root.mainloop()


def main():
    parser=argparse.ArgumentParser(description='Кабельный журнал Excel → DWG')
    parser.add_argument('--source');parser.add_argument('--template',default=str(DEFAULT_TEMPLATE))
    parser.add_argument('--outdir',default=str(APP_DIR/'Результат'));parser.add_argument('--panel',default='')
    parser.add_argument('--max-rows',type=int,default=25);parser.add_argument('--target')
    parser.add_argument('--skip-first',action='store_true',help='Не включать первую кабельную строку каждой книги Excel')
    args=parser.parse_args()
    if args.source:
        pythoncom.CoInitialize()
        try:run_job(args.source,args.template,args.outdir,args.panel,args.max_rows,args.target,skip_first=args.skip_first)
        finally:pythoncom.CoUninitialize()
    else:gui()


if __name__=='__main__':
    main()
