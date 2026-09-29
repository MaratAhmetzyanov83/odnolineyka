# -*- coding: utf-8 -*-
"""check.py - эталонная проверка (регресс-тест) книг «Однолинейка».

    py check.py                         проверить все эталоны (baseline/*.json)
    py check.py --etalon ostrov         только один эталон
    py check.py --book Новый.xlsx       другую книгу (напр. новый шаблон с данными эталона)
                                        против baseline эталона (--etalon, по умолчанию ostrov)
    py check.py --update-baseline       принять текущее состояние как эталон
    py check.py --log SX_log.txt        сверить журнал LISP (SXDRAW) с ожиданиями из снимка

Код возврата: 0 - изменений нет, 1 - есть изменения / ошибка.
"""
import argparse
import json
import os
import re
import sys
from collections import OrderedDict

import snapshot

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))        # корень репозитория
BASELINE_DIR = os.path.join(HERE, "baseline")

# имя эталона -> путь к книге (от корня репозитория)
ETALONS = OrderedDict([
    ("ostrov", os.path.join("1_Таблицы_Excel", "Однолинейка_пример_ЖК-Остров.xlsx")),
    ("grsh_zhd", os.path.join("1_Таблицы_Excel", "Объект ЖД", "Однолинейка ГРЩ-ЖД.xlsx")),
])

SECTION_TITLES = OrderedDict([
    ("мета", "Структура книги"),
    ("ошибки в ячейках", "Ошибки в ячейках"),
    ("счётчики", "Счётчики"),
    ("статус (светофор)", "«Однолинейка» - статус (светофор)"),
    ("однолинейка (строки)", "«Однолинейка»"),
    ("исходные данные (линии)", "«Исходные данные»"),
    ("нагрузки щита", "«Нагрузка_щитов»"),
    ("расчёт нагрузок (ПЗ)", "«Расчёт_нагрузок»"),
    ("питание щита", "«Питание_щита»"),
    ("нижестоящие щиты", "«Нижестоящие_щиты»"),
    ("в акад (таблица нагрузок)", "«В Акад» (таблица для AutoCAD)"),
    ("спецификация (позиция → кол-во)", "«Спецификация_авто»"),
    ("спецификация: подсчёт", "«Спецификация_авто» - подсчёт"),
    ("спецификация: предупреждение", "«Спецификация_авто» - предупреждение"),
])
SKIP = {"формат", "ожидания LISP (SX_log)"}
INFO_ONLY = {("мета", "версия шаблона")}


# ---- разворачивание снимка в (раздел, строка, поле) -> значение --------------------------------
def flat(snap):
    out = OrderedDict()

    def add(sec, row, field, val):
        out[(sec, row, field)] = val

    for sec, obj in snap.items():
        if sec in SKIP:
            continue
        if sec == "мета":
            add(sec, "версия шаблона", "значение", obj.get("версия шаблона"))
            add(sec, "листы книги", "список", ", ".join(obj.get("листы", [])))
            for sheet, cols in obj.get("заголовки", {}).items():
                for c, h in cols.items():
                    add(sec, "заголовки «%s»" % sheet, "столбец " + c, h)
        elif sec == "ошибки в ячейках":
            for sheet, d in obj.items():
                add(sec, sheet, "всего", d["всего"])
                for col, types in d["по столбцам"].items():
                    for t, n in types.items():
                        add(sec, sheet, "столбец %s: %s" % (col, t), n)
        elif sec == "счётчики":
            for grp, d in obj.items():
                for k, v in d.items():
                    add(sec, grp, k, v)
        elif sec == "статус (светофор)":
            add(sec, "сводка", "AE1", obj.get("сводка AE1"))
            for k, v in obj.get("строк по статусам (светофор, AE)", {}).items():
                add(sec, "строк по статусам", k, v)
        elif sec == "спецификация: предупреждение":
            add(sec, "C1", "текст", obj)
        elif isinstance(obj, dict):
            for row, d in obj.items():
                if isinstance(d, dict):
                    for f, v in d.items():
                        add(sec, row, f, v)
                else:
                    add(sec, row, "кол-во", d)
    return out


def fmt(v):
    if v is None:
        return "—"
    if isinstance(v, float):
        return format(v, ".6g")
    return str(v)


def equal(a, b, tol):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool):
        return abs(a - b) <= tol + 1e-9 * max(abs(a), abs(b))
    return a == b


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def line_name(snap, sec, row):
    """Подпись строки: для «Однолинейки» добавляем номер линии."""
    sd = snap.get(sec)
    d = sd.get(row) if isinstance(sd, dict) else None
    if isinstance(d, dict) and d.get("Линия"):
        return "%s [%s]" % (row, d["Линия"])
    return row


def diff(base, cur, tol):
    """-> (изменения: [(раздел, строка, [(поле, было, стало)], тип)], инфо: [строка])"""
    fb, fc = flat(base), flat(cur)
    rows = OrderedDict()
    info = []
    keys = list(fb.keys()) + [k for k in fc.keys() if k not in fb]
    for k in keys:
        sec, row, f = k
        a, b = fb.get(k), fc.get(k)
        # для счётчиков ошибок отсутствие = 0
        if sec == "ошибки в ячейках" and f != "всего":
            a = 0 if a is None else a
            b = 0 if b is None else b
        if equal(a, b, tol):
            continue
        if (sec, row) in INFO_ONLY:
            info.append("Версия шаблона: %s → %s (справочно)" % (fmt(a), fmt(b)))
            continue
        rows.setdefault((sec, row), []).append((f, a, b))
    changes = []
    base_rows = {(s, r) for (s, r, _f) in fb}
    cur_rows = {(s, r) for (s, r, _f) in fc}
    for (sec, row), fields in rows.items():
        if (sec, row) not in cur_rows:
            kind = "исчезла"
        elif (sec, row) not in base_rows:
            kind = "появилась"
        else:
            kind = "изменена"
        snap = cur if kind != "исчезла" else base
        changes.append((sec, line_name(snap, sec, row), fields, kind))
    order = list(SECTION_TITLES.keys())
    changes.sort(key=lambda c: order.index(c[0]) if c[0] in order else 99)
    return changes, info


def render(name, book_name, changes, info, max_lines, full, row_limit=10, field_limit=6):
    lines = []
    total_fields = sum(len(c[2]) for c in changes)
    if not changes:
        lines.append("[%s] OK — изменений нет  (%s)" % (name, book_name))
        lines += ["    " + i for i in info]
        return lines
    lines.append("[%s] ИЗМЕНЕНИЯ: %d полей в %d строках  (%s)" % (name, total_fields, len(changes), book_name))
    lines += ["    " + i for i in info]
    cur_sec = None
    shown_rows = 0
    budget = 10 ** 9 if full else max_lines
    hidden_fields = hidden_rows = 0
    for sec, row, fields, kind in changes:
        if sec != cur_sec:
            n_rows = sum(1 for c in changes if c[0] == sec)
            n_f = sum(len(c[2]) for c in changes if c[0] == sec)
            lines.append("  Лист/раздел %s: %d полей в %d строках" % (SECTION_TITLES.get(sec, sec), n_f, n_rows))
            cur_sec, shown_rows = sec, 0
        if (not full and shown_rows >= row_limit) or len(lines) >= budget:
            hidden_rows += 1
            hidden_fields += len(fields)
            continue
        shown_rows += 1
        if kind in ("появилась", "исчезла"):
            sample = ", ".join("%s=%s" % (f, fmt(b if kind == "появилась" else a)) for f, a, b in fields[:3])
            lines.append("    строка %s: %s (%d полей: %s%s)" % (row, kind, len(fields), sample,
                                                                ", …" if len(fields) > 3 else ""))
            continue
        parts = ["%s: %s → %s" % (f, fmt(a), fmt(b)) for f, a, b in (fields if full else fields[:field_limit])]
        more = "" if full or len(fields) <= field_limit else "; … ещё %d" % (len(fields) - field_limit)
        lines.append("    строка %s — %s%s" % (row, "; ".join(parts), more))
    if hidden_rows:
        lines.append("  … ещё %d строк (%d полей) не показано; полный список: --full или --report файл.txt"
                     % (hidden_rows, hidden_fields))
    return lines


# ---- журнал LISP -------------------------------------------------------------------------------
def unescape_lisp(s):
    return re.sub(r"\\U\+([0-9A-Fa-f]{4})", lambda m: chr(int(m.group(1), 16)), s)


def parse_log(path):
    txt = None
    for enc in ("utf-8-sig", "cp1251", "cp866"):
        try:
            with open(path, encoding=enc) as f:
                txt = f.read()
            break
        except UnicodeDecodeError:
            continue
    if txt is None:
        raise RuntimeError("не удалось прочитать журнал " + path)
    r = {"errors": [], "fatal": [], "missing": [], "panel": None, "recs": None, "draw": None, "sheets": None}
    for ln in txt.splitlines():
        ln = ln.rstrip()
        if ln.startswith("ERROR "):
            r["errors"].append(ln[6:])
        elif ln.startswith("FATAL"):
            r["fatal"].append(ln)
        m = re.match(r"block (\S+) MISSING", ln)
        if m:
            r["missing"].append(m.group(1))
        m = re.match(r"recs=(\d+) draw=(\d+)", ln)
        if m:
            r["recs"], r["draw"] = int(m.group(1)), int(m.group(2))
        m = re.match(r"sheets=(.*)$", ln)
        if m:
            r["sheets"] = sorted(int(x) for x in re.findall(r"\d+", m.group(1)))
        m = re.match(r'first place \d+ row=\("((?:[^"\\]|\\.)*)"', ln)
        if m:
            r["panel"] = unescape_lisp(m.group(1))
    return r


def check_log(path, expect, panel_hint=None):
    """expect: {панель: {recs, draw, sheets}} -> (строки отчёта, число проблем)"""
    lg = parse_log(path)
    out, bad = [], 0
    if lg["fatal"]:
        bad += len(lg["fatal"])
        out += ["    " + x for x in lg["fatal"]]
    if lg["errors"]:
        bad += len(lg["errors"])
        out.append("    ошибок этапов в журнале (ERROR): %d" % len(lg["errors"]))
        out += ["      " + e[:150] for e in lg["errors"][:5]]
    if lg["missing"]:
        out.append("    ВНИМАНИЕ: нет блоков (заменители): %s" % ", ".join(lg["missing"]))
    panel = lg["panel"] or panel_hint
    if panel not in expect:
        if len(expect) == 1:
            panel = next(iter(expect))
        else:
            out.append("    щит журнала (%r) не найден среди ожидаемых: %s" % (panel, ", ".join(expect)))
            return out, bad + 1
    e = expect[panel]
    for key, label in (("recs", "recs (строк щита)"), ("draw", "draw (рисуемых строк)")):
        if lg[key] is None:
            out.append("    в журнале нет строки recs=/draw= (журнал неполный)")
            bad += 1
            break
        if lg[key] != e[key]:
            out.append("    %s: ожидалось %s → в журнале %s" % (label, e[key], lg[key]))
            bad += 1
    if lg["sheets"] is None:
        out.append("    в журнале нет строки sheets= (этап «sheets» не выполнен)")
        bad += 1
    elif lg["sheets"] != e["sheets"]:
        out.append("    листы схемы (индексы с 0): ожидалось %s → в журнале %s" % (e["sheets"], lg["sheets"]))
        bad += 1
    head = "SX_log (щит %s): %s" % (panel, "ОШИБКИ" if bad else "OK — счётчики и листы совпали со снимком")
    return [head] + out, bad


# ---- основное ----------------------------------------------------------------------------------
def baseline_path(name):
    return os.path.join(BASELINE_DIR, name + ".json")


def load_baseline(name):
    p = baseline_path(name)
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--etalon", action="append", choices=list(ETALONS), help="имя эталона (можно несколько)")
    ap.add_argument("--book", help="другая книга для проверки против baseline эталона (--etalon, по умолчанию ostrov)")
    ap.add_argument("--update-baseline", action="store_true", help="принять текущее состояние как эталон")
    ap.add_argument("--log", help="SX_log.txt: сверить счётчики LISP с ожиданиями из снимка")
    ap.add_argument("--backend", default="auto", choices=["auto", "excel", "libreoffice", "cache"])
    ap.add_argument("--tol", type=float, default=1e-5, help="допуск для чисел (по умолчанию 1e-5)")
    ap.add_argument("--max-lines", type=int, default=70, help="лимит строк отчёта (по умолчанию 70)")
    ap.add_argument("--full", action="store_true", help="показать все различия")
    ap.add_argument("--report", help="записать полный отчёт в файл")
    a = ap.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    names = a.etalon or list(ETALONS)
    if a.book:
        names = (a.etalon or ["ostrov"])[:1]
    full_report, rc = [], 0

    # только журнал: без пересчёта книги, ожидания берём из baseline
    if a.log and not a.book and not a.update_baseline and not a.etalon:
        expect = OrderedDict()
        for n in ETALONS:
            b = load_baseline(n)
            if b:
                for p, v in b.get("ожидания LISP (SX_log)", {}).items():
                    expect.setdefault(p, v)
        lines, bad = check_log(a.log, expect)
        print("\n".join(lines))
        return 1 if bad else 0

    for n in names:
        path = os.path.abspath(a.book) if a.book else os.path.join(ROOT, ETALONS[n])
        if not os.path.exists(path):
            full_report.append("[%s] ОШИБКА: книга не найдена: %s" % (n, path))
            rc = 1
            continue
        try:
            cur, warns = snapshot.build(path, a.backend, log=lambda m: print("  " + m))
        except Exception as e:                                      # noqa
            full_report.append("[%s] ОШИБКА: не удалось прочитать книгу: %s" % (n, e))
            rc = 1
            continue
        full_report += ["  " + w for w in warns]
        base = load_baseline(n)
        if a.update_baseline:
            if base is not None:
                ch, info = diff(base, cur, a.tol)
                full_report += render(n, os.path.basename(path), ch, info, 15, False)
            os.makedirs(BASELINE_DIR, exist_ok=True)
            cur["мета"]["книга"] = ETALONS[n].replace("\\", "/") if not a.book else cur["мета"]["книга"]
            snapshot.dump(cur, baseline_path(n))
            full_report.append("[%s] baseline обновлён: %s" % (n, baseline_path(n)))
        elif base is None:
            full_report.append("[%s] нет baseline (%s). Создайте: py check.py --update-baseline" % (n, baseline_path(n)))
            rc = 1
        else:
            ch, info = diff(base, cur, a.tol)
            full_report += render(n, os.path.basename(path), ch, info, a.max_lines, a.full)
            if ch:
                rc = 1
            if a.report:
                with open(a.report, "a", encoding="utf-8") as f:
                    f.write("\n".join(render(n, os.path.basename(path), ch, info, 10 ** 9, True)) + "\n")
        if a.log:
            lines, bad = check_log(a.log, cur.get("ожидания LISP (SX_log)", {}))
            full_report += lines
            if bad:
                rc = 1
    print("\n".join(full_report))
    print("Итог:", "ЕСТЬ ИЗМЕНЕНИЯ / ОШИБКИ (код 1)" if rc else "OK (код 0)")
    return rc


if __name__ == "__main__":
    sys.exit(main())
