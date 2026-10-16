# -*- coding: utf-8 -*-
"""Замечания по объекту.

1. Из книг — то, что уже проверил Excel: статус строки «Однолинейки» (AE), категория, проверки питающей линии.
2. Свои, по всему объекту — то, чего одна книга не видит: связи щитов устарели, фидер мал, селективность,
   фидер без книги щита, щит без питания.
"""
from __future__ import annotations

from .model import Issue, Project

SEL_K = 1.6          # селективность: In фидера ≥ 1,6 · In отходящего (ориентир, как в книге)
TOL = 0.01           # 1 % — значения связи считаются одинаковыми


def _ok(t: str) -> bool:
    t = (t or "").strip().lower()
    return t == "" or t.startswith(("ok", "ок", "—", "-", "да", "обеспечена"))


def _diff(a, b) -> bool:
    if a is None or b is None:
        return False
    return abs(a - b) > TOL * max(abs(a), abs(b), 1e-9)


def _f(x, d=1):
    if x is None:
        return "—"
    s = f"{x:.{d}f}".rstrip("0").rstrip(".") if d else f"{x:.0f}"
    return s.replace(".", ",")


def run_checks(prj: Project):
    add = prj.issues.append
    panels = {p.id: p for p in prj.panels()}

    for b in prj.books:
        if b.error:
            add(Issue(2, f"книга не прочитана: {b.error}", source="объект", where=b.file))

    for p in panels.values():
        # ---- 1. из Excel
        no_kz = 0
        for ln in p.lines:
            if ln.level and ln.status:
                add(Issue(ln.level, ln.status, p.id, ln.number, "Excel", f"Однолинейка!AE{ln.row_odn}"))
            if ln.category_check and not _ok(ln.category_check):
                add(Issue(2, f"категория {ln.category}: {ln.category_check}", p.id, ln.number, "Excel",
                          f"Исходные данные!BV{ln.row_src}"))
            if ln.ikz_check.startswith("не проверено"):
                no_kz += 1
        for n in p.row_notes:
            add(Issue(n["level"], n["status"], p.id, n["qf"], "Excel", f"Однолинейка!AE{n['row']}"))
        if no_kz:
            add(Issue(1, f"КЗ не проверен на {no_kz} линиях: нет данных питания щита", p.id, source="Excel",
                      where="Сеть_щита!C12"))
        for k, v in p.checks.items():
            if not _ok(v):
                lvl = 2 if k in ("Аппарат и кабель", "Отключение КЗ") else 1
                add(Issue(lvl, f"питающая линия, {k.lower()}: {v}", p.id, source="Excel", where="Сеть_щита!C38"))

        # ---- 2. объект: питание задано?
        if not p.parent and not p.supply.from_file and not p.supply.manual:
            add(Issue(1, "не задано, откуда питается щит («Связать с вышестоящим» или «Сеть_щита» раздел 2)",
                      p.id, source="объект", where="Сеть_щита!C12"))

    # ---- 2. объект: фидеры и связи
    fed_by = {}
    for child in panels.values():
        if not child.parent:
            continue
        par = panels[child.parent]
        if not child.parent_line:
            add(Issue(1, f"не найдена линия-фидер в щите {par.name}", child.id, source="объект",
                      where="Сеть_щита!G6"))
            continue
        fed_by.setdefault((par.id, child.parent_line), []).append(child)
        ln = next((x for x in par.lines if x.number == child.parent_line), None)
        if ln is None:
            add(Issue(2, f"в щите {par.name} нет линии {child.parent_line} (фидер удалён?)", child.id,
                      source="объект", where="Сеть_щита!G6"))
            continue
        where_par = f"Исходные данные!D{ln.row_src}"
        # то, что знает нижестоящий о фидере, против книги вышестоящего
        if _diff(ln.in_a, child.supply.in_a):
            add(Issue(2, f"связь устарела: In фидера у {par.name} {_f(ln.in_a)} А, в книге щита {_f(child.supply.in_a)} А "
                         f"— «Обновить связи»", child.id, ln.number, "объект", "Сеть_щита!F6"))
        if _diff(ln.length, child.supply.length):
            add(Issue(1, f"связь устарела: длина фидера у {par.name} {_f(ln.length)} м, в книге щита "
                         f"{_f(child.supply.length)} м", child.id, ln.number, "объект", "Сеть_щита!K6"))
        # то, что знает вышестоящий о нагрузке нижестоящего
        seen = next((e for e in par.feeders_seen if e["line"] == ln.number), None)
        if seen and _diff(seen["p_inst"], child.p_inst):
            add(Issue(2, f"нагрузка {child.name} в книге {par.name} устарела: Руст {_f(seen['p_inst'])} кВт, "
                         f"сейчас {_f(child.p_inst)} кВт — «Обновить связи» в книге {par.name}",
                      par.id, ln.number, "объект", f"Сеть_щита!D{seen['row']}"))
        # фидер против нагрузки
        if ln.in_a is not None and child.i_calc is not None and ln.in_a < child.i_calc:
            add(Issue(2, f"фидер мал: In {_f(ln.in_a)} А < Iр щита {child.name} {_f(child.i_calc)} А",
                      par.id, ln.number, "объект", where_par))
        outs = [x.in_a for x in child.lines if x.in_a and x.group_ks != "Ввод"]
        if ln.in_a and outs and ln.in_a < SEL_K * max(outs):
            add(Issue(1, f"селективность: In фидера {_f(ln.in_a)} А < 1,6 × {_f(max(outs))} А "
                         f"(наибольший отходящий в {child.name})", par.id, ln.number, "объект", where_par))
    for (pid, line), kids in fed_by.items():
        if len(kids) > 1:
            add(Issue(2, f"линия {line} указана фидером у нескольких щитов: " + ", ".join(k.name for k in kids),
                      pid, line, "объект"))

    # фидеры («Щит» в группе Кс), у которых нет книги нижестоящего
    for p in panels.values():
        for ln in p.lines:
            if ln.group_ks == "Щит" and not ln.feeds:
                add(Issue(1, f"фидер «{ln.name}»: книги щита нет в папке объекта (или не связана)",
                          p.id, ln.number, "объект", f"Исходные данные!D{ln.row_src}"))

    prj.issues.sort(key=lambda i: (-i.level, i.panel, i.line))
