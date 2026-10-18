# -*- coding: utf-8 -*-
"""Связи щитов: то, чего не видит одна книга. Работают по p.parent / p.parent_line (services.linking)."""
from .registry import Hit, rule
from .util import differ, f

G = "Связи щитов"
SEL_K = 1.6          # селективность: In фидера ≥ 1,6 · In отходящего (ориентир, как в книге)


def _is_input(ln):
    return ln.number.strip().lower().startswith("ввод")


def _feeders(prj):
    """(нижестоящий, вышестоящий, линия-фидер) для связанных щитов."""
    for child in prj.panels():
        if child.parent and child.parent_line:
            par = prj.panel(child.parent)
            ln = prj.line(child.parent, child.parent_line)
            if par and ln:
                yield child, par, ln


@rule("links.parent_missing", "Вышестоящий щит не найден", G)
def parent_missing(prj):
    for p in prj.panels():
        if p.link_error:
            yield Hit(2, p.link_error, p.id, ref=p.refs.get("supply"))


@rule("links.no_supply", "Не задано питание щита", G)
def no_supply(prj):
    for p in prj.panels():
        if not p.parent and not p.supply.from_doc and not p.supply.manual and not p.link_error:
            yield Hit(1, "не задано, откуда питается щит («Связать с вышестоящим» или «Сеть_щита» раздел 2)",
                      p.id, ref=p.refs.get("supply_manual") or p.refs.get("supply"))


@rule("links.no_feeder_line", "Нет линии-фидера у вышестоящего", G)
def no_feeder_line(prj):
    for p in prj.panels():
        par = prj.panel(p.parent) if p.parent else None
        if not par:
            continue
        if not p.parent_line:
            yield Hit(1, f"не найдена линия-фидер в щите {par.name}", p.id, ref=p.refs.get("supply"))
        elif prj.line(par.id, p.parent_line) is None:
            yield Hit(2, f"в щите {par.name} нет линии {p.parent_line} (фидер удалён?)", p.id,
                      ref=p.refs.get("supply"))


@rule("links.feeder_stale", "Нижестоящий знает старый фидер", G)
def feeder_stale(prj):
    for child, par, ln in _feeders(prj):
        if differ(ln.in_a, child.supply.in_a):
            yield Hit(2, f"связь устарела: In фидера {ln.number} у {par.name} {f(ln.in_a)} А, в книге щита "
                         f"{f(child.supply.in_a)} А — «Обновить связи»", child.id, "", child.refs.get("supply"))
        if differ(ln.input.length, child.supply.length):
            yield Hit(1, f"связь устарела: длина фидера {ln.number} у {par.name} {f(ln.input.length)} м, в книге щита "
                         f"{f(child.supply.length)} м", child.id, "", child.refs.get("supply"))


@rule("links.load_stale", "Вышестоящий считает по старой нагрузке", G)
def load_stale(prj):
    for child, par, ln in _feeders(prj):
        seen = next((e for e in par.feeders_seen if e.line == ln.number), None)
        if seen and differ(seen.p_inst, child.result.p_inst):
            yield Hit(2, f"нагрузка {child.name} в книге {par.name} устарела: Руст {f(seen.p_inst)} кВт, "
                         f"сейчас {f(child.result.p_inst)} кВт — «Обновить связи» в книге {par.name}",
                      par.id, ln.number, seen.ref)


@rule("links.feeder_small", "Фидер меньше нагрузки щита", G)
def feeder_small(prj):
    for child, par, ln in _feeders(prj):
        if ln.in_a is not None and child.result.i_calc is not None and ln.in_a < child.result.i_calc:
            yield Hit(2, f"фидер мал: In {f(ln.in_a)} А < Iр щита {child.name} {f(child.result.i_calc)} А",
                      par.id, ln.number, ln.refs.get("scheme") or ln.refs.get("input"))


@rule("links.selectivity", "Селективность фидера", G)
def selectivity(prj):
    """In фидера ≥ 1,6 × наибольший отходящий автомат нижестоящего (ориентир; подтверждать по картам)."""
    for child, par, ln in _feeders(prj):
        inp = next((x for x in child.lines if _is_input(x) and x.in_a), None)
        outs = [x.in_a for x in child.lines if x.in_a and not _is_input(x)]
        base, by = (inp.in_a, f"ввод {inp.number} в {child.name}") if inp else \
                   ((max(outs), f"наибольший отходящий в {child.name}") if outs else (None, ""))
        if ln.in_a and base and ln.in_a < SEL_K * base:
            yield Hit(1, f"селективность: In фидера {f(ln.in_a)} А < 1,6 × {f(base)} А ({by}) — "
                         f"у щита {child.name} «Фидер ✎» → «Подобрать In»", par.id, ln.number,
                      ln.refs.get("scheme") or ln.refs.get("input"))


@rule("links.feeder_twice", "Линия-фидер у нескольких щитов", G)
def feeder_twice(prj):
    seen = {}
    for child, par, ln in _feeders(prj):
        seen.setdefault((par.id, ln.number), []).append(child)
    for (pid, line), kids in seen.items():
        if len(kids) > 1:
            yield Hit(2, f"линия {line} указана фидером у нескольких щитов: " + ", ".join(k.name for k in kids),
                      pid, line)


@rule("links.feeder_no_book", "Фидер без книги щита", G)
def feeder_no_book(prj):
    for p in prj.panels():
        for ln in p.lines:
            if ln.input.group_ks == "Щит" and not ln.feeds:
                yield Hit(1, f"фидер «{ln.name}»: книги щита нет в папке объекта (или не связана)",
                          p.id, ln.number, ln.refs.get("input"))


@rule("links.duplicate_panel", "Щит в нескольких книгах", G)
def duplicate_panel(prj):
    """Один щит в двух книгах — обычно резервная копия. «Не читать эту книгу» у лишней."""
    seen = {}
    for p in prj.panels():
        seen.setdefault(p.name.strip().lower(), []).append(p)
    for ps in seen.values():
        if len(ps) > 1:
            for p in ps:
                others = ", ".join(q.doc for q in ps if q is not p)
                yield Hit(1, f"щит {p.name} есть и в другой книге: {others} — лишнюю исключите («Не читать эту книгу»)", p.id)


@rule("links.mismatch", "Связь программы ≠ связь в книге", G)
def mismatch(prj):
    for p in prj.panels():
        if p.link_src != "программа":
            continue
        par = prj.panel(p.parent)
        if not p.book_parent:
            yield Hit(1, f"связь {par.name} / {p.parent_line or '?'} есть только в программе — в книге щита её нет "
                         f"(Excel не получает ΔU и Iкз от вышестоящего): «Связь…» → «Записать в книги»", p.id)
        elif (p.book_parent, p.book_line) != (p.parent, p.parent_line):
            bp = prj.panel(p.book_parent)
            yield Hit(1, f"в программе: {par.name} / {p.parent_line or '?'}, в книге: {bp.name if bp else '?'} / "
                         f"{p.book_line or '?'} — «Связь…» → «Записать в книги» или «Убрать связь программы»", p.id)


@rule("links.input_no_device", "Ввод без аппарата", G)
def input_no_device(prj):
    """Строка ввода без аппарата (E пусто): нет QS / автомата ввода. «авто» — Excel подберёт по фидеру."""
    for p in prj.panels():
        for ln in p.lines:
            if _is_input(ln) and not ln.scheme.breaker and not ln.scheme.breaker_set:
                yield Hit(1, f"ввод {ln.number} без аппарата — в «Автомат» впишите «авто» (QS по фидеру) или текст",
                          p.id, ln.number, ln.refs.get("scheme") or ln.refs.get("input"))
