# -*- coding: utf-8 -*-
"""Реестр проверок. Проверка = функция с декоратором @rule; она получает объект и выдаёт замечания.

    @rule("links.feeder_small", "Фидер меньше нагрузки щита", group="Связи щитов")
    def feeder_small(prj):
        for child in prj.panels():
            ...
            yield Hit(2, "фидер мал: …", panel=par.id, line=ln.number, ref=ln.refs.get("input"))

id правила — «группа_файла.имя»: по нему окно фильтрует, а пользователь сможет отключать правила.
Правило не меняет модель. Падение одного правила не роняет остальные (замечание «проверка … не выполнена»).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Optional

from ..core.model import Issue, Project, Ref
from .hints import hint_for


@dataclass
class Hit:
    level: int
    text: str
    panel: str = ""
    line: str = ""
    ref: Optional[Ref] = None


@dataclass
class Rule:
    id: str
    title: str
    group: str
    fn: Callable = field(repr=False)
    doc: str = ""


_RULES: dict = {}


def rule(rule_id: str, title: str, group: str):
    def deco(fn):
        if rule_id in _RULES:
            raise ValueError(f"правило {rule_id} уже есть")
        _RULES[rule_id] = Rule(rule_id, title, group, fn, (fn.__doc__ or "").strip())
        return fn
    return deco


def rules():
    return list(_RULES.values())


def run(prj: Project, disabled=()) -> list:
    out = []
    for r in _RULES.values():
        if r.id in disabled:
            continue
        try:
            for h in r.fn(prj) or ():
                out.append(Issue(r.id, h.level, h.text, h.panel, h.line, h.ref or Ref(), r.group, hint_for(h.text, r.doc)))
        except Exception as e:  # noqa: BLE001 — ошибка в правиле не должна ронять окно
            out.append(Issue(r.id, 1, f"проверка «{r.title}» не выполнена: {e}", group=r.group))
    out.sort(key=lambda i: (-i.level, i.panel, i.line))
    return out
