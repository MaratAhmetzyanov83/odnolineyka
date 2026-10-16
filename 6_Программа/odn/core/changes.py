# -*- coding: utf-8 -*-
"""Правки — команды. Окно не меняет модель напрямую: оно создаёт Change, сервис проверяет и применяет.

Change — «у линии L щита P поле input.length_plan: было 30, стало 45».
Из списка правок строится:
  - запись в источник (адаптер документа: sources.<kind>.writer) с резервной копией;
  - журнал (кто, когда, что) и отмена (обратная правка).
Шаг 2 плана. Сейчас: класс, проверка пути к полю и применение к модели в памяти (для расчёта «что если»).
"""
from __future__ import annotations

import dataclasses
import time
from dataclasses import dataclass, field
from typing import Any

from .model import Project

# что можно править из окна: часть линии / щита -> разрешённые поля
EDITABLE = {
    "line": {"name", "input.p_kw", "input.length_plan", "input.cable_code", "input.cores", "input.s_min",
             "input.in_manual", "input.parallel", "scheme.breaker", "scheme.bus", "scheme.place"},
    "panel": {"supply.from_doc", "supply.manual"},
}


@dataclass
class Change:
    panel: str                    # id щита
    path: str                     # «input.length_plan», «scheme.breaker», «supply.manual»…
    new: Any
    line: str = ""                # номер линии; пусто — поле щита
    old: Any = None               # заполняется при применении
    at: str = field(default_factory=lambda: time.strftime("%Y-%m-%d %H:%M:%S"))

    def inverse(self) -> "Change":
        return Change(self.panel, self.path, self.old, self.line, self.new)


def _target(prj: Project, ch: Change):
    if ch.line:
        obj = prj.line(ch.panel, ch.line)
        kind = "line"
    else:
        obj = prj.panel(ch.panel)
        kind = "panel"
    if obj is None:
        raise KeyError(f"нет {'линии ' + ch.line if ch.line else 'щита'} в {ch.panel}")
    if ch.path not in EDITABLE[kind]:
        raise PermissionError(f"поле {ch.path} не правится из окна")
    *head, last = ch.path.split(".")
    for h in head:
        obj = getattr(obj, h)
    if not dataclasses.is_dataclass(obj) or last not in {f.name for f in dataclasses.fields(obj)}:
        raise KeyError(f"нет поля {ch.path}")
    return obj, last


def apply(prj: Project, ch: Change) -> Change:
    """Применить к модели в памяти; вернуть правку с заполненным old."""
    obj, attr = _target(prj, ch)
    ch.old = getattr(obj, attr)
    setattr(obj, attr, ch.new)
    return ch
