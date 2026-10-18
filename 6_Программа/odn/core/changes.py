# -*- coding: utf-8 -*-
"""Правки — команды. Окно не меняет модель напрямую: оно присылает Change, сервис проверяет, пишет в источник,
ведёт журнал и умеет отменить (обратная правка).

Change — «у линии L щита P поле input.length_plan: было 30, стало 45».
Какие поля правятся и какого они типа — FIELDS (одно место). Адрес ячейки — в адаптере источника (layout.py).
"""
from __future__ import annotations

import dataclasses
import time
from dataclasses import dataclass, field
from typing import Any

from .model import Project

IN_SERIES = [6, 10, 13, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630]
S_SERIES = [0.5, 0.75, 1.5, 2.5, 4, 6, 10, 16, 25, 35, 50, 70, 95, 120, 150, 185, 240]

# путь поля -> что это. kind: line / panel; type: num / int / text; options: стандартные значения для списка
FIELDS = {
    "name":              dict(kind="line", type="text", title="Наименование (ручное)"),
    "input.p_kw":        dict(kind="line", type="num", title="Pу, кВт", min=0),
    "input.length_plan": dict(kind="line", type="num", title="Lпл, м", min=0),
    "input.cable_code":  dict(kind="line", type="text", title="Марка (код)"),
    "input.cores":       dict(kind="line", type="int", title="Жил", options=[2, 3, 4, 5], min=1),
    "input.s_min":       dict(kind="line", type="num", title="Sмин, мм²", options=S_SERIES, min=0),
    "input.in_manual":   dict(kind="line", type="num", title="In ручной, А", options=IN_SERIES, min=0),
    "input.parallel":    dict(kind="line", type="int", title="Кабелей параллельно", options=[1, 2, 3, 4], min=1),
    "scheme.breaker_set": dict(kind="line", type="text", title="Автомат («авто» или текст)", options=["авто"],
                               when="group_head"),
    "scheme.bus":        dict(kind="line", type="text", title="Шина"),
    "supply.manual":     dict(kind="panel", type="text", title="Питается от (вручную)"),
}


class EditError(ValueError):
    pass


@dataclass
class Change:
    panel: str                    # id щита
    path: str                     # путь поля (FIELDS)
    new: Any
    line: str = ""                # номер линии; пусто — поле щита
    old: Any = None               # заполняется при применении
    at: str = field(default_factory=lambda: time.strftime("%Y-%m-%d %H:%M:%S"))

    def inverse(self) -> "Change":
        return Change(self.panel, self.path, self.old, self.line, self.new)


def convert(path: str, value):
    """Значение из окна -> значение поля. Пусто -> None (очистить ячейку)."""
    spec = FIELDS.get(path)
    if spec is None:
        raise EditError(f"поле {path} не правится из окна")
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return None if spec["type"] != "text" else ""
    if spec["type"] == "text":
        return str(value).strip()
    try:
        x = float(str(value).replace(",", ".").replace(" ", ""))
    except ValueError:
        raise EditError(f"{spec['title']}: «{value}» — не число") from None
    if "min" in spec and x < spec["min"]:
        raise EditError(f"{spec['title']}: не меньше {spec['min']}")
    if spec["type"] == "int":
        if not x.is_integer():
            raise EditError(f"{spec['title']}: нужно целое")
        return int(x)
    return x


def target(prj: Project, ch: Change):
    """(объект, имя атрибута) для правки; проверка, что поле правится."""
    spec = FIELDS.get(ch.path)
    if spec is None:
        raise EditError(f"поле {ch.path} не правится из окна")
    if spec["kind"] == "line":
        if not ch.line:
            raise EditError(f"{spec['title']}: нужна линия")
        obj = prj.line(ch.panel, ch.line)
        if obj is None:
            raise EditError(f"нет линии {ch.line} в {ch.panel}")
        if spec.get("when") == "group_head" and obj.scheme.qf and not obj.scheme.group_head:
            raise EditError(f"линия {ch.line}: автомат задаётся в первой строке группы {obj.scheme.qf}")
    else:
        obj = prj.panel(ch.panel)
        if obj is None:
            raise EditError(f"нет щита {ch.panel}")
    *head, last = ch.path.split(".")
    for h in head:
        obj = getattr(obj, h)
    if not dataclasses.is_dataclass(obj) or last not in {f.name for f in dataclasses.fields(obj)}:
        raise EditError(f"нет поля {ch.path}")
    return obj, last


def apply(prj: Project, ch: Change) -> Change:
    """Проверить, привести тип, применить к модели в памяти; вернуть правку с заполненным old."""
    ch.new = convert(ch.path, ch.new)
    obj, attr = target(prj, ch)
    ch.old = getattr(obj, attr)
    setattr(obj, attr, ch.new)
    return ch
