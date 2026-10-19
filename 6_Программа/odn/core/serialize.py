# -*- coding: utf-8 -*-
"""JSON <-> модель. from_dict терпит лишние и недостающие поля (старый кэш, новая версия модели)."""
from __future__ import annotations

import dataclasses
import typing
from functools import lru_cache

from . import model

# поля-списки: какой класс внутри
LIST_ITEMS = {
    (model.Project, "documents"): model.Document,
    (model.Project, "issues"): model.Issue,
    (model.Project, "dismissed"): model.Issue,
    (model.Document, "panels"): model.Panel,
    (model.Panel, "lines"): model.Line,
    (model.Panel, "feeders_seen"): model.FeederSeen,
    (model.Panel, "notes"): model.Note,
    (model.Panel, "nodes"): model.Node,
    (model.Panel, "inputs_more"): model.Supply,
}
# поля-словари: какой класс у значений
DICT_ITEMS = {
    (model.Line, "refs"): model.Ref,
    (model.Panel, "refs"): model.Ref,
}


def to_dict(obj):
    return dataclasses.asdict(obj)


@lru_cache(maxsize=None)
def _hints(cls):
    return typing.get_type_hints(cls, vars(model))


def from_dict(cls, d):
    if d is None:
        return None
    hints = _hints(cls)
    kw = {}
    for f in dataclasses.fields(cls):
        if f.name not in d:
            continue
        v = d[f.name]
        t = hints.get(f.name)
        item = LIST_ITEMS.get((cls, f.name))
        ditem = DICT_ITEMS.get((cls, f.name))
        if item and isinstance(v, list):
            v = [from_dict(item, x) for x in v]
        elif ditem and isinstance(v, dict):
            v = {k: from_dict(ditem, x) for k, x in v.items()}
        elif dataclasses.is_dataclass(t) and isinstance(v, dict):
            v = from_dict(t, v)
        kw[f.name] = v
    return cls(**kw)
