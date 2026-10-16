# -*- coding: utf-8 -*-
"""Отпечаток данных щита — короткий хэш того, что попадает в чертёж и журналы.

Зачем: SXDRAW / «КЖ в DWG» запишут в DWG отпечаток, по которому строили. Окно сравнит его с текущим:
не совпал — «схема устарела». Отпечаток не зависит от порядка чтения и от результатов «для себя»
(статусы, проверки) — только от того, что видно на схеме / в КЖ.
"""
from __future__ import annotations

import dataclasses
import hashlib
import json

from .model import Panel


def _line_key(ln):
    s = dataclasses.asdict(ln.scheme)
    return [ln.number, ln.name, s, ln.input.p_kw, ln.input.length, ln.result.s_mm2, ln.result.du]


def panel_revision(p: Panel) -> str:
    data = {
        "name": p.name,
        "supply": dataclasses.asdict(p.supply),
        "lines": sorted((_line_key(ln) for ln in p.lines), key=lambda x: str(x[0])),
    }
    raw = json.dumps(data, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]
