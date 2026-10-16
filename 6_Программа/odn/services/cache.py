# -*- coding: utf-8 -*-
"""Кэш прочитанных документов: файл не менялся (mtime) — берём из кэша. Лежит в 6_Программа/кэш/ (не в Git)."""
from __future__ import annotations

import json
import os

from ..core import serialize
from ..core.model import SCHEMA, Document

BASE = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "кэш")


def _path(folder: str) -> str:
    key = "".join(ch if ch.isalnum() else "_" for ch in os.path.abspath(folder))[-120:]
    return os.path.join(BASE, key + ".json")


def load(folder: str) -> dict:
    """{файл: Document} — только совместимой версии."""
    try:
        with open(_path(folder), encoding="utf-8") as f:
            d = json.load(f)
        if d.get("schema") != SCHEMA:
            return {}
        return {x["file"]: serialize.from_dict(Document, x) for x in d.get("documents", [])}
    except (OSError, ValueError, KeyError, TypeError):
        return {}


def save(folder: str, docs) -> None:
    try:
        os.makedirs(BASE, exist_ok=True)
        with open(_path(folder), "w", encoding="utf-8") as f:
            json.dump({"schema": SCHEMA, "documents": [serialize.to_dict(d) for d in docs if not d.error]},
                      f, ensure_ascii=False)
    except OSError:
        pass
