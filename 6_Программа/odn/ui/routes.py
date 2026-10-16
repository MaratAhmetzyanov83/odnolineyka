# -*- coding: utf-8 -*-
"""API окна. Новый запрос — функция с @route; получает App и тело запроса (dict), возвращает dict / bytes."""
from __future__ import annotations

import json
import threading

from .. import checks
from ..core import serialize
from ..core.model import Ref
from ..services.project import load_project, open_at

_ROUTES = {}


def route(method: str, path: str):
    def deco(fn):
        _ROUTES[(method, path)] = fn
        return fn
    return deco


def find(method, path):
    return _ROUTES.get((method, path))


class App:
    """Состояние окна: папка, загруженный объект. Один объект на окно."""

    def __init__(self, folder):
        self.folder = folder
        self.lock = threading.Lock()
        self.prj = None
        self.payload = b""
        self.reload()

    def reload(self):
        with self.lock:
            self.prj = load_project(self.folder)
            data = serialize.to_dict(self.prj)
            data["rules"] = [{"id": r.id, "title": r.title, "group": r.group} for r in checks.rules()]
            self.payload = json.dumps(data, ensure_ascii=False).encode("utf-8")


@route("GET", "/api/project")
def get_project(app: App, body):
    return app.payload


@route("POST", "/api/reload")
def post_reload(app: App, body):
    app.reload()
    return {"ok": True}


@route("POST", "/api/open")
def post_open(app: App, body):
    ref = serialize.from_dict(Ref, body.get("ref") or {})
    return {"ok": True, "msg": open_at(app.prj, ref)}
