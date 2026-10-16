# -*- coding: utf-8 -*-
"""API окна. Новый запрос — функция с @route; получает App и тело запроса (dict), возвращает dict / bytes."""
from __future__ import annotations

import json
import threading
import uuid

from .. import checks
from ..core import serialize
from ..core.model import Ref
from ..core.changes import EditError
from ..services import edit, export
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
        self.write_lock = threading.Lock()
        self.prj = None
        self.payload = b""
        self.reload()

    def reload(self, progress=print):
        with self.lock:
            self.prj = load_project(self.folder, progress=progress)
            data = serialize.to_dict(self.prj)
            data["rules"] = [{"id": r.id, "title": r.title, "group": r.group} for r in checks.rules()]
            self.payload = json.dumps(data, ensure_ascii=False).encode("utf-8")

    def reload_quiet(self):
        self.reload(progress=lambda s: None)


@route("GET", "/api/project")
def get_project(app: App, body):
    return app.payload


@route("POST", "/api/reload")
def post_reload(app: App, body):
    app.reload()
    return {"ok": True}


class File:
    """Ответ-файл (скачивание)."""
    def __init__(self, data: bytes, name: str, ctype: str):
        self.data, self.name, self.ctype = data, name, ctype


@route("GET", "/api/options")
def get_options(app: App, body):
    return edit.options(app.prj)


@route("GET", "/api/journal")
def get_journal(app: App, body):
    return {"ok": True, "rows": edit.read_journal(app.prj)[::-1]}


def _after_write(app: App, res: dict):
    if res.get("docs"):
        app.reload()                      # перечитать изменённые книги (кэш: остальные — из кэша)
    else:
        app.reload_quiet()                # модель в памяти могла поменяться — вернуть как в файлах
    return res


@route("POST", "/api/edit")
def post_edit(app: App, body):
    with app.write_lock:
        try:
            res = edit.apply_changes(app.prj, body.get("changes") or [])
        except EditError as e:
            app.reload_quiet()
            return {"ok": False, "errors": [str(e)]}
        return _after_write(app, res)


@route("POST", "/api/undo")
def post_undo(app: App, body):
    with app.write_lock:
        try:
            res = edit.undo_last(app.prj)
        except EditError as e:
            app.reload_quiet()
            return {"ok": False, "errors": [str(e)]}
        return _after_write(app, res)


@route("POST", "/api/export")
def post_export(app: App, body):
    name = (body.get("title") or "Таблица").replace("/", "-").replace("\\", "-")
    data = export.to_xlsx(body.get("title", ""), body.get("columns") or [], body.get("rows") or [])
    token = uuid.uuid4().hex
    app.files = {token: File(data, name + ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    return {"ok": True, "url": f"/api/file?t={token}"}


@route("GET", "/api/file")
def get_file(app: App, body):
    f = getattr(app, "files", {}).pop(body.get("t", ""), None)
    if f is None:
        return {"ok": False, "error": "файл уже скачан или устарел"}
    return f


@route("POST", "/api/open")
def post_open(app: App, body):
    ref = serialize.from_dict(Ref, body.get("ref") or {})
    return {"ok": True, "msg": open_at(app.prj, ref)}
