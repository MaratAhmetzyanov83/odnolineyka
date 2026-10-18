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
from ..services import edit, export, fixes, link_writer, structure
from ..services.project import add_exclusion, load_project, open_at, recheck

_ROUTES = {}


def route(method: str, path: str):
    def deco(fn):
        _ROUTES[(method, path)] = fn
        return fn
    return deco


def find(method, path):
    return _ROUTES.get((method, path))


def _payload(prj):
    data = serialize.to_dict(prj)
    for d in data["documents"]:
        d.pop("cable_journal", None)                # большие строки КЖ окну не нужны
    return data


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
            data = _payload(self.prj)
            data["rules"] = [{"id": r.id, "title": r.title, "group": r.group} for r in checks.rules()]
            self.payload = json.dumps(data, ensure_ascii=False).encode("utf-8")

    def refresh(self):
        """Без чтения книг: структура, связи, расчёт, проверки — и новый JSON для окна."""
        with self.lock:
            recheck(self.prj)
            data = _payload(self.prj)
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


@route("GET", "/api/fixes")
def get_fixes(app: App, body):
    """«Исправить ошибки»: предложения правок (ничего не пишет)."""
    return {"items": fixes.propose(app.prj)}


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


@route("POST", "/api/sync")
def post_sync(app: App, body):
    """«Синхронизировать книги»: все связанные книги — пересчёт вместе и сохранение."""
    with app.write_lock:
        res = edit.sync_books(app.prj)
        if res.get("ok"):
            app.reload()
        return res


@route("POST", "/api/undo")
def post_undo(app: App, body):
    with app.write_lock:
        try:
            res = edit.undo_last(app.prj)
        except EditError as e:
            app.reload_quiet()
            return {"ok": False, "errors": [str(e)]}
        return _after_write(app, res)


@route("POST", "/api/exclude")
def post_exclude(app: App, body):
    rel = body.get("file", "")
    if not rel or not any(d.file == rel for d in app.prj.documents):
        return {"ok": False, "errors": ["нет такой книги в объекте"]}
    with app.write_lock:
        add_exclusion(app.folder, rel)
        app.reload_quiet()
    return {"ok": True}


@route("GET", "/api/link/suggest")
def get_link_suggest(app: App, body):
    """Линия-фидер у выбранного вышестоящего щита для щита child — по имени щита."""
    from ..services.linking import suggest_feeder
    par, child = app.prj.panel(body.get("parent", "")), app.prj.panel(body.get("child", ""))
    return {"ok": True, "line": suggest_feeder(par, child) if par and child else ""}


@route("POST", "/api/link/write")
def post_link_write(app: App, body):
    """Вариант Б: связь в книги (функция надстройки). Удачно — связь программы снимается: теперь она в книге."""
    with app.write_lock:
        try:
            res = link_writer.write_link(app.prj, body.get("panel", ""), body.get("parent", ""),
                                         str(body.get("line", "")).strip(), int(body.get("input") or 1))
        except EditError as e:
            return {"ok": False, "errors": [str(e)]}
        except Exception as e:  # noqa: BLE001 — Excel / COM
            return {"ok": False, "errors": [f"Excel: {e}"]}
        try:
            structure.operate(app.prj, {"op": "clear_link", "panel": body.get("panel", "")})
        except EditError:
            pass
        app.reload()
        return res


@route("POST", "/api/link/undo")
def post_link_undo(app: App, body):
    with app.write_lock:
        res = link_writer.undo_last(app.prj)
        if res.get("ok"):
            app.reload()
        return res


@route("POST", "/api/structure")
def post_structure(app: App, body):
    with app.write_lock:
        try:
            res = structure.operate(app.prj, body)
        except EditError as e:
            return {"ok": False, "errors": [str(e)]}
        app.refresh()
        return res


@route("POST", "/api/structure/undo")
def post_structure_undo(app: App, body):
    with app.write_lock:
        res = structure.undo_last(app.prj)
        app.refresh()
        return res


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
