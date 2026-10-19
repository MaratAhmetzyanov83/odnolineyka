# -*- coding: utf-8 -*-
"""API окна. Новый запрос — функция с @route; получает App и тело запроса (dict), возвращает dict / bytes."""
from __future__ import annotations

import json
import threading
import time
import uuid

from .. import checks
from ..core import serialize
from ..core.model import Ref
from ..core.changes import EditError
from ..services import book_patch, dismiss, edit, export, fixes, kj_writer, link_writer, structure
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
        self.bg = {}                      # фон: пересчёт связанных книг после записи (окно спрашивает /api/bg)
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


@route("POST", "/api/book/patch")
def post_book_patch(app: App, body):
    """Доработка книги скриптом 5_Проверка (например, слаботочные)."""
    with app.write_lock:
        try:
            res = book_patch.run_patch(app.prj, body.get("doc", ""), body.get("name", ""))
        except EditError as e:
            return {"ok": False, "errors": [str(e)]}
        app.reload()
        return res


@route("GET", "/api/kj/cables")
def get_kj_cables(app: App, body):
    """Кабели листа «Кабельный_журнал» книги и убранные из КЖ."""
    try:
        return kj_writer.cables(app.prj, body.get("doc", ""))
    except EditError as e:
        return {"ok": False, "errors": [str(e)]}


@route("POST", "/api/kj/skip")
def post_kj_skip(app: App, body):
    """Убрать кабели из КЖ / вернуть: <книга>.kj.skip.txt (книга не меняется)."""
    with app.write_lock:
        try:
            res = kj_writer.set_skip(app.prj, body.get("doc", ""), body.get("skip") or [])
        except EditError as e:
            return {"ok": False, "errors": [str(e)]}
        app.refresh()                     # «КЖ: актуален / устарел» — без убранных кабелей
        return res


@route("POST", "/api/kj")
def post_kj(app: App, body):
    """КЖ в DWG: выпуск / обновление (функция надстройки SX_KJTo). items: [{doc, dwg}], dwg "" — из отметки."""
    with app.write_lock:
        try:
            res = kj_writer.issue(app.prj, body.get("items") or [])
        except EditError as e:
            return {"ok": False, "errors": [str(e)]}
        except Exception as e:  # noqa: BLE001 — Excel / AutoCAD / COM
            return {"ok": False, "errors": [f"Excel: {e}"]}
        if res.get("done"):
            app.reload_quiet()            # отметки .kj.txt — новое состояние «КЖ»
        return res


@route("POST", "/api/dismiss")
def post_dismiss(app: App, body):
    """Снять замечания вручную / вернуть (restore). items: [{rule, panel, line, text}], note — причина."""
    with app.write_lock:
        try:
            if body.get("restore"):
                res = dismiss.restore(app.prj, body.get("items") or [])
            else:
                res = dismiss.dismiss(app.prj, body.get("items") or [], body.get("note", ""))
        except EditError as e:
            return {"ok": False, "errors": [str(e)]}
        app.refresh()                     # проверки заново: снятые уходят в prj.dismissed
        return res


@route("GET", "/api/fixes")
def get_fixes(app: App, body):
    """«Исправить ошибки»: предложения правок (ничего не пишет)."""
    return {"items": fixes.propose(app.prj)}


@route("GET", "/api/journal")
def get_journal(app: App, body):
    return {"ok": True, "rows": edit.read_journal(app.prj)[::-1]}


BUSY = "идёт пересчёт связанных книг после прошлой записи — подождите, пока он закончится"


def _busy(app: App):
    return {"ok": False, "errors": [BUSY + f" ({app.bg.get('done', 0)} из {app.bg.get('total', 0)})"]}         if app.bg.get("running") else None


def _start_bg(app: App, res: dict):
    """Связанные книги — в фоне. Замок записи переходит фону: следующая запись ждёт его конца (окно — сообщение)."""
    related, edited = res.get("related") or [], res.get("docs") or []
    if not related:
        app.write_lock.release()
        return
    prj, backups = app.prj, res.get("backups") or {}
    app.bg = {"running": True, "done": 0, "total": len(related), "docs": related, "errors": [], "id": uuid.uuid4().hex[:8]}

    def step(n):
        app.bg["done"] = min(app.bg["done"] + n, app.bg["total"])

    def work():
        try:
            r = edit.refresh_related(prj, edited, related, backups, done=step)
            app.bg.update(errors=r["errors"], timing=r["timing"], again=r["again"])
        except Exception as e:  # noqa: BLE001
            app.bg["errors"] = [f"связанные книги не пересчитаны: {e} — «Синхронизировать книги»"]
        finally:
            try:
                app.reload_quiet()
            finally:
                app.bg["running"] = False
                app.write_lock.release()
    threading.Thread(target=work, daemon=True).start()


@route("POST", "/api/spec/ok")
def post_spec_ok(app: App, body):
    """«Спецификация верна»: запомнить аппараты и лист книги (services.spec_stamp)."""
    from ..services import spec_stamp
    res = spec_stamp.accept(app.prj, body.get("doc", ""))
    if res.get("ok"):
        app.refresh()
    return res


@route("GET", "/api/bg")
def get_bg(app: App, body):
    return {"ok": True, **app.bg}


def _write(app: App, fn):
    """Правка / отмена: изменённые книги — сразу, связанные — в фоне."""
    busy = _busy(app)
    if busy:
        return busy
    app.write_lock.acquire()
    try:
        res = fn()
    except EditError as e:
        app.reload_quiet()
        app.write_lock.release()
        return {"ok": False, "errors": [str(e)]}
    except BaseException:
        app.write_lock.release()
        raise
    try:
        res = _after_write(app, res)
        res.pop("backups", None)
        res["bg"] = bool(res.get("docs") and res.get("related"))
    finally:
        if res.get("bg"):
            _start_bg(app, res)
        else:
            app.write_lock.release()
    return res


def _after_write(app: App, res: dict):
    if res.get("docs"):
        t = time.time()
        app.reload()                      # перечитать изменённые книги (кэш: остальные — из кэша)
        res.setdefault("timing", {})["чтение"] = round(time.time() - t, 1)
    else:
        app.reload_quiet()                # модель в памяти могла поменяться — вернуть как в файлах
    return res


@route("POST", "/api/edit")
def post_edit(app: App, body):
    return _write(app, lambda: edit.apply_changes(app.prj, body.get("changes") or []))


@route("POST", "/api/sync")
def post_sync(app: App, body):
    """«Синхронизировать книги»: все связанные книги — пересчёт вместе и сохранение."""
    if _busy(app):
        return _busy(app)
    with app.write_lock:
        res = edit.sync_books(app.prj)
        if res.get("ok"):
            app.reload()
        return res


@route("POST", "/api/undo")
def post_undo(app: App, body):
    return _write(app, lambda: edit.undo_last(app.prj))


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
