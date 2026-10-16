# -*- coding: utf-8 -*-
"""Окно программы: маленький веб-сервер на этом ПК (127.0.0.1), страница открывается в браузере.

GET  /                 — страница
GET  /api/project      — объект (JSON)
POST /api/reload       — перечитать изменившиеся книги
POST /api/open         — открыть книгу в Excel на нужной ячейке {file, where}
"""
from __future__ import annotations

import json
import os
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from .project import load_project

WEB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "web")


class State:
    def __init__(self, folder):
        self.folder = folder
        self.lock = threading.Lock()
        self.data = None
        self.reload()

    def reload(self):
        with self.lock:
            prj = load_project(self.folder)
            self.data = json.dumps(prj.to_dict(), ensure_ascii=False).encode("utf-8")


def open_in_excel(path: str, where: str):
    """Книга — в Excel пользователя (видимый), лист и ячейка — выделить. Без pywin32 — просто открыть файл."""
    sheet, _, addr = where.partition("!")
    try:
        import win32com.client  # type: ignore
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True
        wb = None
        for w in xl.Workbooks:
            if os.path.normcase(w.FullName) == os.path.normcase(path):
                wb = w
                break
        if wb is None:
            wb = xl.Workbooks.Open(path)
        wb.Activate()
        if sheet and addr:
            ws = wb.Worksheets(sheet)
            ws.Activate()
            xl.Goto(ws.Range(addr), True)
        return "ok"
    except Exception as e:  # noqa: BLE001
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
            return f"открыт файл (ячейку найдите сами: {where}); {e}"
        return f"не открыт: {e}"


def make_handler(state: State):
    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _send(self, code, body: bytes, ctype="application/json; charset=utf-8"):
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _json(self, obj, code=200):
            self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"))

        def do_GET(self):
            if self.path in ("/", "/index.html"):
                with open(os.path.join(WEB, "index.html"), "rb") as f:
                    self._send(200, f.read(), "text/html; charset=utf-8")
            elif self.path == "/api/project":
                self._send(200, state.data)
            else:
                self._send(404, b"{}")

        def do_POST(self):
            n = int(self.headers.get("Content-Length") or 0)
            body = json.loads(self.rfile.read(n) or b"{}")
            if self.path == "/api/reload":
                try:
                    state.reload()
                    self._json({"ok": True})
                except Exception as e:  # noqa: BLE001
                    self._json({"ok": False, "error": str(e)}, 500)
            elif self.path == "/api/open":
                rel = body.get("file", "")
                path = os.path.abspath(os.path.join(state.folder, rel))
                if not path.startswith(os.path.abspath(state.folder)) or not os.path.isfile(path):
                    self._json({"ok": False, "error": "нет файла"}, 400)
                    return
                self._json({"ok": True, "msg": open_in_excel(path, body.get("where", ""))})
            else:
                self._json({"ok": False}, 404)

    return H


def serve(folder: str, port: int = 8765, open_browser=True):
    state = State(folder)
    for p in range(port, port + 20):
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", p), make_handler(state))
            break
        except OSError:
            continue
    else:
        raise SystemExit("нет свободного порта")
    url = f"http://127.0.0.1:{srv.server_address[1]}/"
    print(f"Окно объекта: {url}  (закрыть — Ctrl+C или закрыть это окно)")
    if open_browser:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
