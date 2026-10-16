# -*- coding: utf-8 -*-
"""Сервер окна: только 127.0.0.1 (этот ПК). Статика — ui/static, API — ui/routes.py."""
from __future__ import annotations

import json
import mimetypes
import os
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from . import routes

STATIC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
mimetypes.add_type("text/javascript", ".js")


def make_handler(app):
    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _send(self, code, body, ctype="application/json; charset=utf-8"):
            if isinstance(body, routes.File):
                from urllib.parse import quote
                self.send_response(code)
                self.send_header("Content-Type", body.ctype)
                self.send_header("Content-Disposition", f"attachment; filename=\"odnolineyka.xlsx\"; filename*=UTF-8''{quote(body.name)}")
                self.send_header("Content-Length", str(len(body.data)))
                self.end_headers()
                self.wfile.write(body.data)
                return
            if not isinstance(body, bytes):
                body = json.dumps(body, ensure_ascii=False).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _api(self, method):
            fn = routes.find(method, self.path.split("?")[0])
            if not fn:
                return self._send(404, {"ok": False, "error": "нет такого запроса"})
            body = {}
            if "?" in self.path:
                from urllib.parse import parse_qsl
                body = dict(parse_qsl(self.path.split("?", 1)[1]))
            if method == "POST":
                n = int(self.headers.get("Content-Length") or 0)
                body = json.loads(self.rfile.read(n) or b"{}")
            try:
                self._send(200, fn(app, body))
            except Exception as e:  # noqa: BLE001
                self._send(500, {"ok": False, "error": str(e)})

        def do_GET(self):
            if self.path.startswith("/api/"):
                return self._api("GET")
            rel = self.path.split("?")[0].lstrip("/") or "index.html"
            path = os.path.normpath(os.path.join(STATIC, rel))
            if not path.startswith(STATIC) or not os.path.isfile(path):
                return self._send(404, b"", "text/plain")
            with open(path, "rb") as f:
                ctype = (mimetypes.guess_type(path)[0] or "application/octet-stream") + "; charset=utf-8"
                self._send(200, f.read(), ctype)

        def do_POST(self):
            self._api("POST")

    return H


def serve(folder: str, port: int = 8765, open_browser=True):
    app = routes.App(folder)
    for p in range(port, port + 20):
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", p), make_handler(app))
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
