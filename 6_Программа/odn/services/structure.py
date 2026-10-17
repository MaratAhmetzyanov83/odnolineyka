# -*- coding: utf-8 -*-
"""Структура щитов (каскад узлов) — данные самой программы, не Excel.

Файл: <папка объекта>/_окно/структура.json
    {"format": 1, "panels": {"<id щита>": {"nodes": [{id, kind, parent, bus, device, rcd_ma, note}],
                                            "lines": {"<номер линии>": "<id узла>"}}}}
Операции окна (op): add_node, update_node, delete_node, move_lines. Каждая — запись в журнал «_окно/журнал.jsonl»
(kind «структура», состояние щита до операции) — отмена восстанавливает его.
"""
from __future__ import annotations

import dataclasses
import json
import os
import time
import uuid

from ..core.changes import EditError
from ..core.model import Node, Project

FORMAT = 1
KINDS = ("УЗО", "АВДТ", "Автомат", "Рубильник")
NODE_INPUT = ("id", "kind", "parent", "bus", "device", "rcd_ma", "note")
SERVICE_DIR = "_окно"


def _path(folder):
    return os.path.join(folder, SERVICE_DIR, "структура.json")


def load(folder) -> dict:
    try:
        with open(_path(folder), encoding="utf-8") as f:
            d = json.load(f)
        return d if d.get("format") == FORMAT else {"format": FORMAT, "panels": {}}
    except (OSError, ValueError):
        return {"format": FORMAT, "panels": {}}


def save(folder, data):
    os.makedirs(os.path.join(folder, SERVICE_DIR), exist_ok=True)
    tmp = _path(folder) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, _path(folder))


def apply(prj: Project, data: dict | None = None):
    """Наложить структуру на прочитанный объект: panel.nodes, line.node."""
    data = data or load(prj.folder)
    for p in prj.panels():
        st = data["panels"].get(p.id, {})
        p.nodes = [Node(**{k: v for k, v in n.items() if k in NODE_INPUT}) for n in st.get("nodes", [])]
        where = st.get("lines", {})
        for ln in p.lines:
            ln.node = where.get(ln.number, "")


# ---------------------------------------------------------------- операции
def _panel_state(data, pid):
    return data["panels"].setdefault(pid, {"nodes": [], "lines": {}})


def _next_id(st):
    used = {n["id"] for n in st["nodes"]}
    i = 1
    while f"У{i}" in used:
        i += 1
    return f"У{i}"


def _check_parent(st, node_id, parent):
    """Родитель существует и не создаёт цикл."""
    if not parent:
        return
    ids = {n["id"]: n for n in st["nodes"]}
    if parent not in ids:
        raise EditError(f"нет узла {parent}")
    p = parent
    while p:
        if p == node_id:
            raise EditError(f"узел {node_id} не может стоять под собой (цикл)")
        p = ids[p].get("parent", "")


def _clean(fields):
    out = {}
    for k in ("kind", "parent", "bus", "device", "rcd_ma", "note"):
        if k in fields:
            v = fields[k]
            if k == "kind" and v not in KINDS:
                raise EditError(f"тип узла: {', '.join(KINDS)}")
            if k == "rcd_ma":
                v = int(v) if str(v or "").strip() else None
            if k in ("parent", "bus", "device", "note"):
                v = str(v or "").strip()
            out[k] = v
    if out.get("device", "x") == "":
        out["device"] = "авто"
    return out


def operate(prj: Project, op: dict) -> dict:
    """Выполнить операцию, записать файл и журнал. Возвращает {ok, node?}."""
    pid = op.get("panel", "")
    panel = prj.panel(pid)
    if panel is None:
        raise EditError("нет щита")
    data = load(prj.folder)
    st = _panel_state(data, pid)
    before = json.loads(json.dumps(st))
    numbers = {ln.number for ln in panel.lines}
    kind = op.get("op")
    res = {"ok": True}
    if kind == "add_node":
        f = _clean(op)
        node = {"id": _next_id(st), "kind": f.get("kind", "УЗО"), "parent": f.get("parent", ""), "bus": f.get("bus", ""),
                "device": f.get("device", "авто"), "rcd_ma": f.get("rcd_ma"), "note": f.get("note", "")}
        _check_parent(st, node["id"], node["parent"])
        st["nodes"].append(node)
        for n in op.get("lines", []):
            if n not in numbers:
                raise EditError(f"нет линии {n}")
            st["lines"][n] = node["id"]
        res["node"] = node["id"]
        what = f"узел {node['id']} ({node['kind']}), линий {len(op.get('lines', []))}"
    elif kind == "update_node":
        node = next((n for n in st["nodes"] if n["id"] == op.get("id")), None)
        if node is None:
            raise EditError(f"нет узла {op.get('id')}")
        f = _clean(op.get("fields", {}))
        if "parent" in f:
            _check_parent(st, node["id"], f["parent"])
        node.update(f)
        what = f"узел {node['id']}: " + ", ".join(f"{k} = {v}" for k, v in f.items())
    elif kind == "delete_node":
        node = next((n for n in st["nodes"] if n["id"] == op.get("id")), None)
        if node is None:
            raise EditError(f"нет узла {op.get('id')}")
        up = node.get("parent", "")
        for n in st["nodes"]:                       # дети узла — к его родителю
            if n.get("parent") == node["id"]:
                n["parent"] = up
        for ln, nid in list(st["lines"].items()):
            if nid == node["id"]:
                if up:
                    st["lines"][ln] = up
                else:
                    del st["lines"][ln]
        st["nodes"].remove(node)
        what = f"узел {node['id']} убран, линии — к {up or 'шине'}"
    elif kind == "move_lines":
        target = op.get("node", "")
        if target and target not in {n["id"] for n in st["nodes"]}:
            raise EditError(f"нет узла {target}")
        for n in op.get("lines", []):
            if n not in numbers:
                raise EditError(f"нет линии {n}")
            if target:
                st["lines"][n] = target
            else:
                st["lines"].pop(n, None)
        what = f"линии {', '.join(op.get('lines', []))} -> {target or 'шина'}"
    else:
        raise EditError(f"неизвестная операция {kind}")
    save(prj.folder, data)
    _journal(prj, {"kind": "структура", "panel": pid, "doc": panel.doc, "what": what, "before": before})
    return res


def _journal(prj, entry):
    from .edit import _journal as jpath, _user
    entry = {"batch": f"{time.strftime('%Y-%m-%d_%H%M%S')}|структура|{entry['panel']}|{uuid.uuid4().hex[:8]}",
             "at": time.strftime("%d.%m.%Y %H:%M:%S"), "user": _user(), "changes": [], **entry}
    with open(jpath(prj), "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def undo_last(prj: Project) -> dict:
    from .edit import read_journal, _journal as jpath
    rows = [r for r in read_journal(prj, 10_000) if r.get("kind") == "структура" and not r["undone"]]
    if not rows:
        return {"ok": False, "errors": ["нечего отменять в структуре"]}
    last = rows[-1]
    data = load(prj.folder)
    data["panels"][last["panel"]] = last["before"]
    save(prj.folder, data)
    with open(jpath(prj), "a", encoding="utf-8") as f:
        f.write(json.dumps({"undo": last["batch"], "at": time.strftime("%d.%m.%Y %H:%M:%S")}, ensure_ascii=False) + "\n")
    return {"ok": True, "undone": last["what"]}


def node_dict(n: Node) -> dict:
    return dataclasses.asdict(n)
