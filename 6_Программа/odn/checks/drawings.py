# -*- coding: utf-8 -*-
"""Чертежи: актуальность однолинейных схем по отметкам SXDRAW (services.drawing_stamp)."""
from .registry import Hit, rule

G = "Чертежи"


def _name(path):
    return path.replace("\\", "/").rsplit("/", 1)[-1]          # путь Windows — и на Linux


@rule("drawings.stale", "Однолинейка устарела", G)
def stale(prj):
    """Схема построена по другим данным: перестроить SXUPDATE в чертеже."""
    for p in prj.panels():
        for d in p.drawings:
            if d["state"] == "stale":
                yield Hit(2, f"однолинейка устарела ({_name(d['dwg'])}, построена {d['at']}): {d['why']} "
                             f"— SXUPDATE в чертеже", p.id)


@rule("drawings.check", "Однолинейка: проверить", G)
def check(prj):
    for p in prj.panels():
        for d in p.drawings:
            if d["state"] == "check":
                yield Hit(1, f"однолинейка — проверить ({_name(d['dwg'])}, построена {d['at']}): {d['why']}", p.id)


@rule("drawings.kj_stale", "КЖ устарел", G)
def kj_stale(prj):
    """Кабельный журнал в DWG не совпадает с листом «Кабельный_журнал»: «КЖ в DWG» в книге."""
    for d in prj.documents:
        if d.kj.get("state") == "stale" and d.panels:
            yield Hit(2, f"КЖ устарел ({_name(d.kj['dwg'])}, записан {d.kj['at']}): {d.kj['why']} — «КЖ в DWG» в книге",
                      d.panels[0].id)
