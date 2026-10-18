# -*- coding: utf-8 -*-
"""То, что уже проверил источник (Excel): статусы строк, категория, питающая линия.

Когда расчёт переедет в Python, эти правила заменятся своими (calc + checks), а эти останутся для сверки.
"""
import re

from ..core.model import Ref
from .registry import Hit, rule
from .util import ok_text

G = "Из книг"
DU_RE = re.compile(r"^потери [\d,.]+ % > [\d,.]+ %$")
ERR_WORDS = ("дубль", "ПЕРЕГРУЗ", "ЗАЩИЩ", "КАНАЛ", "без 30")      # как «Однолинейка» BU: ошибка, иначе внимание


def _err(text):
    return any(w in text for w in ERR_WORDS)


@rule("source.unreadable", "Книга не прочитана", G)
def unreadable(prj):
    for d in prj.documents:
        if d.error:
            yield Hit(2, f"книга не прочитана: {d.error}", ref=Ref(file=d.file))


@rule("source.row_status", "Статус строки «Однолинейки»", G)
def row_status(prj):
    for p in prj.panels():
        for ln in p.lines:
            status = ln.result.status
            if ln.rcd_node:                       # 30 мА дал узел выше (УЗО / АВДТ) — замечание Excel снято
                status = "; ".join(x for x in status.split("; ") if x.strip() != "розетка без 30 мА")
            r = ln.result
            du = r.du_sum_eff if r.du_sum_eff is not None else r.du_sum
            if (r.du_extra or (du is not None and r.du_sum is not None and du < r.du_sum - 1e-9)) \
                    and du is not None and r.du_max_eff is not None and du <= r.du_max_eff + 1e-9:
                # ΔU в допуске (от ВРУ здания, с надбавкой за длину) — замечание книги «потери … > …» снято
                status = "; ".join(x for x in status.split("; ") if not DU_RE.match(x.strip()))
            if ln.result.level and status:
                lvl = ln.result.level if status == ln.result.status else (2 if _err(status) else 1)
                yield Hit(lvl, status, p.id, ln.number, ln.refs.get("status"))
        for n in p.notes:
            yield Hit(n.level, n.text, p.id, n.label, n.ref)


@rule("source.category", "Категория надёжности не обеспечена", G)
def category(prj):
    for p in prj.panels():
        for ln in p.lines:
            if ln.result.category_check and not ok_text(ln.result.category_check):
                yield Hit(2, f"категория {ln.result.category}: {ln.result.category_check}", p.id, ln.number,
                          ln.refs.get("input"))


@rule("source.kz_unchecked", "КЗ не проверен (нет данных питания)", G)
def kz_unchecked(prj):
    for p in prj.panels():
        n = sum(1 for ln in p.lines if ln.result.ikz_check.startswith("не проверено"))
        if n:
            yield Hit(1, f"КЗ не проверен на {n} линиях: нет данных питания щита", p.id,
                      ref=p.refs.get("supply_manual") or p.refs.get("supply"))


@rule("source.supply_checks", "Проверки питающей линии", G)
def supply_checks(prj):
    for p in prj.panels():
        for k, v in p.result.checks.items():
            if not ok_text(v):
                lvl = 2 if k in ("Аппарат и кабель", "Отключение КЗ") else 1
                yield Hit(lvl, f"питающая линия, {k.lower()}: {v}", p.id, ref=p.refs.get("checks"))


@rule("source.du_over", "ΔU больше допустимой (от ВРУ здания)", G)
def du_over(prj):
    """ΔU от ВРУ здания (программа) больше допуска, а книга этого не видит (ВРУ задан только в программе)."""
    for p in prj.panels():
        for ln in p.lines:
            r = ln.result
            if r.du_sum_eff is None or r.du_max_eff is None or "потери" in (r.status or ""):
                continue
            if r.du_sum_eff > r.du_max_eff + 1e-9:
                yield Hit(1, f"потери {r.du_sum_eff:.2f} % > {r.du_max_eff:.2f} % (от ВРУ здания)".replace(".", ","),
                          p.id, ln.number, ln.refs.get("input"))
