# -*- coding: utf-8 -*-
"""Модель объекта: Объект -> Документы (книги, позже — свой файл) -> Щиты -> Линии.

Правила модели:
- Ядро не знает, откуда данные (Excel, свой файл, DWG). Это знают адаптеры в `odn.sources`.
- У линии и щита три части:
    input  — что задаёт инженер (мощность, длина, марка…);
    scheme — что стоит на схеме (место, QF, автомат, клемма…);
    result — что получено расчётом (Iр, сечение, ΔU, Iкз…). `result.by` — кто считал: «excel» или «python».
  Так расчёт переносится в Python по частям и сверяется с Excel.
- Новые поля добавлять со значением по умолчанию — старые JSON и кэш читаются (см. serialize.from_dict).
- Ссылки между объектами — по id (строки), не по объектам: модель сериализуется в JSON без потерь.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

SCHEMA = 2            # версия формата JSON; менять при несовместимых правках модели


# ---------------------------------------------------------------- место в источнике (куда идти править)
@dataclass
class Ref:
    file: str = ""        # путь относительно папки объекта
    sheet: str = ""
    cell: str = ""

    def __str__(self):
        return f"{self.sheet}!{self.cell}" if self.sheet else self.file


# ---------------------------------------------------------------- линия
@dataclass
class LineInput:
    p_kw: Optional[float] = None        # Pу, кВт
    u_v: Optional[float] = None         # Uн, В
    phases: Optional[int] = None
    count: Optional[int] = None         # кол-во потребителей (для длины)
    length_plan: Optional[float] = None
    length: Optional[float] = None      # L итог, м
    cable_code: str = ""                # код марки (мс, мфр…)
    cores: Optional[int] = None
    cos: Optional[float] = None
    s_min: Optional[float] = None       # сечение ручное / мин. по коду, мм²
    in_manual: Optional[float] = None   # In ручной, А
    group_ks: str = ""                  # Быт / Сил / Пост / Нет / Щит
    parallel: Optional[int] = None      # кабелей параллельно
    low_current: bool = False           # слаботочная
    section: Optional[int] = None       # секция щита (ввод): 1 / 2 / 3, пусто = 1 («Исходные данные» BB)


@dataclass
class LineScheme:
    place: Optional[int] = None         # № места
    sheet: Optional[int] = None         # лист схемы
    bus: str = ""                       # шина (пусто = основная)
    qf: str = ""                        # QF1, QFD3…
    breaker: str = ""                   # «2P С16А 30мА» (итог на схеме)
    breaker_set: str = ""               # что задано: «авто» или текст (правится; только у первой строки группы)
    group_head: bool = False            # первая строка группы (у неё свой автомат)
    in_a: Optional[float] = None        # In автомата группы на схеме, А
    devices: list = field(default_factory=list)
    terminal: str = ""
    phase: str = ""                     # L1 / L1,L2,L3
    cable: str = ""                     # кабель текстом: марка (жилы×сечение) L-ΔU


@dataclass
class LineResult:
    by: str = ""                        # кто считал: excel / python
    i_calc: Optional[float] = None      # Iр, А
    i_start: Optional[float] = None     # Iпуск, А
    in_by_current: Optional[float] = None   # In по току, А
    s_mm2: Optional[float] = None       # сечение, мм²
    du: Optional[float] = None          # ΔU линии, %
    du_sum: Optional[float] = None      # ΔU от ВРУ, %
    du_max: Optional[float] = None      # допустимая, %
    ikz_min: Optional[float] = None     # Iкз min в конце, А
    ikz_ratio: Optional[float] = None   # Iкз min / In
    ikz_check: str = ""
    ikz_max: Optional[float] = None     # кА
    category: str = ""
    fire: str = ""
    category_check: str = ""
    du_path_len: Optional[float] = None # длина проводки от ВРУ: линия + фидеры выше, м (calc.du)
    du_extra: Optional[float] = None    # надбавка к ΔUдоп за длину > 100 м, % (ГОСТ Р 50571.5.52, прил. G)
    du_max_eff: Optional[float] = None  # допустимая ΔU с надбавкой, %
    du_from: str = ""                   # от какого ВРУ здания считаются потери (id щита)
    du_sum_eff: Optional[float] = None  # ΔU от ВРУ здания, % (книга минус ΔU до шин ВРУ, если книга его не знает)
    status: str = ""                    # статус строки (как в источнике)
    level: int = 0                      # 0 ок, 1 внимание, 2 ошибка


@dataclass
class Line:
    number: str                         # номер линии — уникален в щите
    panel: str = ""                     # id щита
    name: str = ""
    input: LineInput = field(default_factory=LineInput)
    scheme: LineScheme = field(default_factory=LineScheme)
    result: LineResult = field(default_factory=LineResult)
    refs: dict = field(default_factory=dict)    # где в источнике: {"input": Ref, "scheme": Ref, ...} как dict
    feeds: str = ""                     # id щита, который питает эта линия (фидер)
    feeds_input: int = 0                # какой ввод того щита питает эта линия (1–3)
    node: str = ""                      # узел щита, под которым линия (структура программы); пусто — на шине
    rcd_node: str = ""                  # узел выше с 30 мА (УЗО / АВДТ) — заполняет calc.nodes

    @property
    def id(self):
        return f"{self.panel} # {self.number}"

    @property
    def in_a(self):
        """In автомата линии: со схемы, иначе — по току."""
        return self.scheme.in_a if self.scheme.in_a is not None else self.result.in_by_current


# ---------------------------------------------------------------- щит
@dataclass
class Supply:
    """Откуда питается щит — как записано в книге щита."""
    from_doc: str = ""                  # файл вышестоящего (как в книге, может быть относительным)
    from_panel: str = ""                # имя щита-источника
    manual: str = ""                    # вручную: «ТП», «ГРЩ»…
    qf: str = ""
    breaker: str = ""
    cable: str = ""
    length: Optional[float] = None
    in_a: Optional[float] = None
    du_bus: Optional[float] = None
    ikz_max: Optional[float] = None
    input: int = 1                      # номер ввода щита (1 — основной, 2–3 — «Вводы 2 и 3»)
    # связь ввода (services.linking): от какого щита и линии
    parent: str = ""
    parent_line: str = ""


@dataclass
class PanelScheme:
    inputs: Optional[int] = None        # число вводов
    avr: str = ""
    icu: Optional[float] = None         # отключающая способность аппаратов, кА
    is_vru: str = ""                    # «да» — ВРУ здания (книга: «Сеть_щита» C26:E26): ΔU линий от его шин


@dataclass
class PanelResult:
    by: str = ""
    p_inst: Optional[float] = None      # Руст, кВт
    p_calc: Optional[float] = None      # Рр, кВт
    i_calc: Optional[float] = None      # Iр, А
    cos: Optional[float] = None
    du_bus: Optional[float] = None      # ΔU до шин, %
    zp_bus: Optional[float] = None
    ikz_min_bus: Optional[float] = None
    ikz_max_bus: Optional[float] = None
    category: str = ""                  # обеспечиваемая вводами
    checks: dict = field(default_factory=dict)   # проверки питающей линии: название -> текст
    errors: int = 0                     # ошибок по оценке источника
    warnings: int = 0


@dataclass
class FeederSeen:
    """Что вышестоящий щит записал о нижестоящем (для сверки связей)."""
    line: str = ""
    doc: str = ""
    p_inst: Optional[float] = None
    i_calc: Optional[float] = None
    in_a: Optional[float] = None
    check: str = ""
    ref: Ref = field(default_factory=Ref)


@dataclass
class Note:
    """Статус строки источника, не привязанный к линии (например, строка группы без номера)."""
    level: int = 0
    text: str = ""
    label: str = ""
    ref: Ref = field(default_factory=Ref)


@dataclass
class Node:
    """Узел каскада: аппарат над группой (УЗО, АВДТ, автомат, рубильник). Данные — программы («_окно/структура.json»)."""
    id: str                             # У1, У2… — уникален в щите
    kind: str = "УЗО"                   # УЗО / АВДТ / Автомат / Рубильник
    parent: str = ""                    # узел выше; пусто — на шине
    bus: str = ""                       # шина узла верхнего уровня (пусто = основная)
    device: str = "авто"                # «авто» или текст аппарата («УЗО 2P 40А 30мА»)
    rcd_ma: Optional[int] = None        # ток утечки, мА (УЗО / АВДТ; пусто = 30)
    note: str = ""
    # расчёт (calc.nodes)
    designation: str = ""               # QD1, QF5…
    i_calc: Optional[float] = None      # Iр узла: сумма Iр линий под ним по наиболее нагруженной фазе, без Кс
    phase_i: dict = field(default_factory=dict)   # {"L1": …, "L2": …, "L3": …}
    phases: int = 1
    poles: str = ""
    in_a: Optional[float] = None
    text: str = ""                      # «УЗО 2P 40А 30мА»
    lines_all: list = field(default_factory=list)  # номера всех линий под узлом (все уровни)
    level: int = 1                      # уровень: 1 — на шине


@dataclass
class Panel:
    id: str                             # «документ / щит» — уникален в объекте
    name: str = ""
    doc: str = ""                       # документ (путь относительно папки объекта)
    supply: Supply = field(default_factory=Supply)
    scheme: PanelScheme = field(default_factory=PanelScheme)
    result: PanelResult = field(default_factory=PanelResult)
    lines: list = field(default_factory=list)            # list[Line]
    feeders_seen: list = field(default_factory=list)     # list[FeederSeen]
    notes: list = field(default_factory=list)            # list[Note]
    nodes: list = field(default_factory=list)            # list[Node] — структура щита (каскад)
    inputs_more: list = field(default_factory=list)      # list[Supply] — вводы 2 и 3 («Сеть_щита» строки 6–11)
    refs: dict = field(default_factory=dict)             # именованные места: {"supply": Ref, ...}
    # заполняет services.linking
    parent: str = ""
    parent_line: str = ""
    link_error: str = ""
    link_src: str = ""                  # откуда связь: «книга» / «программа»
    vru: bool = False                   # ВРУ здания (итог: программа, иначе книга) — calc.du
    vru_src: str = ""                   # «книга» / «программа»
    book_parent: str = ""               # связь, как записана в книге (для сверки со связью программы)
    book_line: str = ""
    # заполняет core.revision
    revision: str = ""
    # заполняет services.drawing_stamp: построения щита по чертежам (SXDRAW), последнее на каждый DWG
    drawings: list = field(default_factory=list)          # [{dwg, at, book_time, state ok/stale/check, why}]


# ---------------------------------------------------------------- документ, объект
@dataclass
class Document:
    file: str                           # путь относительно папки объекта
    kind: str = ""                      # адаптер: excel_v312, …
    external: bool = False              # книга из другой папки: найдена по связям (вверх / вниз)
    mtime: float = 0.0
    error: str = ""
    panels: list = field(default_factory=list)           # list[Panel]


@dataclass
class Issue:
    rule: str                           # id правила (checks)
    level: int                          # 2 ошибка, 1 внимание
    text: str
    panel: str = ""
    line: str = ""
    ref: Ref = field(default_factory=Ref)
    group: str = ""                     # группа правила — для фильтров окна


@dataclass
class Project:
    name: str = ""
    folder: str = ""
    schema: int = SCHEMA
    loaded: str = ""
    documents: list = field(default_factory=list)        # list[Document]
    issues: list = field(default_factory=list)           # list[Issue]

    # ---- индексы (не сериализуются: строятся reindex())
    def reindex(self):
        self._panels = {p.id: p for d in self.documents for p in d.panels}
        self._kids = {}
        for p in self._panels.values():
            if p.parent:
                self._kids.setdefault(p.parent, []).append(p)
        return self

    def panel_raw(self, pid):
        """Щит по id без reindex (во время связывания)."""
        return next(p for d in self.documents for p in d.panels if p.id == pid)

    def panels(self):
        return list(self._panels.values())

    def panel(self, pid) -> Optional[Panel]:
        return self._panels.get(pid)

    def children(self, pid):
        return self._kids.get(pid, [])

    def line(self, pid, number) -> Optional[Line]:
        p = self.panel(pid)
        return next((ln for ln in p.lines if ln.number == number), None) if p else None
