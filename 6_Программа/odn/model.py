# -*- coding: utf-8 -*-
"""Модель объекта: объект -> щиты -> линии. Хранится в JSON («объект.json»).

Шаг 1: данные приходят из книг Excel v3.12 (импорт). Расчётные поля (Iр, In, S, ΔU, Iкз) пока — из Excel.
Потом расчёт переедет в Python и будет сверяться с этими же полями.
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Optional

FORMAT = 1


@dataclass
class Issue:
    """Замечание. level: 2 — ошибка, 1 — предупреждение."""
    level: int
    text: str
    panel: str = ""
    line: str = ""
    source: str = ""          # «Excel» — статус из книги; «объект» — проверка программы
    where: str = ""           # лист!ячейка в книге, куда смотреть


@dataclass
class Line:
    number: str                         # номер линии (D «Исходных данных»)
    name: str = ""
    panel: str = ""
    place: Optional[int] = None         # № места («Однолинейка» B)
    sheet: Optional[int] = None         # лист схемы
    bus: str = ""                       # шина (пусто = основная)
    # ввод
    p_kw: Optional[float] = None        # Pу, кВт
    u_v: Optional[float] = None
    phases: Optional[int] = None
    count: Optional[int] = None         # кол-во потребителей
    length_plan: Optional[float] = None
    length: Optional[float] = None      # L итог, м
    cable_code: str = ""                # марка (код: мс, мфр…)
    cores: Optional[int] = None
    cos: Optional[float] = None
    s_min: Optional[float] = None       # сечение ручное / мин.
    in_manual: Optional[float] = None   # In ручной
    group_ks: str = ""                  # Быт / Сил / Пост / Нет / Щит
    parallel: Optional[int] = None
    low_current: bool = False
    # расчёт (из Excel)
    i_calc: Optional[float] = None      # Iр, А
    i_start: Optional[float] = None     # Iпуск, А
    in_a: Optional[float] = None        # In автомата, А
    s_mm2: Optional[float] = None       # сечение, мм²
    du: Optional[float] = None          # ΔU линии, %
    du_sum: Optional[float] = None      # ΔU от ВРУ, %
    du_max: Optional[float] = None      # допустимая, %
    ikz_min: Optional[float] = None     # Iкз min в конце, А
    ikz_ratio: Optional[float] = None
    ikz_check: str = ""
    ikz_max: Optional[float] = None     # кА
    category: str = ""
    fire: str = ""
    category_check: str = ""
    cable: str = ""                     # кабель текстом (марка (жилы×сечение) L-ΔU)
    # «Однолинейка»
    qf: str = ""                        # QF1, QFD3…
    breaker: str = ""                   # «2P С16А 30мА»
    devices: list = field(default_factory=list)
    terminal: str = ""
    phase: str = ""                     # L1 / L1,L2,L3
    status: str = ""                    # статус строки (AE)
    level: int = 0                      # 0 ок, 1 предупреждение, 2 ошибка
    row_src: Optional[int] = None       # строка «Исходных данных»
    row_odn: Optional[int] = None       # строка «Однолинейки»
    feeds: str = ""                     # id щита, который питает эта линия (если фидер)


@dataclass
class Supply:
    """Откуда питается щит (раздел 1 или 2 «Сеть_щита»)."""
    from_file: str = ""                 # файл вышестоящего (как в книге, относительный)
    from_panel: str = ""                # имя щита-источника
    manual: str = ""                    # раздел 2: «ТП», «ГРЩ»… (если раздел 1 пуст)
    qf: str = ""
    breaker: str = ""
    cable: str = ""
    length: Optional[float] = None
    in_a: Optional[float] = None
    du_bus: Optional[float] = None      # ΔU от ВРУ до шин этого щита, %
    ikz_max: Optional[float] = None     # кА на шинах (от вышестоящего)


@dataclass
class Panel:
    id: str                             # «файл.xlsx / щит» — уникально в объекте
    name: str
    file: str                           # путь относительно папки объекта
    supply: Supply = field(default_factory=Supply)
    parent: str = ""                    # id вышестоящего щита (если найден в папке)
    parent_line: str = ""               # номер линии-фидера у вышестоящего
    inputs: Optional[int] = None
    avr: str = ""
    category: str = ""                  # обеспечиваемая вводами
    # итоги (из Excel)
    p_inst: Optional[float] = None      # Руст, кВт
    p_calc: Optional[float] = None      # Рр, кВт
    i_calc: Optional[float] = None      # Iр, А
    cos: Optional[float] = None
    du_bus: Optional[float] = None      # ΔU до шин, %
    zp_bus: Optional[float] = None
    ikz_min_bus: Optional[float] = None
    ikz_max_bus: Optional[float] = None
    icu: Optional[float] = None
    checks: dict = field(default_factory=dict)   # проверки «Сеть_щита» 38–41
    excel_errors: Optional[int] = None
    excel_warnings: Optional[int] = None
    lines: list = field(default_factory=list)    # list[Line]
    feeders_seen: list = field(default_factory=list)  # раздел Б: что вышестоящий «видит» о нижестоящих
    row_notes: list = field(default_factory=list)     # статусы строк «Однолинейки» без линии (row, qf, status, level)


@dataclass
class Book:
    file: str
    mtime: float = 0.0
    error: str = ""                    # не прочиталась
    panels: list = field(default_factory=list)   # list[Panel]


@dataclass
class Project:
    name: str
    folder: str
    format: int = FORMAT
    imported: str = ""
    books: list = field(default_factory=list)    # list[Book]
    issues: list = field(default_factory=list)   # list[Issue] — все замечания объекта

    def panels(self):
        for b in self.books:
            yield from b.panels

    def to_dict(self):
        return asdict(self)
