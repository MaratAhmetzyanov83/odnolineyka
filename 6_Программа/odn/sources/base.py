# -*- coding: utf-8 -*-
"""Интерфейс источника данных."""
from __future__ import annotations

from ..core.changes import Change
from ..core.model import Document

_REGISTRY: list = []


class Source:
    kind = ""                       # имя адаптера, пишется в Document.kind
    extensions: tuple = ()          # какие файлы смотреть (.xlsx…)

    def can_read(self, path: str) -> bool:
        """Быстро: мой ли это файл (по расширению / имени листов)."""
        return path.lower().endswith(self.extensions)

    def read(self, path: str, rel: str) -> Document:
        """Прочитать файл в Document. Не бросать исключений из-за содержимого — писать Document.error."""
        raise NotImplementedError

    def cells_for(self, prj, changes: "list[Change]"):
        """Правки -> что и куда писать в файле (для журнала и записи). Не умеет — NotImplementedError."""
        raise NotImplementedError(f"{self.kind}: запись не поддерживается")

    def write(self, path: str, cells) -> None:
        """Записать в файл (резерв делает сервис). Ошибка для пользователя — EditError."""
        raise NotImplementedError(f"{self.kind}: запись не поддерживается")

    def open_at(self, path: str, ref) -> str:
        """Открыть файл у пользователя на нужном месте. Вернуть сообщение для окна."""
        raise NotImplementedError


def register(src: Source) -> Source:
    _REGISTRY.append(src)
    return src


def registry():
    return list(_REGISTRY)


def source_for(path: str):
    return next((s for s in _REGISTRY if s.can_read(path)), None)


def by_kind(kind: str):
    return next((s for s in _REGISTRY if s.kind == kind), None)
