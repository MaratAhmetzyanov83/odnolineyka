# -*- coding: utf-8 -*-
"""Источники данных (адаптеры). Каждый умеет узнать свой файл и прочитать его в Document.

Добавить источник (свой файл объекта, другая версия таблицы, DWG):
  1. модуль в odn/sources/<имя>/ с классом-наследником Source;
  2. register(<Класс>()) в конце модуля и импорт ниже.
Порядок регистрации = приоритет: первый, кто сказал can_read() — читает.
"""
from .base import Source, register, registry, source_for  # noqa: F401
from .excel_v312 import reader as _excel_v312  # noqa: F401,E402  — регистрирует себя
