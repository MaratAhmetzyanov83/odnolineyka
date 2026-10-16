# -*- coding: utf-8 -*-
"""Проверки объекта. Новый файл правил — добавить в импорт ниже."""
from .registry import Hit, rule, rules, run  # noqa: F401
from . import source_status, links  # noqa: F401,E402  — регистрируют правила
