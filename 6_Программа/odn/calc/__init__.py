# -*- coding: utf-8 -*-
"""Расчёт в Python (шаг 3 плана). Пока пусто: результаты линий и щитов приходят из Excel (result.by = "excel").

Договор:
- calc.run(prj) заполняет line.result / panel.result и ставит result.by = "python";
- чистые функции по темам: loads.py (Руст, Кс, Рр, Iр), select.py (автомат, сечение), du.py (ΔU), kz.py (Iкз),
  categories.py; таблицы норм — norms.py (из «Справочной»), с источником (СП 256, ПУЭ…);
- сверка: compare.py сравнивает python и excel по каждой линии; пока есть расхождения — главный Excel.
- Подключается шагом step_calc в services/project.py.
"""
