# -*- coding: utf-8 -*-
"""Адреса книги «Однолинейка» v3.12 — ВСЕ в этом файле.

Таблица поменялась (новый столбец, сдвиг строки) — править только здесь.
Сверено с шапками листов и с надстройкой (SX_Ribbon.bas: CollectInfo, SubSheet).
Поле модели -> столбец. Путь поля — как в модели: «input.length», «result.du_sum».
"""

SRC, SPLIT, ODN, NET = "Исходные данные", "Разбивка_по_щитам", "Однолинейка", "Сеть_щита"
KJ = "Кабельный_журнал"          # КЖ: P:W (графы 1–8) со строки 5, строка — где есть маркировка (как SX_KJ.bas)
KJ_FIRST, KJ_COL, KJ_N = 5, 16, 8
REQUIRED = (SRC, ODN)            # без них — не книга «Однолинейка»
FIRST_ROW = 3                    # первая строка данных на листах линий
LAST_ROW = 1000
NET_ROWS = 140

# ---------------------------------------------------------------- «Исходные данные»: одна строка = одна линия
SRC_NUMBER = "D"
SRC_NAME = ("C", "AJ", "B")      # ручное, итоговое, автоматическое — первое непустое
SRC_FIELDS = {
    "input.p_kw": "E", "input.u_v": "F", "input.count": "G", "input.phases": "H", "input.length_plan": "I",
    "input.length": "M", "input.cable_code": "N", "input.cores": "O", "input.cos": "P", "input.s_min": "Q",
    "input.in_manual": "R", "input.group_ks": "BC", "input.parallel": "CC", "input.low_current": "CB",
    "input.section": "BB",
    "result.i_calc": "T", "result.in_by_current": "U", "result.s_mm2": "V", "result.du": "AB",
    "result.i_start": "AT", "result.du_max": "BF", "result.du_sum": "BG", "result.ikz_min": "BI",
    "result.ikz_ratio": "BJ", "result.ikz_check": "BK", "result.ikz_max": "BM", "result.category": "BP",
    "result.fire": "BQ", "result.category_check": "BV",
    "scheme.cable": "AO",
}
# куда идти править поле (для правок из окна — шаг 2): ручные столбцы
SRC_EDIT = {"input.p_kw": "E", "input.length_plan": "I", "input.cable_code": "N", "input.cores": "O",
            "input.s_min": "Q", "input.in_manual": "R", "input.parallel": "CC", "name": "C"}

# ---------------------------------------------------------------- «Разбивка_по_щитам»: линия -> щит
SPLIT_PANEL, SPLIT_LINE = "F", "H"

# ---------------------------------------------------------------- «Однолинейка»: одна строка = место щита
ODN_PANEL, ODN_LINE, ODN_QF, ODN_BREAKER = "A", "R", "G", "H"
ODN_FIELDS = {
    "scheme.place": "B", "scheme.sheet": "C", "scheme.bus": "D", "scheme.terminal": "P",
    "scheme.cable": "Q", "scheme.phase": "V",
}
ODN_DEVICES = ("I", "J", "K")
ODN_IN = "AU"                    # номинал автомата группы (итог)
ODN_STATUS = "AE"                # статус строки
ODN_LEVEL = "BU"                 # 2 ошибка, 1 предупреждение
ODN_EDIT = {"scheme.breaker_set": "E", "scheme.bus": "D"}
ODN_BREAKER_SET = "E"

# ---------------------------------------------------------------- «Сеть_щита»
NET_PANEL_ROWS = (6, 7, 8)       # раздел 1: до 3 щитов книги; столбцы C/D/E раздела 3 — в том же порядке
NET_PANEL_COLS = "CDE"
NET_SUPPLY = {                   # раздел 1, строка щита
    "name": "B", "du_bus": "C", "ikz_max": "E", "in_a": "F", "from_doc": "G", "qf": "H",
    "breaker": "I", "cable": "J", "length": "K", "from_panel": "L",
}
NET_MANUAL_FROM = (12, "C")      # раздел 2: «Питается от» (ТП, ГРЩ…)
NET_EDIT = {"supply.manual": NET_MANUAL_FROM}
NET_RESULT_ROWS = {              # раздел 3: строка -> поле (столбец щита из NET_PANEL_COLS)
    29: "p_calc", 30: "i_calc", 31: "cos", 33: "du_bus", 34: "zp_bus", 35: "ikz_min_bus", 36: "ikz_max_bus",
}
NET_CHECK_ROWS = {38: "Аппарат и кабель", 39: "Отключение КЗ", 40: "Термостойкость", 41: "Селективность"}
NET_ICU_ROW = 42
NET_VRU_ROW = 26                 # «Этот щит — ВРУ здания» (patch_vru.py): «да» / пусто, столбец щита
NET_LOAD_ROW = 45                # раздел 4: Руст щита (строка 45 + № щита), столбец C
NET_INPUTS_ROW, NET_AVR_ROW, NET_CAT_ROW = 65, 66, 68
NET_SUB_ROW, NET_SUB_COUNT = 94, 30      # раздел Б: фидеры к нижестоящим
# «Вводы 2 и 3» (patch_inputs.py): U4 = метка; щит k (1–3): строки 6 + (k-1)·2 (ввод 2) и +1 (ввод 3)
NET_IN23_MARK = (4, "U", "ВВОДЫ 2 И 3")
NET_IN23 = {"input": "V", "du_bus": "W", "ikz_max": "Y", "in_a": "Z", "from_doc": "AA",
            "qf": "AL", "breaker": "AM", "cable": "AN", "length": "AO", "from_panel": "AP"}
NET_SUB = {"line": "B", "doc": "C", "p_inst": "D", "i_calc": "J", "in_a": "N", "check": "O"}
