# -*- coding: utf-8 -*-
"""make_build_v5.py - собрать build_v5.py: ядро build_v4.py (листы «Щиты», «Режимы», журнал, итоги…) + новые части v5.

    py make_build_v5.py        -> build_v5.py

Новые части: lines_v5_block.py (таблица «Линии» v5), build_v5_tail.py (перенос данных, сборка книги .xlsm, макросы).
"""
import io
import os

HERE = os.path.dirname(os.path.abspath(__file__))
V4 = os.path.join(HERE, "..", "..", "2026-10-03_v4", "scripts", "build_v4.py")


def rd(p):
    return io.open(p, encoding="utf-8").read()


def cut(s, start_marker, end_marker):
    """индексы блока: от начала строки-рамки перед start_marker до начала строки-рамки перед end_marker."""
    i = s.index(start_marker)
    i = s.rindex("# ----", 0, i)
    j = s.index(end_marker, i)
    j = s.rindex("# ----", 0, j)
    return i, j


def main():
    s = rd(V4)
    # 1. шапка
    i = s.index('"""', 0)
    j = s.index('"""', i + 3) + 3
    s = s[:i] + ('"""build_v5.py - книга однолинейки v5 (один файл на объект, .xlsm с макросами). СОБРАН make_build_v5.py — правьте исходники.\n\n'
                 '    py build_v5.py --data <obj.json> --base "<книга v3.12>" --out "<Однолинейка_v5_объект.xlsm>" [--defaults]\n\n'
                 'Отличия от v4: в «Линиях» короткая строка (серое — выбор таблицы, синее — своё), маркер группы «гр», автоподбор автомата и\n'
                 'сечения, фазы — автобалансировка по автоматам, кнопки щитов с карточкой, лист «Экспликация», макросы книги.\n"""') + s[j:]
    s = s.replace('from model_v312 import ExcelSession, log  # noqa: E402',
                  'from model_v312 import ExcelSession, log  # noqa: E402\n'
                  'sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "2026-10-03_v4", "scripts"))\n'
                  'sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))')
    # 2. таблица «Линии»
    i, j = cut(s, "# Таблица «Линии»", "# Таблица «Щиты»")
    s = s[:i] + rd(os.path.join(HERE, "lines_v5_block.py")).rstrip() + "\n\n" + s[j:]
    # 3. «Щиты»
    rep = [
        ('Линии[Р, кВт],Линии[Щит],[@Щит],Линии[Фаза],"L1,L2,L3"', 'Линии[_Р линии, кВт],Линии[Щит],[@Щит],Линии[_фаза итог],"L1,L2,L3"'),
        ('SUMIFS(Линии[Р, кВт],Линии[Щит],[@Щит],Линии[Фаза],"L{k}"', 'SUMIFS(Линии[_Р линии, кВт],Линии[Щит],[@Щит],Линии[_фаза итог],"L{k}"'),
        ('scol("Iкз max на шинах, кА", "c", \'=IF(ISNUMBER([@[Zmax на шинах, Ом]]),1.05*(Uн_лин*1000/SQRT(3))/[@[Zmax на шинах, Ом]]/1000,"")\')',
         'scol("Iкз max на шинах, кА", "c", \'=IF(ISNUMBER([@[Zmax на шинах, Ом]]),1.05*(Uн_лин*1000/SQRT(3))/[@[Zmax на шинах, Ом]]/1000,"")\')\n'
         '# для подбора сечений в «Линиях»: те же ΔU и Z до шин, но по сечениям фидеров «по току» (иначе цикл щит → фидер → щит)\n'
         'scol("_ΔU фидера (подбор)", "h", \'=IF([@[_строка фидера]]="","",INDEX(Линии[_ΔU по току, %],[@[_строка фидера]]))\')\n'
         'scol("ΔU до шин (подбор), %", "h", f\'=SUMPRODUCT({ONPATH},IFERROR(--[_ΔU фидера (подбор)],0)+IFERROR(--[_ΔU источника],0))\')\n'
         'scol("_Zфидера (подбор)", "h", \'=IF([@[_строка фидера]]="","",INDEX(Линии[_Zкабеля по току],[@[_строка фидера]]))\')\n'
         'scol("Zпетли на шинах (подбор), Ом", "h",\n'
         '     f\'=IF(XLOOKUP([@Корень],[Щит],[_Zисточника],"")="","",SUMPRODUCT({ONPATH},IFERROR(--[_Zисточника],0)+IFERROR(--[_Zфидера (подбор)],0))+SUMPRODUCT({ONPATH})*R_пер)\')'),
        ('"Лист задаёт настройки ВСЕГО объекта. Обычно меняют только C3 (название), C8 (методика) и C9 (kп.к). Остальное — по необходимости.",',
         '"Лист задаёт настройки ВСЕГО объекта. Обычно меняют только C3 (название), C8 (методика), C9 (kп.к) и C10–C12 (наименования, марка для пожарных линий).",\n'
         '    "C10 Наименование линии: «по коду и № помещения» — «Розеточная группа, пом. 105»; «по экспликации» — название помещения из таблицы «Экспликация» (лист «Коды»). '
         'C11 — писать ли «пом.». C12 — марка кабеля для пожарных линий (обозначение из «Коды» → «Кабели»). '
         'C13–C14 — линии до 50 В (LED 24 В): запас мощности БП (1,3) и допустимые потери от БП (5 %).",'),
    ]
    for a, b in rep:
        assert a in s, "нет в build_v4.py: " + a[:90]
        s = s.replace(a, b)
    import core_patches_v5 as cp
    for a, b in cp.REP:
        assert s.count(a) == 1, "правка ядра не найдена (или не одна): " + a[:90]
        s = s.replace(a, b)
    c0, c1 = cp.CUT_LOADS_CATEGORIES
    i = s.index(c0)
    j = s.index(c1, i)
    s = s[:i] + s[j:]
    for a, b in cp.REP_AFTER_CUT:
        assert a in s, a
        s = s.replace(a, b)
    # 3б. дополнительные столбцы «Щитов» (расклад фаз) — перед функцией F2
    k = s.index("\n\ndef F2(")
    s = s[:k] + "\n\n" + rd(os.path.join(HERE, "shields_v5_extra.py")).rstrip() + s[k:]
    # 4. повтор заголовков (Excel не различает регистр)
    s = s.replace('_dup = [h for h, n in __import__("collections").Counter(c["head"] for c in L).items() if n > 1] +',
                  '_dup = [h for h, n in __import__("collections").Counter(c["head"].lower() for c in L).items() if n > 1] +')
    # 5. хвост: перенос данных, сборка, сверка
    i = s.index("def prepare_rows(data):")
    i = s.rindex("# ----", 0, i)
    s = s[:i] + rd(os.path.join(HERE, "docs_v5_block.py")).rstrip() + "\n\n\n" + rd(os.path.join(HERE, "build_v5_tail.py"))
    out = os.path.join(HERE, "build_v5.py")
    io.open(out, "w", encoding="utf-8").write(s)
    compile(s, out, "exec")
    print("записан", out, len(s.splitlines()), "строк")


if __name__ == "__main__":
    main()
