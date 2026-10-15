# -*- coding: utf-8 -*-
"""Довести книгу v3.12 (старый объект) до текущей версии шаблона — все обновления по порядку.

    py update_book.py <книга.xlsx> [...]   — книги закрыты; повторный запуск ничего не портит

Порядок:
  1. tidy_book.py          — «Пульт», порядок и цвет листов, оформленная «Справочная», уровни столбцов
  2. fix_text_format.py    — TEXT(x,"0.0") -> FIXED (#ЗНАЧ! на русском Excel)
  3. patch_audit.py        — аудит, ошибки 1–4: ΔU по остатку, Ко только с бытом, «мфрх», КЗ «не проверено»
  4. merge_network.py      — один лист «Сеть_щита» (питание + нижестоящие + токи КЗ)
  5. patch_inputs.py       — два ввода: Вводы 2 и 3, нагрузка по вводам для вышестоящего
  6. patch_validation.py   — валидация, ошибки 15–18: ввод от Iр щита, 120 мм², фазы без «Нет», дыры I/K/L
  7. patch_review.py       — замечания 19–22: двигатели (КПД, пуск, D), Icu, освещение МОП (LM), мощность ввода
  8. patch_feeder.py       — данные фидера для схемы (QF, автомат, кабель, L, щит-источник), селективность ввода (QS)
  9. patch_categories.py   — категория шины по цепочке источников: ввод(ы), ДГУ, ИБП (блок «ДГУ и ИБП»)
 10. patch_assembly.py     — лист «Сборка_щита» (CPDRAW v2.0), «Каталог» N для провода / шин / корпуса, столбец Q
 11. patch_vru.py          — «Сеть_щита» строка 26 «ВРУ здания»: ΔU линий от его шин при связи с вышестоящим
 12. patch_parallel.py     — «Исходные данные»: «Кабелей параллельно n» (сечение, ΔU, Zп, Iкз, марка, Iдоп)
 13. patch_char.py         — характеристика автомата B / C / D сама по Iкз (ПУЭ 1.7.79), «Однолинейка» CC, CD
 14. patch_input_du.py     — «Исходные данные» BG: ΔU от ВРУ в строке ввода без двойного учёта кабеля ввода
 15. patch_ks.py           — Кс «Сил» сам по числу ЭП в щите (лифты, двигатели; «Справочная» CE:CN, «Нагрузка_щитов» BB:BD)
 16. patch_km.py           — узел КМ (АПС): контакторы «Справочная» CY:DD, «Однолинейка» CE:CI (подбор КМ), «Каталог» N
После объединения листов старые ручные связи между файлами рвутся: в каждом щите — «Связать с вышестоящим».
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
STEPS = ["tidy_book.py", "fix_text_format.py", "patch_audit.py", "merge_network.py", "patch_inputs.py",
         "patch_validation.py", "patch_review.py", "patch_feeder.py", "patch_categories.py",
         "patch_assembly.py", "patch_vru.py", "patch_parallel.py", "patch_char.py", "patch_input_du.py",
         "patch_ks.py", "patch_km.py"]


def main(paths):
    for p in paths:
        p = os.path.abspath(p)
        print(f"=== {os.path.basename(p)}")
        env = dict(os.environ, PYTHONIOENCODING="utf-8")    # шаги печатают ², Δ — консоль cp1251 падала
        for step in STEPS:
            r = subprocess.run([sys.executable, os.path.join(HERE, step), p], env=env,
                               capture_output=True, text=True, encoding="utf-8", errors="replace")
            out = (r.stdout or "").strip().splitlines()[1:]           # первая строка — имя книги
            print(f"  {step}: " + ("; ".join(x.strip() for x in out) if out else "ok"))
            if r.returncode != 0:
                print("  ОШИБКА:\n" + (r.stderr or "")[-1500:])
                print("  книга не доведена — остальные шаги пропущены")
                break


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1:])
