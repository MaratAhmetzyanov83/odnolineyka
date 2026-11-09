cp_main : dialog {
  label = "Компоновка электрощита — CPDRAW";
  : boxed_column {
    label = "Данные из Excel";
    : text { key = "file"; width = 75; }
    : button { key = "browse"; label = "Другая книга Excel..."; fixed_width = true; }
    : edit_box { key = "library"; label = "DWG с вашими блоками:"; edit_width = 55; }
    : popup_list { key = "panel"; label = "Щит:"; width = 35; }
  }
  : boxed_column {
    label = "Размещение модульных аппаратов по рядам";
    : row {
      : edit_box { key = "nrows"; label = "Рядов (0 — авто):"; edit_width = 5; }
      : edit_box { key = "slots"; label = "Модулей в ряду:"; edit_width = 5; }
    }
    : row {
      : edit_box { key = "pitch"; label = "Шаг модуля, мм:"; edit_width = 5; }
      : edit_box { key = "reserve"; label = "Свободное место, %:"; edit_width = 5; }
    }
    : edit_box { key = "railpitch"; label = "Межосевое расстояние DIN-реек, мм:"; edit_width = 6; }
    : edit_box { key = "devh"; label = "Минимальная высота аппаратов, мм:"; edit_width = 6; }
    : text { label = "Шаг между осями одинаковый для рядов аппаратов и клемм."; }
    : edit_box { key = "textheight"; label = "Высота подписей аппаратов, мм:"; edit_width = 6; }
    : edit_box { key = "xttextheight"; label = "Высота подписей клемм, мм:"; edit_width = 6; }
  }
  : boxed_column {
    label = "Сборка щита";
    : row {
      : toggle { key = "comb"; label = "Фазы L1-L2-L3 и гребёнки, провода"; }
      : toggle { key = "dense"; label = "Плотная раскладка"; }
    }
    : row {
      : toggle { key = "npe"; label = "Кросс-модуль N и PE"; }
      : toggle { key = "excel"; label = "Записать в Excel (лист «Сборка_щита»)"; }
    }
    : toggle { key = "rails"; label = "Рисовать DIN-рейки и корпус"; }
    : edit_box { key = "case"; label = "Корпус: артикул (пусто — подбор, «-» — без корпуса):"; edit_width = 16; }
    : edit_box { key = "caseseries"; label = "Серия для подбора (ST, STE, CQE N, RAM Base, FRAME; пусто — все):"; edit_width = 12; }
  }
  : toggle { key = "oneblock"; label = "Создавать единым блоком"; }
  : text { key = "status"; width = 75; label = ""; }
  ok_cancel;
}
cp_sizes : dialog {
  label = "Ваши блоки для компоновки щита";
  : text { label = "Выберите ваш блок для каждого типа. Габариты измеряются по блоку; геометрия сохраняется."; }
  : list_box { key = "items"; width = 78; height = 12; }
  : text { key = "model"; width = 78; }
  : popup_list { key = "state"; label = "Исполнение динамического блока:"; width = 65; }
  : row {
    : popup_list { key = "block"; label = "Блок:"; width = 32; }
    : edit_box { key = "scale"; label = "Масштаб:"; edit_width = 5; }
    : edit_box { key = "rotation"; label = "Угол, °:"; edit_width = 5; }
    : button { key = "apply"; label = "Применить"; fixed_width = true; }
  }
  : text { key = "status"; width = 78; label = ""; }
  ok_cancel;
}
