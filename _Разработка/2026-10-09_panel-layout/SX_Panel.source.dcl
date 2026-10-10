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
      : edit_box { key = "nrows"; label = "Количество рядов:"; edit_width = 5; }
      : edit_box { key = "slots"; label = "Модулей в ряду:"; edit_width = 5; }
    }
    : row {
      : edit_box { key = "pitch"; label = "Шаг модуля, мм:"; edit_width = 5; }
      : edit_box { key = "reserve"; label = "Свободное место, %:"; edit_width = 5; }
    }
    : edit_box { key = "gap"; label = "Зазор между аппаратами соседних рядов, мм:"; edit_width = 6; }
    : edit_box { key = "devh"; label = "Минимальная высота аппаратов, мм:"; edit_width = 6; }
    : text { label = "Ваши аппараты и клеммы XT; клеммы размещаются отдельным нижним рядом."; }
  }
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
