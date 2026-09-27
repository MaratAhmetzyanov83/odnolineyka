// SX_Schema.dcl v3 — окно команды SXDRAW (кодировка Windows-1251)
sx_main : dialog {
  label = "Однолинейка из Excel — SXDRAW";
  : boxed_column {
    label = "Данные Excel";
    : row {
      : edit_box { key = "path"; label = "Файл:"; edit_width = 58; allow_accept = false; }
      : button   { key = "browse"; label = "Обзор..."; fixed_width = true; width = 12; }
    }
    : popup_list { key = "sheet"; label = "Лист:"; width = 50; }
    : popup_list { key = "panel"; label = "Щит:"; width = 30; }
  }
  : boxed_column {
    label = "Чертёж";
    : row {
      : text   { key = "ptinfo"; label = ""; width = 44; }
      : button { key = "pick"; label = "Указать точку <"; fixed_width = true; width = 22; }
    }
    : row {
      : edit_box { key = "lib"; label = "Библиотека блоков (DWG):"; edit_width = 44; allow_accept = false; }
      : button   { key = "libbrowse"; label = "Обзор..."; fixed_width = true; width = 12; }
    }
    : row {
      : edit_box { key = "step"; label = "Шаг столбца, мм:"; edit_width = 6; }
      : edit_box { key = "nsheet"; label = "Столбцов на лист A3:"; edit_width = 5; }
    }
    : toggle   { key = "clear"; label = "Стереть прошлую отрисовку (слои SX_*)"; }
  }
  : text { key = "status"; label = ""; width = 70; }
  ok_cancel;
}
