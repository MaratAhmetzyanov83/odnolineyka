// SX_Layout.dcl v0.2 — окно команды SXLAYOUT (кодировка Windows-1251)
sxl_main : dialog {
  label = "Компоновка щита из Excel — SXLAYOUT";
  : boxed_column {
    label = "Данные Excel";
    : row {
      : edit_box { key = "path"; label = "Файл:"; edit_width = 58; allow_accept = false; }
      : button   { key = "browse"; label = "Обзор..."; fixed_width = true; width = 12; }
    }
    : popup_list { key = "panel"; label = "Щит:"; width = 30; }
  }
  : boxed_column {
    label = "Корпус (база SX_Корпуса.csv) и ограничения по месту";
    : row {
      : edit_box { key = "maxw"; label = "Ширина не более, мм:"; edit_width = 6; }
      : edit_box { key = "maxh"; label = "Высота не более:"; edit_width = 6; }
      : edit_box { key = "maxd"; label = "Глубина не более:"; edit_width = 6; }
    }
    : row {
      : popup_list { key = "inst"; label = "Установка:"; width = 30; }
      : popup_list { key = "series"; label = "Серия:"; width = 30; }
    }
    : popup_list { key = "encl"; label = "Корпус:"; width = 72; }
    : text { key = "enclinfo"; label = ""; width = 72; }
    : row {
      : edit_box { key = "reserve"; label = "Запас по рейкам, %:"; edit_width = 5; }
      : toggle   { key = "newrail"; label = "Каждый тип аппаратов с новой рейки"; }
    }
  }
  : boxed_column {
    label = "Вставка";
    : row {
      : text   { key = "ptinfo"; label = ""; width = 50; }
      : button { key = "pick"; label = "Указать точку <"; fixed_width = true; width = 22; }
    }
    : row {
      : edit_box { key = "lib"; label = "База блоков компоновки (DWG):"; edit_width = 44; allow_accept = false; }
      : button   { key = "libbrowse"; label = "Обзор..."; fixed_width = true; width = 12; }
    }
    : toggle { key = "clear"; label = "Удалить прежнюю компоновку (слои SXL_*)"; }
  }
  : text { key = "status"; label = ""; width = 70; }
  ok_cancel;
}
