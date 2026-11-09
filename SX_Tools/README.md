# SX Tools — сводный инструмент для AutoCAD 2025

Вкладка **«SX»** на ленте AutoCAD вместо четырёх отдельных `APPLOAD`. Всё нужное лежит в этой папке.
Частые команды — на панелях, редкие — в выпадающей части панели (стрелка ▾ под названием панели).

| Панель вкладки | Команды | Файл |
|---|---|---|
| Однолинейка | SXDRAW, SXUPDATE, SXUPDSEL, SXUPDSHEET, SXBLOCKS, SXCLEAR | `CAD/Однолинейка/SX_Schema.lsp` |
| Компоновка щита | CPDRAW, CPUPDATE, CPEDIT, CPEXPLODE, CPDELETE | `CAD/Компоновка/SX_Panel.lsp` |
| Длины трасс | DLN, DLN_UPD, DLN_CHK, DLN_SVR, DLN_RAMKA, DLN_XLS, DLN_PREV, DLN_NEXT, DLN_CLR | `CAD/Длины_трасс/DLN.lsp` |
| Каталог блоков | SXCAT, SXGO, SXHOME | `CAD/Блоки/SX_Каталог.lsp` |
| Сервис | кабельный журнал (CableJournal.exe), образец КЖ, инструкции, перезагрузка LISP | `Tools/`, `Docs/` |

## Установка (один раз)
1. Двойной щелчок `Install.cmd` — делает ссылку (junction) в `%APPDATA%\Autodesk\ApplicationPlugins\SX_Tools.bundle`. Права администратора не нужны.
2. Перезапустите AutoCAD. Вкладка «SX» появится на ленте, LISP-файлы загрузятся в каждый чертёж.
3. Команда `SXT` — вернуть и открыть вкладку, `SXTRELOAD` — перезагрузить LISP после правки.
4. Снять — `Uninstall.cmd`. Папка остаётся.

При первой загрузке загрузчик один раз добавляет подпапки `CAD\…` в путь поддержки файлов AutoCAD и папку SX_Tools в `TRUSTEDPATHS` (нужно для `.dcl`, `.csv`, баз блоков). Старый `APPLOAD` этих файлов — убрать из «Автозагрузки», иначе грузятся дважды.

## Состав
- `Contents/SXTools.dll` — вкладка ленты и загрузчик (.NET 8, AutoCAD 2025). Исходник: `src/`. Пересборка: `Build.cmd` (AutoCAD закрыть).
- `CAD/SX_Tools.lsp` — загрузчик: грузит все `*.lsp` из подпапок `CAD/`.
- `CAD/…` — копии рабочих файлов из `2_AutoCAD`. Структура та же, поэтому `SX_Panel` и `SX_Schema` находят `..\Блоки` как раньше.
- `Update_from_project.cmd` — подтянуть свежие версии из `2_AutoCAD`/`4_Инструкции` (оригиналы остаются главными).

Базы блоков (`SX_Компоновка_пользовательские.dwg`, 62 МБ) в GitHub не идут — см. `.gitignore`.
