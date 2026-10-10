@echo off
chcp 65001 >nul
cd /d "%~dp0"
where git >nul 2>&1 || goto nogit
git config user.name >nul 2>&1 || git config user.name "%USERNAME%"
git config user.email >nul 2>&1 || git config user.email "%USERNAME%@%COMPUTERNAME%"
set OPEN=
for /r %%f in (~$*) do set OPEN=%%~nxf
for /r %%f in (*.dwl) do set OPEN=%%~nxf
if defined OPEN goto opened

echo Получаю изменения с GitHub...
git pull --rebase --autostash
if errorlevel 1 goto conflict
echo.
echo Готово: папка обновлена.
goto end

:conflict
echo.
echo ВНИМАНИЕ: одни и те же файлы менялись и здесь, и на другом ПК:
git diff --name-only --diff-filter=U
git rebase --abort >nul 2>&1
echo.
echo Ничего не потеряно: папка оставлена как была до обновления.
echo Откройте Claude в этой папке и скажите: "конфликт при получении".
goto end

:opened
echo.
echo Открыт файл в Excel или AutoCAD (найден %OPEN%).
echo Сохраните и закройте таблицы и чертежи из этой папки, потом запустите снова.
echo Если всё закрыто, а сообщение повторяется - удалите этот файл: он остался после сбоя.
goto end

:nogit
echo Не найден Git. Установите Git for Windows: https://git-scm.com/download/win
goto end

:end
echo.
pause
