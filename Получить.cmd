@echo off
chcp 65001 >nul
cd /d "%~dp0"
where git >nul 2>&1 || goto nogit
git config user.name >nul 2>&1 || git config user.name "%USERNAME%"
git config user.email >nul 2>&1 || git config user.email "%USERNAME%@%COMPUTERNAME%"
set OPEN=
rem замки Excel (~$) и AutoCAD (.dwl) скрытые - for /r их не видит, нужен dir /a
for /f "delims=" %%f in ('dir /s /b /a-d "~$*" "*.dwl" 2^>nul') do set OPEN=%%~nxf
rem надстройка, подключённая в Excel, замков не оставляет, но файл занят - git зависает на "Unlink ... (y/n)"
if exist "3_Макросы_Excel\Однолинейка.xlam" ren "3_Макросы_Excel\Однолинейка.xlam" "Однолинейка.xlam.chk" 2>nul && ren "3_Макросы_Excel\Однолинейка.xlam.chk" "Однолинейка.xlam" || if exist "3_Макросы_Excel\Однолинейка.xlam" set OPEN=Однолинейка.xlam - надстройка загружена в Excel
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
echo Если всё закрыто, а сообщение повторяется - удалите этот файл-замок (~$... или .dwl): он остался после сбоя. Надстройку .xlam не удалять - закройте Excel.
goto end

:nogit
echo Не найден Git. Установите Git for Windows: https://git-scm.com/download/win
goto end

:end
echo.
pause
