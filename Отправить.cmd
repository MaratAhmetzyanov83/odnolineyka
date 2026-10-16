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

set GDRIVE=G:\Мой диск\odnolineyka
if not exist "G:\Мой диск\" goto nogdrive
echo Копирую папку на Google Диск...
robocopy "%CD%" "%GDRIVE%" /MIR /XD .git /XF ~$* *.dwl *.dwl2 /R:1 /W:2 /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 (echo Не все файлы скопировались на Google Диск - запустите позже.) else (echo Google Диск: копия обновлена.)
goto gitstep
:nogdrive
echo Google Диск не найден (диск G:) - копия не обновлена. Запустите Google Drive.
:gitstep

echo Сохраняю изменения этого ПК...
git add -A
powershell -NoProfile -Command "[Console]::OutputEncoding=[Text.Encoding]::UTF8; $n=0; git -c core.quotepath=false diff --cached --name-only --diff-filter=AM | %%{ if ((Test-Path -LiteralPath $_) -and ((Get-Item -LiteralPath $_).Length -gt 95MB)) { git reset -q -- $_; Write-Host ('   ' + $_); $n++ } }; exit $n"
if errorlevel 1 echo Эти файлы больше 95 МБ - GitHub их не принимает, они остались только на этом ПК.
git diff --cached --quiet
if errorlevel 1 git commit -q -m "%COMPUTERNAME% %date% %time:~0,5%"
echo Сверяю с GitHub...
git pull --rebase --autostash
if errorlevel 1 goto conflict
git push -q
if errorlevel 1 goto pusherr
echo.
echo Готово: изменения на GitHub. На другом ПК нажмите "Получить.cmd".
goto end

:conflict
echo.
echo ВНИМАНИЕ: одни и те же файлы менялись и здесь, и на другом ПК:
git diff --name-only --diff-filter=U
git rebase --abort >nul 2>&1
echo.
echo Ничего не потеряно: ваши правки сохранены на этом ПК, но на GitHub не ушли.
echo Откройте Claude в этой папке и скажите: "конфликт при отправке".
goto end

:pusherr
echo.
echo Не удалось отправить на GitHub: нет интернета или не выполнен вход.
echo Правки сохранены на этом ПК - запустите "Отправить.cmd" позже.
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
