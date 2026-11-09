@echo off
rem Install SX Tools into AutoCAD: junction in ApplicationPlugins (no admin needed).
set "BUNDLE=%APPDATA%\Autodesk\ApplicationPlugins\SX_Tools.bundle"
if not exist "%APPDATA%\Autodesk\ApplicationPlugins" mkdir "%APPDATA%\Autodesk\ApplicationPlugins"
if exist "%BUNDLE%" rmdir "%BUNDLE%"
powershell -NoProfile -Command "New-Item -ItemType Junction -Path $env:BUNDLE -Target '%~dp0.' | Out-Null"
if exist "%BUNDLE%\PackageContents.xml" (echo OK. Restart AutoCAD. Palette opens by itself; command SXT shows it.) else (echo ERROR: junction not created.)
pause
