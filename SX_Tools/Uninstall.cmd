@echo off
rem Remove the junction only. Files in this folder stay.
set "BUNDLE=%APPDATA%\Autodesk\ApplicationPlugins\SX_Tools.bundle"
if exist "%BUNDLE%" rmdir "%BUNDLE%"
echo Removed. Restart AutoCAD.
pause
