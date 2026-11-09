@echo off
rem Copy fresh .lsp/.dcl/.csv/docs from ..\2_AutoCAD and ..\4_Instructions into this folder.
py "%~dp0sync_from_project.py" || python "%~dp0sync_from_project.py"
pause
