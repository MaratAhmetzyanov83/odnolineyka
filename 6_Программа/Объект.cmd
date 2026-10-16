@echo off
chcp 65001 >nul
cd /d "%~dp0"
py объект.py %*
if errorlevel 1 pause
