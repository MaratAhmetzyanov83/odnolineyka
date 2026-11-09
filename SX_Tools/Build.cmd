@echo off
rem Rebuild Contents\SXTools.dll (needs .NET 8 SDK and AutoCAD 2025). Close AutoCAD first.
rem Other AutoCAD folder: Build.cmd "D:\...\AutoCAD 2025"
if "%~1"=="" (dotnet build "%~dp0src\SXTools.csproj" -c Release -o "%~dp0Contents") else (dotnet build "%~dp0src\SXTools.csproj" -c Release -o "%~dp0Contents" -p:AcadDir=%1)
pause
