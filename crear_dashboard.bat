@echo off
REM Arrastra un archivo de Excel o CSV sobre este .bat para crear su dashboard.
cd /d "%~dp0"
if "%~1"=="" (python dashboard_excel.py) else (python dashboard_excel.py "%~1")
pause
