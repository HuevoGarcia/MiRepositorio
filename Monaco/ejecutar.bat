@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Primero ejecuta instalar.bat para preparar Python y las dependencias.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" monaco.py %*
set "MONACO_EXIT_CODE=%ERRORLEVEL%"
pause
exit /b %MONACO_EXIT_CODE%
