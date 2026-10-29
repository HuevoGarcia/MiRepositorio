@echo off
setlocal
cd /d "%~dp0"

echo Preparando Monaco con Python 3.14...
if exist "C:\Python314\python.exe" (
    "C:\Python314\python.exe" -m venv .venv
) else (
    py -3.14 -m venv .venv
)
if errorlevel 1 goto error_python

".venv\Scripts\python.exe" -c "import sys, struct; sys.exit(0 if sys.version_info[:2] == (3, 14) and struct.calcsize('P') == 8 else 1)"
if errorlevel 1 goto error_python

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto error_dependencias

echo.
echo Instalacion completada. Ejecuta ejecutar.bat.
echo Genera un archivo SQL para revisar y ejecutar manualmente en SSMS.
pause
exit /b 0

:error_python
echo ERROR: se necesita Python 3.14 de 64 bits con venv y el lanzador py,
echo o instalado en C:\Python314\python.exe.
pause
exit /b 1

:error_dependencias
echo ERROR: no se pudieron instalar las dependencias. Revisa el mensaje anterior
echo y la conexion a Internet. Luego vuelve a ejecutar instalar.bat.
pause
exit /b 1
