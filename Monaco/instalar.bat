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

".venv\Scripts\python.exe" -c "import pyodbc, sys; sys.exit(0 if 'ODBC Driver 18 for SQL Server' in pyodbc.drivers() else 1)"
if errorlevel 1 goto error_odbc

echo.
echo Instalacion completada. Ejecuta ejecutar.bat.
echo Servidor: DWHNET - Base: Auxiliar - Autenticacion: usuario actual de Windows.
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

:error_odbc
echo ERROR: falta Microsoft ODBC Driver 18 for SQL Server de 64 bits.
echo Instalalo desde la pagina oficial y vuelve a ejecutar instalar.bat:
echo https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server
pause
exit /b 1
