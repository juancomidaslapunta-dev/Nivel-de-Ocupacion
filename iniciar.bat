@echo off
REM Iniciador — Windows
cd /d "%~dp0"

where py >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no encontrado. Instalar desde https://www.python.org
    echo Asegurarse de marcar "Add Python to PATH" durante la instalacion.
    pause
    exit /b 1
)

if not exist ".venv\" (
    echo Creando entorno virtual...
    py -m venv .venv
)

call .venv\Scripts\activate.bat

echo Instalando dependencias...
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

echo.
echo Iniciando servidor en http://localhost:8080
echo Ctrl+C para detener
echo.

start "" "http://localhost:8080"
py watcher.py
pause
