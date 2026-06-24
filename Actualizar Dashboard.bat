@echo off
chcp 65001 >nul
cd /d "%~dp0"
where python >nul 2>nul || (
  echo No se encontro Python. Instalalo desde https://www.python.org/downloads/  ^(marca "Add to PATH"^).
  pause
  exit /b
)
python -c "import openpyxl" >nul 2>nul || (
  echo Instalando libreria openpyxl ^(solo la primera vez^)...
  python -m pip install openpyxl
)
python "actualizar_dashboard.py"
if errorlevel 1 pause
