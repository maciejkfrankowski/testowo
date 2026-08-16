@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal

set "PY="
python --version >nul 2>nul && set "PY=python"
if not defined PY ( py -3 --version >nul 2>nul && set "PY=py -3" )
if not defined PY (
  echo Nie znaleziono Pythona.
  echo Zainstaluj z https://www.python.org/downloads/ i ZAZNACZ "Add Python to PATH".
  pause & exit /b 1
)

%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Brakuje biblioteki openpyxl - instaluje...
  %PY% -m pip install --quiet --upgrade pip
  %PY% -m pip install --quiet -r requirements.txt
  if errorlevel 1 (
    echo.
    echo Instalacja nie powiodla sie. Sprobuj recznie:  %PY% -m pip install openpyxl
    pause & exit /b 1
  )
)

%PY% start.py
pause
endlocal
