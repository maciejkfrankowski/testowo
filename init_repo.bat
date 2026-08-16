@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal

if "%~1"=="" (
  echo Podaj adres pustego repozytorium, np:
  echo    init_repo.bat https://github.com/macief/biznesradar-mapowanie.git
  pause & exit /b 1
)

call :srodowisko || goto :koniec_blad

echo.
echo == kontrola przed wypchnieciem ==
%PY% slownik_csv.py  || goto :koniec_blad
%PY% test_slownik.py || goto :koniec_blad

where git >nul 2>nul
if errorlevel 1 (
  echo.
  echo Nie znaleziono gita. Zainstaluj z https://git-scm.com/download/win i uruchom ponownie.
  goto :koniec_blad
)

echo.
git init
git add -A
git status --short
echo.
set /p ODP="Wypchnac powyzsze na %~1 ? [t/N] "
if /i not "%ODP%"=="t" (echo Przerwano. & goto :koniec)

git commit -m "Mapowanie sprawozdan na standard biznesradar"
git branch -M main
git remote remove origin 2>nul
git remote add origin %~1
git push -u origin main
echo.
echo Gotowe.
goto :koniec


:srodowisko
rem --- ustala czym uruchamiac Pythona i dociaga brakujace biblioteki ---
set "PY="
python --version >nul 2>nul && set "PY=python"
if not defined PY ( py -3 --version >nul 2>nul && set "PY=py -3" )
if not defined PY (
  echo.
  echo Nie znaleziono Pythona.
  echo Zainstaluj z https://www.python.org/downloads/ i ZAZNACZ "Add Python to PATH".
  exit /b 1
)
for /f "delims=" %%v in ('%PY% --version 2^>^&1') do echo Python: %%v

%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Brakuje biblioteki openpyxl - instaluje...
  %PY% -m pip install --quiet --upgrade pip
  %PY% -m pip install --quiet -r requirements.txt
  if errorlevel 1 (
    echo.
    echo Instalacja nie powiodla sie. Sprobuj recznie:
    echo    %PY% -m pip install openpyxl
    exit /b 1
  )
  echo Zainstalowano.
)
exit /b 0

:koniec_blad
echo.
echo Przerwano z powodu bledu.
:koniec
pause
endlocal
