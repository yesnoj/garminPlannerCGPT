@echo off
REM ============================================================
REM  Crea dist\GarminTrainingPlanner.exe (un unico file, niente Python da installare)
REM  Uso: doppio clic su questo file, oppure da terminale nella cartella del progetto.
REM ============================================================
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (set PY=py) else (set PY=python)

echo.
echo [1/3] Ambiente di build pulito (.venv-build)...
if not exist .venv-build (%PY% -m venv .venv-build || goto :errore)
call .venv-build\Scripts\activate.bat || goto :errore

echo.
echo [2/3] Installazione dipendenze...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt pyinstaller || goto :errore

echo.
echo [3/3] Creazione dell'eseguibile (qualche minuto)...
python -m PyInstaller --noconfirm --clean GarminTrainingPlanner.spec || goto :errore

echo.
echo ============================================================
echo  FATTO:  dist\GarminTrainingPlanner.exe
echo ============================================================
explorer dist
pause
exit /b 0

:errore
echo.
echo *** Errore durante la build: controlla i messaggi qui sopra. ***
pause
exit /b 1
