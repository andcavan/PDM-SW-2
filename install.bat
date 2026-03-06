@echo off
REM =========================================================
REM  install.bat  -  Installazione PDM-SW su nuovo PC
REM  Richiede Python 3.10+ disponibile nel PATH di sistema
REM =========================================================
setlocal
cd /d "%~dp0"

echo.
echo =========================================
echo   PDM-SW  -  Installazione
echo =========================================
echo.

REM -- Verifica Python --
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Python non trovato nel PATH.
    echo Scaricare Python 3.10+ da https://www.python.org
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo [OK] %%v

REM -- Crea virtual environment --
if not exist ".venv\Scripts\python.exe" (
    echo Creazione virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERRORE] Impossibile creare il virtual environment.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment creato.
) else (
    echo [OK] Virtual environment gia' presente.
)

REM -- Installa dipendenze --
echo Installazione dipendenze (potrebbe richiedere qualche minuto)...
.venv\Scripts\pip install -r requirements.txt
if errorlevel 1 (
    echo [ERRORE] Installazione dipendenze fallita.
    pause
    exit /b 1
)
echo [OK] Dipendenze installate.

echo.
echo =========================================
echo   Installazione completata!
echo   Avvia l'applicazione con start.bat
echo =========================================
echo.
pause
