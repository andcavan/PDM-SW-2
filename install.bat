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
    echo Assicurarsi di spuntare "Add Python to PATH" durante l'installazione.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do (
    echo [OK] %%v
    set PYVER=%%v
)

REM -- Verifica versione minima Python 3.10 --
for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do (
    for /f "tokens=1,2 delims=." %%a in ("%%v") do (
        if %%a LSS 3 (
            echo [ERRORE] Python 3.10 o superiore richiesto. Versione trovata: %%v
            pause
            exit /b 1
        )
        if %%a EQU 3 if %%b LSS 10 (
            echo [ERRORE] Python 3.10 o superiore richiesto. Versione trovata: %%v
            pause
            exit /b 1
        )
    )
)

REM -- Crea virtual environment --
if not exist ".venv\Scripts\python.exe" (
    echo Creazione virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERRORE] Impossibile creare il virtual environment.
        echo Provare: python -m pip install --upgrade virtualenv
        pause
        exit /b 1
    )
    echo [OK] Virtual environment creato.
) else (
    echo [OK] Virtual environment gia' presente.
)

REM -- Aggiorna pip all'interno del venv (critico per PyQt6) --
echo Aggiornamento pip...
.venv\Scripts\python.exe -m pip install --upgrade pip
if errorlevel 1 (
    echo [ATTENZIONE] Aggiornamento pip fallito, si continua comunque...
)
echo [OK] pip aggiornato.

REM -- Installa dipendenze --
echo.
echo Installazione dipendenze (potrebbe richiedere qualche minuto)...
.venv\Scripts\pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERRORE] Installazione dipendenze fallita.
    echo.
    echo Possibili cause:
    echo  - Connessione internet assente o bloccata da firewall/proxy
    echo  - Antivirus che blocca il download
    echo  - Spazio su disco insufficiente
    echo.
    echo Provare manualmente:
    echo   .venv\Scripts\pip install PyQt6==6.10.2 --verbose
    pause
    exit /b 1
)
echo [OK] Dipendenze installate.

REM -- Verifica installazione PyQt6 --
.venv\Scripts\python.exe -c "from PyQt6.QtWidgets import QApplication; print('[OK] PyQt6 verificato')"
if errorlevel 1 (
    echo [ERRORE] PyQt6 non funziona correttamente.
    echo Potrebbe mancare il pacchetto Microsoft Visual C++ Redistributable.
    echo Scaricarlo da: https://aka.ms/vs/17/release/vc_redist.x64.exe
    pause
    exit /b 1
)

REM -- Verifica pywin32 --
.venv\Scripts\python.exe -c "import win32api; print('[OK] pywin32 verificato')"
if errorlevel 1 (
    echo [ATTENZIONE] pywin32 non operativo, tentativo post-install...
    .venv\Scripts\python.exe .venv\Scripts\pywin32_postinstall.py -install 2>nul
)

echo.
echo =========================================
echo   Installazione completata con successo!
echo   Avvia l'applicazione con start.bat
echo =========================================
echo.
pause
