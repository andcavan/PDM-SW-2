@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Virtual environment non trovato.
    echo Eseguire install.bat prima di avviare l'applicazione.
    pause
    exit /b 1
)

REM Avvia in background senza console. In caso di crash, l'errore viene
REM scritto in error.txt nella stessa cartella dell'applicazione.
start "" ".venv\Scripts\pythonw.exe" main.py 2>error.txt
