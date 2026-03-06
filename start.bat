@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Virtual environment non trovato.
    echo Eseguire install.bat prima di avviare l'applicazione.
    pause
    exit /b 1
)
start "" ".venv\Scripts\pythonw.exe" main.py
