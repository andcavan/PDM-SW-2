@echo off
REM ===========================================================================
REM  test_macro.bat  -  Simula esattamente il comando VBA per debug
REM  Eseguire da qualsiasi posizione per verificare il bridge
REM ===========================================================================

set "APP_DIR=%~dp0.."
set "PYTHON=%APP_DIR%\.venv\Scripts\python.exe"
set "SCRIPT=%APP_DIR%\macros\sw_bridge.py"

echo ===================================================
echo  PDM-SW Bridge - Test Macro
echo ===================================================
echo APP_DIR : %APP_DIR%
echo PYTHON  : %PYTHON%
echo SCRIPT  : %SCRIPT%
echo.

if not exist "%PYTHON%" (
    echo ERRORE: python.exe non trovato in %PYTHON%
    pause
    exit /b 1
)

if not exist "%SCRIPT%" (
    echo ERRORE: sw_bridge.py non trovato
    pause
    exit /b 1
)

echo --- Test 1: azione "open" (apre app PDM) ---
echo Comando: "%PYTHON%" "%SCRIPT%" --action open
echo.

cd /d "%APP_DIR%"
"%PYTHON%" "%SCRIPT%" --action open

echo.
echo Return code: %ERRORLEVEL%
echo.
echo --- Log file ---
if exist "%APP_DIR%\macros\sw_bridge.log" (
    type "%APP_DIR%\macros\sw_bridge.log"
) else (
    echo Nessun log generato!
)
echo.
pause
