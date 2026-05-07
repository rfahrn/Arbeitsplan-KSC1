@echo off
REM ═══════════════════════════════════════════════════════════
REM  KSC Arbeitsplan – Windows Start-Skript
REM ═══════════════════════════════════════════════════════════

echo.
echo ========================================
echo   KSC Arbeitsplan wird gestartet...
echo ========================================
echo.

REM Prüfen ob Docker läuft
docker info >nul 2>&1
if errorlevel 1 (
    echo [FEHLER] Docker laeuft nicht oder ist nicht installiert.
    echo         Bitte Docker Desktop starten und erneut versuchen.
    echo.
    pause
    exit /b 1
)

REM Container starten (baut beim ersten Mal automatisch)
docker compose up -d --build

if errorlevel 1 (
    echo.
    echo [FEHLER] Start fehlgeschlagen. Siehe Meldungen oben.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Fertig!
echo ========================================
echo.
echo   Im Browser oeffnen: http://localhost:8000
echo.
echo   Excel-Dateien landen im Ordner: .\data\output\
echo.
echo   Container stoppen: stop.bat
echo.

REM Browser automatisch oeffnen
timeout /t 3 /nobreak >nul
start http://localhost:8000

pause
