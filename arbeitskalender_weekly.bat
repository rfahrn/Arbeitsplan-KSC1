@echo off
REM ============================================
REM Arbeitskalender - Wöchentliche Generierung
REM Im Windows Taskplaner einrichten:
REM   Trigger: Jeden Freitag um z.B. 14:00
REM   Aktion:  Dieses .bat starten
REM ============================================

REM --- Pfad anpassen wo die Dateien liegen ---
SET WORKDIR=C:\Users\Peter\Arbeitskalender
SET PYTHON=python

cd /d "%WORKDIR%"

echo.
echo   Arbeitskalender-Generator
echo   =========================
echo.
echo   [1] GUI starten (empfohlen)
echo   [2] Kommandozeile (CLI)
echo   [3] Automatisch generieren (ohne Eingabe)
echo.

set /p CHOICE="  Auswahl (1/2/3): "

IF "%CHOICE%"=="1" (
    echo Starte GUI...
    %PYTHON% arbeitskalender_gui.py
) ELSE IF "%CHOICE%"=="2" (
    echo Starte Kommandozeile...
    %PYTHON% arbeitskalender.py
) ELSE IF "%CHOICE%"=="3" (
    echo Generiere automatisch...
    echo [%date% %time%] Auto-Generierung >> log.txt
    %PYTHON% -c "from arbeitskalender import *; import random; m=get_next_monday(); kw=m.isocalendar()[1]; random.seed(kw); s=build_schedule(kw,m); f=f'Arbeitsplan_KW{kw}_{m.strftime(\"%%Y%%m%%d\")}.xlsx'; write_excel(s,f); print(f'Erstellt: {f}')"
    for %%f in (Arbeitsplan_KW*.xlsx) do start "" "%%f"
) ELSE (
    echo Ungueltige Auswahl.
)

pause