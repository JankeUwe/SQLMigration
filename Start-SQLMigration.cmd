@echo off
:: ============================================================
:: SQLMigration - Starter
:: ============================================================
:: Kopiert das Tool einmalig nach C:\ProgramData\SQLMigration
:: und startet es dann als Administrator (UAC).
::
:: Warum ProgramData?
::   - Nach UAC-Elevation ist das Netzlaufwerk (W:\) nicht mehr
::     erreichbar (Elevation laeuft unter lokalem System-Kontext)
::   - ProgramData ist AppLocker/AV-unbedenklich
::   - Nicht von Cleanup-Scripts betroffen
::
:: Verwendung:
::   Doppelklick vom Share oder UNC-Pfad genuegt.
::   Kein manuelles Kopieren durch den Admin noetig.
::
:: Rolle (optional als Parameter):
::   Start-SQLMigration.cmd Source   -> startet direkt als Quellserver
::   Start-SQLMigration.cmd Target   -> startet direkt als Zielserver
::   Start-SQLMigration.cmd          -> Rollenauswahl per Dialog
:: ============================================================
setlocal EnableDelayedExpansion

set "SRCDIR=%~dp0"
set "LOCALDIR=%ProgramData%\SQLMigration"
set "LOCALPS=%LOCALDIR%\SQL-Migration.ps1"
set "ROLE=%~1"

echo.
echo  SQLMigration - Vorbereitung
echo  ============================================================
echo  Quelle : %SRCDIR%
echo  Ziel   : %LOCALDIR%
if not "%ROLE%"=="" (
    echo  Rolle  : %ROLE%
)
echo.

:: Zielverzeichnis anlegen falls nicht vorhanden
if not exist "%LOCALDIR%" (
    mkdir "%LOCALDIR%"
    if errorlevel 1 (
        echo  FEHLER: Verzeichnis konnte nicht angelegt werden: %LOCALDIR%
        echo  Bitte Script als Administrator ausfuehren.
        pause
        exit /b 1
    )
)

:: Alle Dateien kopieren (ueberschreibt vorhandene - immer aktuelle Version)
:: /E = inkl. Unterverzeichnisse (modules\, config\)
xcopy /Y /Q /E "%SRCDIR%." "%LOCALDIR%\" >nul 2>&1
if errorlevel 1 (
    echo  FEHLER: Kopieren fehlgeschlagen.
    echo  Pruefen: Lesezugriff auf Quelle, Schreibzugriff auf Ziel.
    pause
    exit /b 1
)

echo  Dateien bereit - starte als Administrator ...
echo.

:: Rolle als Argument weitergeben wenn angegeben
if "%ROLE%"=="" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
        "Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%LOCALPS%""' -Verb RunAs"
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
        "Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%LOCALPS%"" -Role %ROLE%' -Verb RunAs"
)

endlocal
