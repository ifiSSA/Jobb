@echo off
REM Start Mengdeuttak GUI

cd /d "%~dp0"

REM Kjør Python-scriptet som starter GUI-en
python Mengdeuttak_GUI.py

REM Hold vinduet åpent hvis det er feil
if %errorlevel% neq 0 (
    echo.
    echo FEIL: GUI kunne ikke startes. Sjekk at Python er installert.
    pause
)
