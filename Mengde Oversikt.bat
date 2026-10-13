@echo off
cd /d "%~dp0"
python mengde_oversikt_gui.py
if errorlevel 1 (
    echo.
    echo [FEIL] Python-skriptet krasjet. Trykk en tast for aa lukke.
    pause >nul
)
