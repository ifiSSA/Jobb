@echo off
cd /d "%~dp0"
title Maalebrev PDF — Lokal

echo.
echo =========================================
echo   Maalebrev PDF til Excel
echo =========================================
echo.

REM Sjekk at Python er installert
where python >nul 2>&1
if errorlevel 1 (
    echo [FEIL] Python er ikkje installert eller ikkje i PATH.
    pause
    exit /b 1
)

REM Installer manglade pakkar automatisk
echo Sjekkar pakkar...
python -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo Installerer streamlit...
    pip install streamlit -q
)
python -c "import pdfplumber" >nul 2>&1
if errorlevel 1 (
    echo Installerer pdfplumber...
    pip install pdfplumber -q
)
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installerer openpyxl...
    pip install openpyxl -q
)
python -c "import pandas" >nul 2>&1
if errorlevel 1 (
    echo Installerer pandas...
    pip install pandas -q
)

echo.
echo Startar — nettlesaren opnar seg automatisk paa http://localhost:8503
echo Lukk dette vindauget for aa stoppe.
echo.

streamlit run "%~dp0pdf_maalebrev_web.py" --server.port 8503 --browser.gatherUsageStats false

pause
