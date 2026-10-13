@echo off
cd /d "%~dp0"
title Manglende poster — PDF til Excel

echo.
echo =========================================
echo   Manglende poster: PDF til Excel
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
echo Opnar vindauget...
echo.

REM Utan argument opnar scriptet GUI-et
python "%~dp0pdf_manglende_poster_til_excel.py"
if errorlevel 1 pause
