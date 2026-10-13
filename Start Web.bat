@echo off
cd /d "%~dp0"
title Mengde Oversikt Web

echo.
echo =========================================
echo   Mengde Oversikt Web
echo =========================================
echo.

REM Sjekk at streamlit er installert
where streamlit >nul 2>&1
if errorlevel 1 (
    echo [FEIL] streamlit er ikkje installert.
    echo Kjoer:  pip install streamlit
    echo.
    pause
    exit /b 1
)

REM Sjekk at ngrok er installert
where ngrok >nul 2>&1
if errorlevel 1 (
    echo [FEIL] ngrok er ikkje installert.
    echo Last ned fraa: https://ngrok.com/download
    echo Legg ngrok.exe i same mappe eller i PATH.
    echo.
    pause
    exit /b 1
)

echo Startar Streamlit paa http://localhost:8501 ...
start "Streamlit" cmd /k "cd /d "%~dp0" && streamlit run mengde_oversikt_web.py --server.port 8501 --server.headless true"

echo Ventar paa at Streamlit startar ...
timeout /t 5 /nobreak >nul

echo.
echo =========================================
echo   Kopieer HTTPS-lenka under og send
echo   til kollegaer:
echo =========================================
echo.
ngrok http 8501

echo.
echo ngrok er stoppa. Streamlit kjoerer framleis i det andre vindauget.
pause
