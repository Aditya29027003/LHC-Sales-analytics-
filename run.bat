@echo off
REM Launches the Sales Analytics Tracker in the default browser.
REM Double-click this file to start.

cd /d "%~dp0"

where streamlit >nul 2>nul
if errorlevel 1 (
    echo Streamlit not found. Installing required packages...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Failed to install dependencies. Please install Python 3.10+ and re-run.
        pause
        exit /b 1
    )
)

echo Starting Sales Analytics Tracker...
echo Press Ctrl+C in this window to stop the app when you are done.
echo.

python -m streamlit run tracker.py
pause
