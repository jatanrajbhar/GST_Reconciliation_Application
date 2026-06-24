@echo off
echo ============================================================
echo GST Reconciliation Tool - Desktop Application
echo ============================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH!
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Check if dependencies are installed
pip show customtkinter >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    pip install -r requirements_desktop.txt
)

echo.
echo Starting GST Reconciliation Tool...
echo.

python gst_reconciliation_app.py

pause
