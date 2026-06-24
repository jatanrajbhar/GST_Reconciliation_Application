@echo off
echo ============================================================
echo GST Reconciliation Tool - Build Script
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

echo Python found. Installing dependencies...
echo.

REM Install required packages
pip install -r requirements_desktop.txt

echo.
echo Dependencies installed. Starting build process...
echo.

REM Run the build script
python build.py

echo.
echo ============================================================
echo Build process completed!
echo ============================================================
echo.
echo If successful, you can find:
echo   - Portable app: dist\GST Reconciliation Tool\
echo   - Installer:    installer_output\GST_Reconciliation_Tool_Setup.exe
echo.
pause
