@echo off
REM Google Business Profile Reviews Backup
REM This script runs the Google reviews backup application

cd /d "D:\Git_ExcelKidsHub\google-files\GReviews"

echo ========================================
echo ExcelKidsHub Google Reviews Backup
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from https://www.python.org/
    echo.
    pause
    exit /b 1
)

REM Check if virtual environment exists
if not exist "venv\" (
    echo Virtual environment not found.
    echo Creating virtual environment...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
    echo Virtual environment created successfully.
    echo.
)

REM Activate virtual environment
call venv\Scripts\activate.bat

REM Check if dependencies are installed
pip show google-api-python-client >nul 2>&1
if %errorlevel% neq 0 (
    echo Dependencies not installed.
    echo Installing dependencies...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
    echo Dependencies installed successfully.
    echo.
)

REM Check if credentials exist
if not exist "credentials\client_secret.json" (
    echo ERROR: Google OAuth credentials not found
    echo.
    echo Please follow these steps:
    echo 1. Open credentials\README.txt for detailed instructions
    echo 2. Create OAuth credentials in Google Cloud Console
    echo 3. Download client_secret.json
    echo 4. Place it in the credentials folder
    echo.
    pause
    exit /b 1
)

REM Run the Python script
echo Starting Google reviews backup...
echo.
python google_reviews.py

REM Keep window open to read results
if %errorlevel% neq 0 (
    echo.
    echo Script completed with errors.
) else (
    echo.
    echo Script completed successfully.
)

echo.
pause
