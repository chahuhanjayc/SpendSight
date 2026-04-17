@echo off
echo.
echo ================================================
echo            SpendSight -- Starting Up
echo ================================================
echo.

:: Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    echo   ERROR: Python is not installed.
    echo.
    echo   Please install it first:
    echo   https://www.python.org/downloads/
    echo.
    echo   Make sure to check "Add Python to PATH"
    echo   during installation.
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version') do echo   %%i found
echo.

:: Install requirements
echo   Checking requirements...
python -m pip install -r requirements.txt --quiet
echo   Requirements ready.
echo.

:: Run the app
echo   Open http://localhost:5000 in your browser
echo   Press Ctrl+C to stop
echo ================================================
echo.
python app.py
pause
