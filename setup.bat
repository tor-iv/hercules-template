@echo off
REM Carta Cap Table Transformer - Setup Script (Windows)

echo ==================================
echo Carta Cap Table Transformer Setup
echo ==================================
echo.

REM Get the directory where this script is located
cd /d "%~dp0"

REM Check for Python
echo Checking for Python...
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo X Python is not installed or not in PATH.
    echo   Download from: https://www.python.org/downloads/
    echo   IMPORTANT: Check "Add Python to PATH" during installation!
    pause
    exit /b 1
)

python --version
echo.

REM Create virtual environment
echo Creating virtual environment...
if exist ".venv" (
    echo   Virtual environment already exists, skipping creation.
) else (
    python -m venv .venv
    echo Done - Created .venv directory
)

REM Activate virtual environment
echo.
echo Activating virtual environment...
call .venv\Scripts\activate.bat
echo Done - Activated .venv

REM Upgrade pip
echo.
echo Upgrading pip...
python -m pip install --upgrade pip --quiet
echo Done - pip upgraded

REM Install dependencies
echo.
echo Installing dependencies...
pip install -r requirements.txt --quiet
echo Done - Installed: xlwings, pandas, openpyxl, streamlit

REM Install xlwings Excel add-in
echo.
echo Installing xlwings Excel add-in...
echo   (This adds a ribbon tab to Excel for running Python scripts)
xlwings addin install
if %ERRORLEVEL% neq 0 (
    echo   Note: xlwings addin install may require Excel to be closed.
    echo   If it failed, close Excel and run: xlwings addin install
)

echo.
echo ==================================
echo Setup Complete!
echo ==================================
echo.
echo USAGE:
echo.
echo 1. Activate the virtual environment:
echo    .venv\Scripts\activate
echo.
echo 2. Run via command line:
echo    python src\carta_to_cap_table.py ^<carta_export.xlsx^> templates\Cap_Table_Template.xlsx
echo.
echo 3. Run the web interface:
echo    streamlit run src\app.py
echo.
echo 4. For Excel integration, see README.md
echo.
pause
