@echo off
setlocal enabledelayedexpansion

title Auto Template Generator - Custom API
echo =========================================
echo  Auto Template Generator - Custom API
echo =========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python from https://python.org
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo [OK] Python detected.
python --version
echo.

REM Check if .env file exists
if not exist ".env" (
    echo [WARNING] .env file not found in project root!
    echo.
    echo Please create a .env file with your credentials:
    echo.
    echo YOUR_CLIENT_KEY=your-client-key-here
    echo YOUR_PASS_KEY=your-pass-key-here
    echo ENDPOINT_URL=https://your-endpoint-url.com
    echo YOUR_EMAIL=your-email@example.com
    echo.
    pause
    exit /b 1
)

echo [OK] .env file found.
echo.

REM Create virtual environment if it doesn't exist
if not exist "venv" (
    echo [INFO] Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created.
) else (
    echo [OK] Virtual environment exists.
)
echo.

REM Activate virtual environment
echo [INFO] Activating virtual environment...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo [ERROR] Failed to activate virtual environment.
    echo Try running: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
    pause
    exit /b 1
)
echo [OK] Virtual environment activated.
echo.

REM Install dependencies if needed
if not exist "venv\Lib\site-packages\docx" (
    echo [INFO] Installing dependencies (first time setup)...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies.
        pause
        exit /b 1
    )
    echo [OK] Dependencies installed.
) else (
    echo [OK] Dependencies already installed.
)
echo.

REM Find case directories
echo [INFO] Looking for case directories...
set "case_found="
for /d %%D in (cases\*) do (
    if not defined case_found (
        set "case_found=%%D"
        set "case_name=%%~nD"
    )
)

if not defined case_found (
    echo [ERROR] No case directories found in cases/ folder.
    echo Please create a case directory with a sources/ subfolder containing .txt files.
    pause
    exit /b 1
)

echo [OK] Found case: !case_name!
echo.

REM Ask for model name (optional)
echo [INFO] Enter model name for Custom API (press Enter for default: gpt-4o-mini):
set /p model_name="Model: "
if "!model_name!"=="" set "model_name=gpt-4o-mini"
echo.

REM Run the generator
echo =========================================
echo  Running with Custom API
echo  Case: !case_name!
echo  Model: !model_name!
echo =========================================
echo.

python src\generate_doc.py --case cases\!case_name! --client-type custom --model !model_name!

if errorlevel 1 (
    echo.
    echo [ERROR] Generation failed. Check the error message above.
    pause
    exit /b 1
)

echo.
echo =========================================
echo  SUCCESS! Document generated.
echo  Check the output/ folder.
echo =========================================
pause
endlocal
