@echo off
REM ─────────────────────────────────────────────────────────────────────────
REM Multifamily Demand Index — Auto-Launcher (Windows)
REM
REM Double-click this file to start the app.
REM It will automatically:
REM   1. Create a Python virtual environment (first run only)
REM   2. Install dependencies (first run only)
REM   3. Launch the Streamlit app in your default browser
REM ─────────────────────────────────────────────────────────────────────────

cd /d "%~dp0"

set VENV_DIR=.venv
set PYTHON=

REM ── Find Python 3 ──────────────────────────────────────────────────────
where python3 >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set PYTHON=python3
    goto :found
)
where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set PYTHON=python
    goto :found
)

echo.
echo ERROR: Python 3.9+ is required but not found.
echo Please install Python from https://www.python.org/downloads/
echo Make sure to check "Add Python to PATH" during installation.
echo.
pause
exit /b 1

:found
echo Using: %PYTHON%

REM ── Create virtual environment if needed ────────────────────────────────
if not exist "%VENV_DIR%" (
    echo Creating virtual environment...
    %PYTHON% -m venv %VENV_DIR%
)

REM ── Activate and install ────────────────────────────────────────────────
call %VENV_DIR%\Scripts\activate.bat

REM Re-install whenever requirements.txt changes (hash stored in .installed)
for /f "skip=1 tokens=*" %%H in ('certutil -hashfile requirements.txt MD5 2^>nul') do (
    if not defined REQ_HASH set REQ_HASH=%%H
)
set STORED_HASH=
if exist "%VENV_DIR%\.installed" (
    set /p STORED_HASH=<"%VENV_DIR%\.installed"
)
if "%REQ_HASH%" NEQ "%STORED_HASH%" (
    echo Installing dependencies...
    pip install --upgrade pip -q
    pip install -r requirements.txt -q
    echo %REQ_HASH%> "%VENV_DIR%\.installed"
    echo Dependencies installed successfully.
)
REM ── Suppress Streamlit email prompt ────────────────────────────────────────
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit"
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    (echo [general] & echo email = "") > "%USERPROFILE%\.streamlit\credentials.toml"
)
REM ── Launch ──────────────────────────────────────────────────────────────
echo.
echo Launching Multifamily Demand Index App...
echo (Close this window to stop the app)
echo.

streamlit run app.py --server.headless=false --browser.gatherUsageStats=false

pause
