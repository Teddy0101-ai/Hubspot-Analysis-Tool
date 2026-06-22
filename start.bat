@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================================
echo   Hubspot Analysis Tool  -  starting up
echo ============================================================
echo.

REM --- 1) Make sure Python is available -----------------------
where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python was not found on this computer.
  echo.
  echo   Please install Python 3.9 or newer from:
  echo       https://www.python.org/downloads/
  echo   During install, TICK the box "Add Python to PATH".
  echo.
  pause
  exit /b 1
)

REM --- 2) Create a private virtual environment on first run ---
if not exist ".venv\Scripts\python.exe" (
  echo [SETUP] Creating the virtual environment ^(first run only^)...
  python -m venv .venv
  if errorlevel 1 (
    echo [ERROR] Could not create the virtual environment.
    pause
    exit /b 1
  )
)

set "PYEXE=.venv\Scripts\python.exe"

REM --- 3) Install required packages on first run --------------
if not exist ".venv\.installed" (
  echo [SETUP] Installing required packages ^(first run only - needs internet^)...
  echo         This can take a minute or two. Please wait.
  "%PYEXE%" -m pip install --upgrade pip >nul 2>nul
  "%PYEXE%" -m pip install -r requirements.txt
  if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install packages. Check your internet connection and try again.
    pause
    exit /b 1
  )
  echo done> ".venv\.installed"
  echo [OK] Setup complete.
  echo.
)

REM --- 4) Launch the web app ----------------------------------
echo [OK] Starting the tool. Your browser will open automatically.
echo      Keep this window open while you use the tool.
echo      Close this window when you are finished.
echo.
"%PYEXE%" app.py

echo.
echo The tool has stopped.
pause
