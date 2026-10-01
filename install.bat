@echo off
REM ---------------------------------------------------------------------------
REM pptx_translator installer for Windows
REM
REM How to use: double-click this file (install.bat), OR open Command Prompt in
REM this folder and type:  install.bat
REM
REM It creates a private, self-contained Python environment in a ".venv" folder
REM and installs everything the tool needs. It does not change the rest of your
REM computer.
REM ---------------------------------------------------------------------------
setlocal
cd /d "%~dp0"

echo.
echo   pptx_translator - setup (Windows)
echo   ---------------------------------

REM 1. Find Python.
where py >nul 2>nul
if %ERRORLEVEL%==0 (
  set "PY=py -3"
) else (
  where python >nul 2>nul
  if %ERRORLEVEL%==0 (
    set "PY=python"
  ) else (
    echo.
    echo   Python 3 was not found on your computer.
    echo   Please install it from https://www.python.org/downloads/
    echo   During install, TICK the box "Add Python to PATH", then run this again.
    pause
    exit /b 1
  )
)

%PY% --version

REM 2. Create the virtual environment.
echo   Creating a private Python environment in .venv ...
%PY% -m venv .venv

REM 3. Install dependencies.
echo   Installing required packages (this can take a minute) ...
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
  echo.
  echo   Something went wrong installing packages. Please check the messages above.
  pause
  exit /b 1
)

REM 4. Self-test.
echo   Verifying the installation ...
".venv\Scripts\python.exe" -m pptx_translator --help >nul

echo.
echo   All set!  To use the tool, run these commands from this folder:
echo.
echo     Extract text:
echo       .venv\Scripts\python.exe -m pptx_translator extract "MyDeck.pptx" -o "MyDeck.json"
echo.
echo     Patch translations back in:
echo       .venv\Scripts\python.exe -m pptx_translator patch "MyDeck.pptx" "MyDeck.json" -o "MyDeck_translated.pptx"
echo.
pause
