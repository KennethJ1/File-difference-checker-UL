@echo off
REM run_setup.bat - Attempts to ensure Python exists (uses winget if available) and installs project dependencies for the current user.
REM Run this from the project root containing requirements.txt and src\

SETLOCAL ENABLEDELAYEDEXPANSION

:: Check for python
python --version >nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    echo Python is already installed.
    goto :install_pkgs
)

:: Try py launcher
py -3 --version >nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    echo Python launcher available (py). Using py as python.
    set PYCMD=py -3
    goto :install_pkgs
)

:: Try winget if present to install Python
nwhere winget >nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    echo winget found - attempting to install Python (may require network access)...
    winget install --id=Python.Python.3 -e --accept-package-agreements --accept-source-agreements
    IF %ERRORLEVEL% EQU 0 (
        echo winget install finished. Please re-open a new terminal if needed and re-run this script.
    ) ELSE (
        echo winget failed or was blocked. Please install Python manually from https://python.org and re-run this script.
    )
    goto :end
) ELSE (
    echo winget not available. Please install Python manually from https://python.org (choose "Install for current user") and re-run this script.
    goto :end
)

:install_pkgs
IF DEFINED PYCMD (
  echo Using PYCMD=%PYCMD%
  %PYCMD% -m pip install --user --upgrade pip setuptools wheel
  %PYCMD% -m pip install --user -r requirements.txt
) ELSE (
  echo Using system python
  python -m pip install --user --upgrade pip setuptools wheel
  python -m pip install --user -r requirements.txt
)

:done
 echo.
 echo Installation finished (or attempted). To run the GUI:
 echo   python -u src\gui\app_window.py
 echo Or open VS Code, set the interpreter to the installed Python, and run.
ENDLOCAL
PAUSE
