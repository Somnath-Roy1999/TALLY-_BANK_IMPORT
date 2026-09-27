@echo off
title ACFC Tally Voucher Generator
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    set "PY=py"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PY=python"
    ) else (
        echo.
        echo Python is not installed or is not in PATH.
        echo Install Python 3 from https://www.python.org/downloads/
        echo Make sure "Add Python to PATH" is selected.
        pause
        exit /b 1
    )
)

%PY% -m pip install openpyxl
if %errorlevel% neq 0 (
    echo.
    echo Could not install openpyxl.
    pause
    exit /b 1
)

%PY% ACFC_Tally_Voucher_Generator.py
if %errorlevel% neq 0 (
    echo.
    echo The program stopped with an error.
    pause
)
