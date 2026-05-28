@echo off
setlocal
title PM CV Formatter
mode con cols=62 lines=18
cd /d "%~dp0"

REM ── Use bundled Python runtime (no install needed) ────────────────────────────
set "PY=%~dp0runtime\python.exe"

if not exist "%PY%" (
    echo.
    echo  ERROR: runtime\python.exe not found.
    echo  Please re-download the PM CV Formatter ZIP.
    echo.
    pause
    exit /b 1
)

cls
echo.
echo  +----------------------------------------------------------+
echo   PATRICK MORGAN  ^|  CV Formatter
echo  +----------------------------------------------------------+
echo.

"%PY%" "%~dp0setup.py"

if %errorlevel% neq 0 (
    echo.
    echo  Something went wrong. Please contact Will.
    echo.
    pause
)
