@echo off
rem wxemogif launcher - double click me
setlocal
cd /d "%~dp0"

set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY (
    where py >nul 2>nul && set "PY=py -3"
)

if not defined PY (
    echo.
    echo [ERROR] Python not found in PATH.
    echo Please install Python 3.8+ and check "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

%PY% "%~dp0wxemogif.py" %*
if errorlevel 1 (
    echo.
    echo [ERROR] wxemogif exited with an error. See messages above.
    pause
)
endlocal
