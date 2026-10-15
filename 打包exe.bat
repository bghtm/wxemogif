@echo off
rem Build wxemogif.exe - double click me
setlocal
cd /d "%~dp0"

set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY (
    where py >nul 2>nul && set "PY=py -3"
)
if not defined PY (
    echo [ERROR] Python not found in PATH.
    pause
    exit /b 1
)

%PY% "%~dp0build_exe.py" %*
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed. See messages above.
)
pause
endlocal
