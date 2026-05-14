@echo off
setlocal
REM Cloud 7 Payroll Analyzer - double-click launcher (no EXE build required)
cd /d "%~dp0"

set "APP_SCRIPT=%~dp0payroll_analyzer_v1_2.py"
if not exist "%APP_SCRIPT%" (
    echo [ERROR] Missing payroll_analyzer_v1_2.py next to this BAT file.
    pause
    exit /b 1
)

REM Prefer project virtual environment if present
if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" "%APP_SCRIPT%" %*
    goto :after_py
)

where py >nul 2>&1
if %ERRORLEVEL%==0 (
    py -3 "%APP_SCRIPT%" %*
    goto :after_py
)

where python >nul 2>&1
if %ERRORLEVEL%==0 (
    python "%APP_SCRIPT%" %*
    goto :after_py
)

where python3 >nul 2>&1
if %ERRORLEVEL%==0 (
    python3 "%APP_SCRIPT%" %*
    goto :after_py
)

echo [ERROR] Python not found. Install Python 3.10+ from python.org and enable "Add to PATH".
echo Then install dependencies, for example:
echo   py -3 -m pip install PyQt6 pandas numpy pdfplumber pymupdf openpyxl matplotlib
pause
exit /b 1

:after_py
set "EC=%ERRORLEVEL%"
if %EC% neq 0 (
    echo.
    echo The program exited with an error. If you see ModuleNotFoundError, run pip install for the packages above.
    pause
)
exit /b %EC%
