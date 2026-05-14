@echo off
setlocal
cd /d "%~dp0"

set "TOOL=%~dp0pdf_to_excel.py"
if not exist "%TOOL%" (
    echo [ERROR] Missing pdf_to_excel.py next to this BAT file.
    pause
    exit /b 1
)

if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" "%TOOL%" %*
    goto :after_py
)

where py >nul 2>&1
if %ERRORLEVEL%==0 (
    py -3 "%TOOL%" %*
    goto :after_py
)

where python >nul 2>&1
if %ERRORLEVEL%==0 (
    python "%TOOL%" %*
    goto :after_py
)

where python3 >nul 2>&1
if %ERRORLEVEL%==0 (
    python3 "%TOOL%" %*
    goto :after_py
)

echo [ERROR] Python not found.
echo Install: pip install pdfplumber pandas openpyxl
pause
exit /b 1

:after_py
set "EC=%ERRORLEVEL%"
if %EC% neq 0 (
    echo.
    echo Conversion failed. For table PDFs install: pip install pdfplumber pandas openpyxl
    pause
)
exit /b %EC%
