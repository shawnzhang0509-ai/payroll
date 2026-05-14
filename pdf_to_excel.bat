@echo off
setlocal
cd /d "%~dp0"
REM UTF-8 code page: clearer Unicode paths in this window (Chinese Windows defaults to CP936).
chcp 65001 >nul 2>&1
REM UTF-8 mode for Python (helps PDF paths with non-ASCII characters).
set PYTHONUTF8=1

set "TOOL=%~dp0pdf_to_excel.py"
if not exist "%TOOL%" (
    echo [ERROR] Missing pdf_to_excel.py next to this BAT file.
    pause
    exit /b 1
)

REM Avoid non-ASCII in echo lines: cmd misreads UTF-8 batch files under CP936.

REM Double-click runs with no PDF path. Drag-and-drop a .pdf onto this BAT, or pass the path.
if "%~1"=="" (
    echo.
    echo [NO PDF] Do not double-click this BAT alone. It needs the PDF file path.
    echo.
    echo Option A: Drag your .pdf file onto this BAT file and release.
    echo Option B: In cmd, cd to this folder, then run:
    echo   pdf_to_excel.bat "D:\path\to\payslip.pdf"
    echo.
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
    if not "%EC%"=="2" (
        echo Conversion failed. For table PDFs install: pip install pdfplumber pandas openpyxl
    )
    pause
)
exit /b %EC%
