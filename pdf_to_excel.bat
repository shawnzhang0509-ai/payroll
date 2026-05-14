@echo off
setlocal
cd /d "%~dp0"

set "TOOL=%~dp0pdf_to_excel.py"
if not exist "%TOOL%" (
    echo [ERROR] Missing pdf_to_excel.py next to this BAT file.
    pause
    exit /b 1
)

REM Double-click runs with no args — need a PDF path (e.g. drag PDF onto this BAT).
if "%~1"=="" (
    echo.
    echo 【未指定 PDF】不要只双击本 BAT，否则不知道要转哪个文件。
    echo.
    echo 方式一: 把 PDF 拖到本 BAT 图标上，松开鼠标
    echo 方式二: 先打开 cmd，进入本目录后执行:
    echo   pdf_to_excel.bat "D:\路径\工资单.pdf"
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
