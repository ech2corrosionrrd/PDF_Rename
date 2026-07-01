@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if exist ".venv\Scripts\activate.bat" (
    call ".venv\Scripts\activate.bat"
)

where python >nul 2>&1
if errorlevel 1 (
    echo Python не знайдено у PATH. Встановіть Python і позначте "Add to PATH", або активуйте venv.
    pause
    exit /b 1
)

python pdf_rename_expert.py
set "EXITCODE=%ERRORLEVEL%"
if not "%EXITCODE%"=="0" (
    echo.
    echo Помилка запуску. Код виходу: %EXITCODE%
    pause
)
exit /b %EXITCODE%
