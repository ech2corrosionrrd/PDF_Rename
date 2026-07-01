@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo === PDF Rename Expert: збірка під Windows 7 ===
echo Потрібен Python 3.8.x (рекомендовано 3.8.10 x64).
echo.

set "PY38="
where py >nul 2>&1
if not errorlevel 1 (
    py -3.8 -c "import sys" >nul 2>&1
    if not errorlevel 1 set "PY38=py -3.8"
)
if not defined PY38 (
    python -c "import sys; raise SystemExit(0 if sys.version_info[:2]==(3,8) else 1)" >nul 2>&1
    if not errorlevel 1 set "PY38=python"
)

if not defined PY38 (
    echo [Помилка] Не знайдено Python 3.8.
    echo Встановіть Python 3.8.10 з python.org і за потреби увімкніть launcher ^(py^).
    echo https://www.python.org/downloads/release/python-3810/
    pause
    exit /b 1
)

echo Використовується: %PY38%
echo.

if not exist ".venv-win7\Scripts\activate.bat" (
    echo Створення віртуального середовища .venv-win7 ...
    %PY38% -m venv .venv-win7
    if errorlevel 1 (
        echo Не вдалося створити venv.
        pause
        exit /b 1
    )
)

call ".venv-win7\Scripts\activate.bat"
python -m pip install --upgrade pip setuptools wheel
if errorlevel 1 goto :pipfail

echo Встановлення залежностей ^(requirements-win7.txt^) ...
pip install -r requirements-win7.txt
if errorlevel 1 goto :pipfail

echo.
echo Запуск тестів ...
python -m pytest tests -q
if errorlevel 1 (
    echo [Помилка] Тести не пройшли — збірку скасовано.
    pause
    exit /b 1
)

echo.
echo PyInstaller ...
pyinstaller --noconfirm PDF_Rename_Expert.spec
if errorlevel 1 (
    echo [Помилка] PyInstaller завершився з помилкою.
    pause
    exit /b 1
)

echo.
echo Готово: dist\PDF_Rename_Expert.exe
echo rules.json та довідка користувача вшиті у exe ^(PDF_Rename_Expert.spec^).
echo На Windows 7 перевірте наявність VC++ Redistributable та оновлень системи.
pause
exit /b 0

:pipfail
echo [Помилка] pip install не вдався.
pause
exit /b 1
