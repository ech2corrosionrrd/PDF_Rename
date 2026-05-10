@echo off
REM Збірка одним exe для Windows 7+: потрібен Python 3.8.x (остання гілка з підтримкою Win7).
REM Встановлення: py -3.8 -m pip install -r requirements.txt -r requirements-dev.txt
cd /d "%~dp0"
py -3.8 -m PyInstaller PDF_Rename_Expert.spec --noconfirm
if errorlevel 1 exit /b 1
echo Готово: dist\PDF_Rename_Expert.exe
