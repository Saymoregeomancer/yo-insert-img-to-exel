@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem з exe Python і залежності не потрібні
if exist "%~dp0insert_img_to_excel.exe" goto menu

echo Встановлення залежностей...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Не вдалося встановити залежності. Перевірте, що Python є в PATH.
    pause
    exit /b 1
)
echo.

:menu
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_context_menu.ps1"
echo.
pause
