@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Встановлення залежностей...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Не вдалося встановити залежності. Перевірте, що Python є в PATH.
    pause
    exit /b 1
)
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_context_menu.ps1"
echo.
pause
