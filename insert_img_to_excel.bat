@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

echo ==========================================================
echo  Вставка картинок в ексель
if not "%~1"=="" echo  %~1
echo ==========================================================
echo.

python insert_img_to_excel.py %*

echo.
pause
