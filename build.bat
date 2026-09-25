@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Збирає dist\insert_img_to_excel.exe і архів релізу dist\yo-insert-img-to-exel.zip
rem Окреме venv, щоб у exe не потрапили зайві пакети з глобального Python.
set VENV=build\venv
if not exist "%VENV%\Scripts\python.exe" python -m venv "%VENV%" || exit /b 1
"%VENV%\Scripts\python.exe" -m pip install -q -r requirements.txt pyinstaller || exit /b 1
"%VENV%\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --console --name insert_img_to_excel insert_img_to_excel.py || exit /b 1

set PKG=dist\yo-insert-img-to-exel
if exist "%PKG%" rmdir /s /q "%PKG%"
mkdir "%PKG%"
copy /y dist\insert_img_to_excel.exe "%PKG%\" >nul
for %%f in (insert_img_to_excel.bat install.bat uninstall.bat install_context_menu.ps1 README.md) do copy /y %%f "%PKG%\" >nul
powershell -NoProfile -Command "Compress-Archive -Path '%PKG%' -DestinationPath 'dist\yo-insert-img-to-exel.zip' -Force"
echo.
echo Готово: dist\yo-insert-img-to-exel.zip
