@echo off
chcp 65001 >nul
title ساخت Smart ID Finder v6 - فایل اجرایی ویندوز
echo ========================================
echo   ساخت فایل اجرایی Smart ID Finder v6
echo ========================================
echo.

:: بررسی پایتون
python --version >nul 2>&1
if errorlevel 1 (
    echo [خطا] پایتون پیدا نشد. پایتون 3.10 یا 3.12 را نصب کنید.
    pause
    exit /b 1
)

echo [1/4] نصب وابستگی‌ها...
python -m pip install --upgrade pip >nul
python -m pip install PyQt5 openpyxl requests beautifulsoup4 pyinstaller
if errorlevel 1 (
    echo [خطا] نصب کتابخانه‌ها ناموفق بود.
    pause
    exit /b 1
)

echo [2/4] پاک‌سازی ساخت‌های قبلی...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist Smart_ID_Finder_v6.spec del /q Smart_ID_Finder_v6.spec

echo [3/4] ساخت فایل exe با PyInstaller (ممکن است چند دقیقه طول بکشد)...
python -m PyInstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name "Smart_ID_Finder_v6" ^
    --clean ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.styles ^
    --hidden-import=openpyxl.utils ^
    --hidden-import=bs4 ^
    --hidden-import=requests ^
    --hidden-import=PyQt5 ^
    --hidden-import=PyQt5.QtCore ^
    --hidden-import=PyQt5.QtGui ^
    --hidden-import=PyQt5.QtWidgets ^
    --collect-all PyQt5 ^
    Smart_ID_Finder_v6.py

if errorlevel 1 (
    echo.
    echo [خطا] ساخت ناموفق بود. متن خطا را بالا ببینید.
    pause
    exit /b 1
)

echo.
echo [4/4] تمام شد!
echo.
echo فایل اجرایی اینجاست:
echo   dist\Smart_ID_Finder_v6.exe
echo.
echo می‌توانید این فایل را به هر سیستم ویندوزی کپی کنید
echo (نیازی به نصب پایتون روی سیستم مقصد نیست).
echo.
explorer dist
pause
