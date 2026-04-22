@echo off
chcp 65001 >nul
echo ============================================
echo   TOOL CHIA FILE EXCEL - Cai dat va Chay
echo ============================================
echo.

:: Kiểm tra Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [LOI] Khong tim thay Python! Vui long cai dat Python 3.8+ tu python.org
    pause
    exit /b 1
)

echo [1/2] Dang cai dat thu vien can thiet...
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [CANH BAO] Co loi khi cai dat thu vien, thu lai voi pip3...
    pip3 install -r requirements.txt --quiet
)

echo.
echo [2/2] Dang khoi dong tool...
echo.
python chia_file_excel.py

pause
