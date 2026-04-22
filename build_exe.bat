@echo off
chcp 65001 >nul
echo ============================================
echo   DONG GOI TOOL CHIA FILE EXCEL (.exe)
echo ============================================
echo.

:: Kiểm tra Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [LOI] Khong tim thay Python! Vui long cai dat Python 3.8+ tu python.org
    pause
    exit /b 1
)

echo [1/3] Dang cai dat thu vien can thiet...
pip install xlrd==1.2.0 xlwt==1.3.0 openpyxl==3.1.2 pyinstaller --quiet

echo.
echo [2/3] Dang dong goi thanh file .exe...
echo     (Qua trinh nay mat khoang 1-2 phut)
echo.

pyinstaller --noconfirm ^
    --onefile ^
    --windowed ^
    --name "ChiaFileExcel" ^
    --icon=NONE ^
    --add-data "requirements.txt;." ^
    --hidden-import xlrd ^
    --hidden-import xlwt ^
    --hidden-import xlutils ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.cell ^
    --hidden-import openpyxl.cell._writer ^
    --hidden-import openpyxl.workbook ^
    chia_file_excel.py

echo.
if exist "dist\ChiaFileExcel.exe" (
    echo ============================================
    echo   THANH CONG!
    echo ============================================
    echo.
    echo   File .exe: dist\ChiaFileExcel.exe
    echo.
    echo   Ban co the copy file nay di bat ky dau
    echo   va chay truc tiep, khong can cai Python!
    echo ============================================
    
    :: Copy ra thu muc chinh cho tien
    copy "dist\ChiaFileExcel.exe" "ChiaFileExcel.exe" >nul 2>&1
    echo.
    echo   Da copy ra: ChiaFileExcel.exe
) else (
    echo [LOI] Dong goi that bai! Vui long kiem tra log phia tren.
)

echo.
pause
