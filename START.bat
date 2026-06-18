@echo off
chcp 65001 >nul
title DOI CHIEU ACH

echo ============================================================
echo  DOI CHIEU ACH - GL02 vs MIS Hub
echo ============================================================
echo.

REM Chuyen den thu muc chua file bat nay
cd /d "%~dp0"

REM Kiem tra Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [LOI] Khong tim thay Python. Vui long cai dat Python truoc.
    pause
    exit /b 1
)

REM Kiem tra thu muc input
if not exist "file du liẹu" (
    echo [LOI] Khong tim thay thu muc "file du lieu"
    echo       Vui long dat du lieu vao thu muc: %~dp0file du lieu\
    pause
    exit /b 1
)

echo [INFO] Dang chay doi chieu...
echo.

python main.py --input ".\file du liẹu" --output ".\output"

echo.
if errorlevel 1 (
    echo [THAT BAI] Co loi xay ra. Xem thong bao o tren.
) else (
    echo [THANH CONG] Ket qua da duoc xuat vao thu muc: output\
    echo.
    echo Mo thu muc ket qua...
    start "" ".\output"
)

pause
