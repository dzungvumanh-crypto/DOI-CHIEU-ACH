@echo off
chcp 65001 >nul
title DOI CHIEU ACH

echo ============================================================
echo  DOI CHIEU ACH - GL02 vs MIS Hub
echo ============================================================
echo.

REM Chuyen den thu muc chua file bat nay
cd /d "%~dp0"

REM ============================================================
REM BUOC 1: KIEM TRA VA CAI PYTHON
REM ============================================================
python --version >nul 2>&1
if not errorlevel 1 goto :PYTHON_OK

echo [!] Chua co Python. Dang tu dong cai dat...
echo.

REM Thu cai bang winget (Windows 10/11 co san)
winget --version >nul 2>&1
if not errorlevel 1 (
    echo [INFO] Dang cai Python 3.12 qua winget...
    winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
    REM Reload PATH
    set "PATH=%LOCALAPPDATA%\Programs\Python\Python312\;%LOCALAPPDATA%\Programs\Python\Python312\Scripts\;%PATH%"
    set "PATH=%APPDATA%\Python\Python312\Scripts\;%PATH%"
    python --version >nul 2>&1
    if not errorlevel 1 goto :PYTHON_OK
)

REM Neu winget that bai, mo trang tai Python
echo [!] Cai dat tu dong that bai.
echo     Vui long tai Python tai: https://www.python.org/downloads/
echo     - Chon "Windows installer (64-bit)"
echo     - QUAN TRONG: Tick chon "Add Python to PATH" khi cai
echo     - Sau khi cai xong, dong cua so nay va mo lai START.bat
echo.
start https://www.python.org/downloads/
pause
exit /b 1

:PYTHON_OK
echo [OK] Python da san sang.

REM ============================================================
REM BUOC 2: KIEM TRA VA CAI THU VIEN
REM ============================================================
echo [INFO] Kiem tra thu vien Python...
python -c "import pandas, pyzipper, xlsxwriter, calamine, openpyxl" >nul 2>&1
if not errorlevel 1 (
    echo [OK] Thu vien da san sang.
    goto :CHECK_INPUT
)

echo [INFO] Dang cai thu vien can thiet...
python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [LOI] Cai thu vien that bai. Kiem tra ket noi mang.
    pause
    exit /b 1
)
echo [OK] Cai thu vien thanh cong.

REM ============================================================
REM BUOC 3: KIEM TRA THU MUC DU LIEU
REM ============================================================
:CHECK_INPUT
echo.
if not exist "file du liẹu" (
    echo [LOI] Khong tim thay thu muc du lieu!
    echo       Vui long tao thu muc va dat du lieu vao:
    echo       %~dp0file du lieu\
    echo.
    echo       Can co cac file:
    echo         - PDF:  ACH_*_NRT_*_*.pdf
    echo         - GW:   *.xlsx  (co sheet GW)
    echo         - GL02: GL02\*.zip
    echo         - MIS:  MIS_Hub\*.zip
    pause
    exit /b 1
)

REM ============================================================
REM BUOC 4: CHAY DOI CHIEU
REM ============================================================
echo [INFO] Bat dau doi chieu...
echo.

python main.py --input ".\file du liẹu" --output ".\output"

echo.
if errorlevel 1 (
    echo ============================================================
    echo  [THAT BAI] Co loi xay ra. Xem thong bao o tren.
    echo ============================================================
) else (
    echo ============================================================
    echo  [THANH CONG] Ket qua: output\doi_chieu_%date:~6,4%%date:~3,2%%date:~0,2%.xlsx
    echo ============================================================
    echo.
    start "" ".\output"
)

pause
