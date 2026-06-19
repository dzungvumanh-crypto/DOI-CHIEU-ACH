@echo off
chcp 65001 >nul
title DOI CHIEU ACH

echo ============================================================
echo  DOI CHIEU ACH - GL02 vs MIS Hub
echo ============================================================
echo.

REM Chuyen ve thu muc chua START.bat (dung absolute path noi bo)
cd /d "%~dp0"

REM ============================================================
REM BUOC 1: KIEM TRA PYTHON
REM ============================================================
python --version >nul 2>&1
if not errorlevel 1 goto :PYTHON_OK

echo [!] Chua co Python. Dang tu dong cai...
winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements >nul 2>&1
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312\;%LOCALAPPDATA%\Programs\Python\Python312\Scripts\;%PATH%"
python --version >nul 2>&1
if not errorlevel 1 goto :PYTHON_OK

echo [!] Cai tu dong that bai. Tai Python tai:
echo     https://www.python.org/downloads/
echo     (Tick "Add Python to PATH" khi cai)
start https://www.python.org/downloads/
pause & exit /b 1

:PYTHON_OK
echo [OK] Python san sang.

REM ============================================================
REM BUOC 2: KIEM TRA THU VIEN
REM ============================================================
python -c "import pandas, pyzipper, xlsxwriter, calamine, openpyxl" >nul 2>&1
if not errorlevel 1 goto :LIB_OK
echo [INFO] Dang cai thu vien...
python -m pip install -r requirements.txt --quiet
if errorlevel 1 ( echo [LOI] Cai thu vien that bai. & pause & exit /b 1 )
:LIB_OK
echo [OK] Thu vien san sang.

REM ============================================================
REM BUOC 3: XAC DINH THU MUC DU LIEU
REM   - Keo tha FOLDER vao START.bat  -> %1 = duong dan folder do
REM   - Chay truc tiep double-click   -> dung thu muc "input"
REM ============================================================
echo.
set "INPUT_DIR="

if not "%~1"=="" (
    if exist "%~1\" (
        set "INPUT_DIR=%~1"
        echo [OK] Thu muc du lieu: %~1
        goto :ASK_DATE
    )
)

if exist "input\" (
    set "INPUT_DIR=input"
    echo [OK] Dung thu muc mac dinh: input\
    goto :ASK_DATE
)

echo [LOI] Khong tim thay du lieu!
echo.
echo  === CACH SU DUNG ===
echo.
echo  Cach 1 - Keo tha:
echo    Keo FOLDER du lieu tha thang vao file START.bat
echo.
echo  Cach 2 - Thu muc mac dinh:
echo    Dat tat ca file vao thu muc:  input\
echo.
echo  File can co trong folder:
echo    PDF    : ACH_*_NRT_*_*.pdf
echo    GW     : *.xlsx  (co sheet GW)
echo    GL02   : GL02*.zip
echo    MIS DI : *_DI_*.zip  (file T va T-1)
echo    MIS DEN: *_DEN_*.zip (file T va T-1)
echo.
pause & exit /b 1

REM ============================================================
REM BUOC 4: NHAP NGAY DOI CHIEU
REM ============================================================
:ASK_DATE
echo.
for /f "tokens=*" %%d in ('powershell -NoProfile -Command "Get-Date -Format 'dd/MM/yyyy'"') do set "TODAY=%%d"
echo  Ngay hom nay: %TODAY%
set /p "NGAY_DC=  Nhap ngay doi chieu [Enter = %TODAY%]: "
if "%NGAY_DC%"=="" set "NGAY_DC=%TODAY%"
echo [OK] Ngay: %NGAY_DC%

REM ============================================================
REM BUOC 5: CHAY DOI CHIEU
REM   Dung relative path (input, output) vi da cd vao thu muc goc
REM   Tranh truyen path Unicode qua command line
REM ============================================================
echo.
echo [INFO] Bat dau xu ly...
echo.

python main.py --input "%INPUT_DIR%" --output "output" --date "%NGAY_DC%"

echo.
if errorlevel 1 (
    echo ============================================================
    echo  [THAT BAI] Co loi - xem thong bao phia tren.
    echo ============================================================
) else (
    echo ============================================================
    echo  [THANH CONG] Ket qua da xuat vao: output\
    echo ============================================================
    start "" "output"
)

pause
