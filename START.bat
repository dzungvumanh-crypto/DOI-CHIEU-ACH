@echo off
chcp 65001 >nul
title DOI CHIEU ACH

echo ============================================================
echo  DOI CHIEU ACH - GL02 vs MIS Hub
echo ============================================================
echo.

cd /d "%~dp0"

REM ============================================================
REM BUOC 1: KIEM TRA PYTHON
REM ============================================================
python --version >nul 2>&1
if not errorlevel 1 goto :PYTHON_OK

echo [!] Chua co Python. Dang tu dong cai dat...
winget --version >nul 2>&1
if not errorlevel 1 (
    echo [INFO] Dang cai Python 3.12 qua winget...
    winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
    set "PATH=%LOCALAPPDATA%\Programs\Python\Python312\;%LOCALAPPDATA%\Programs\Python\Python312\Scripts\;%PATH%"
    python --version >nul 2>&1
    if not errorlevel 1 goto :PYTHON_OK
)
echo [!] Cai dat tu dong that bai.
echo     Tai Python tai: https://www.python.org/downloads/
echo     Tick chon "Add Python to PATH" khi cai dat.
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
python -m pip install -r "%~dp0requirements.txt" --quiet
if errorlevel 1 ( echo [LOI] Cai thu vien that bai. & pause & exit /b 1 )
echo [OK] Thu vien san sang.

:LIB_OK

REM ============================================================
REM BUOC 3: XAC DINH THU MUC DU LIEU
REM   - Keo tha folder vao START.bat  → dung folder do
REM   - Chay truc tiep                → dung thu muc mac dinh
REM ============================================================
if not "%~1"=="" (
    if exist "%~1\" (
        set "INPUT_DIR=%~1"
        echo [OK] Thu muc du lieu: %~1
        goto :ASK_DATE
    )
)

REM Khong co keo tha - dung thu muc mac dinh
if exist "%~dp0file du liẹu" (
    set "INPUT_DIR=%~dp0file du liẹu"
    echo [OK] Thu muc du lieu: %INPUT_DIR%
    goto :ASK_DATE
)

echo [LOI] Chua chon thu muc du lieu!
echo.
echo  Cach 1: Keo tha FOLDER du lieu vao file START.bat
echo  Cach 2: Dat file vao thu muc:  %~dp0file du lieu\
echo.
echo  Thu muc can chua 5 loai file:
echo    - 1 file PDF   : ACH_*_NRT_*_*.pdf
echo    - 1 file GW    : *.xlsx  (co sheet GW)
echo    - 1 file GL02  : GL02*.zip
echo    - 2 file MIS DI : *_DI_*.zip  (ngay T va T-1)
echo    - 2 file MIS DEN: *_DEN_*.zip (ngay T va T-1)
pause & exit /b 1

REM ============================================================
REM BUOC 4: NHAP NGAY DOI CHIEU
REM ============================================================
:ASK_DATE
echo.
for /f "tokens=*" %%d in ('powershell -Command "Get-Date -Format 'dd/MM/yyyy'"') do set TODAY=%%d
echo  Ngay hom nay: %TODAY%
set /p "NGAY_DC=  Nhap ngay doi chieu [Enter = %TODAY%]: "
if "%NGAY_DC%"=="" set "NGAY_DC=%TODAY%"
echo [OK] Ngay doi chieu: %NGAY_DC%

REM ============================================================
REM BUOC 5: CHAY DOI CHIEU
REM ============================================================
echo.
echo [INFO] Bat dau xu ly...
echo.

python "%~dp0main.py" --input "%INPUT_DIR%" --output "%~dp0output" --date "%NGAY_DC%"

echo.
if errorlevel 1 (
    echo ============================================================
    echo  [THAT BAI] Co loi. Xem thong bao phia tren.
    echo ============================================================
) else (
    echo ============================================================
    echo  [THANH CONG] Ket qua xuat vao: %~dp0output\
    echo ============================================================
    start "" "%~dp0output"
)

pause
