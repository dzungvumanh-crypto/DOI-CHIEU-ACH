@echo off
chcp 65001 >nul
title DOI CHIEU ACH

echo ============================================================
echo  DOI CHIEU ACH - GL02 vs MIS Hub
echo ============================================================
echo.

cd /d "%~dp0"

REM ============================================================
REM BUOC 1: TIM PYTHON
REM  - Uu tien Python cuc bo (_python\) neu da chay INSTALL.bat
REM  - Du phong: Python he thong neu co san
REM ============================================================
set "PYTHON_EXE="

if exist "%~dp0_python\python.exe" (
    set "PYTHON_EXE=%~dp0_python\python.exe"
    echo [OK] Python cuc bo: _python\python.exe
    goto :CHECK_LIBS
)

python --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_EXE=python"
    echo [OK] Python he thong san sang.
    goto :CHECK_LIBS
)

echo [LOI] Chua co Python!
echo       Hay chay INSTALL.bat truoc.
pause & exit /b 1

REM ============================================================
REM BUOC 2: KIEM TRA THU VIEN
REM ============================================================
:CHECK_LIBS
"%PYTHON_EXE%" -c "import pandas, pyzipper, xlsxwriter, openpyxl, tqdm; import python_calamine" >nul 2>&1
if not errorlevel 1 goto :FIND_DATA

echo [INFO] Thu vien chua du, dang cai...
if exist "%~dp0_setup\packages" (
    "%PYTHON_EXE%" -m pip install ^
        --find-links="%~dp0_setup\packages" ^
        --no-index ^
        -r requirements.txt --quiet
) else (
    "%PYTHON_EXE%" -m pip install -r requirements.txt --quiet
)
if errorlevel 1 (
    echo [LOI] Cai thu vien that bai! Hay chay INSTALL.bat truoc.
    pause & exit /b 1
)

REM ============================================================
REM BUOC 3: XAC DINH THU MUC DU LIEU
REM   - Keo tha FOLDER vao START.bat -> %1 = duong dan folder do
REM   - Double-click truc tiep       -> dung thu muc "input"
REM ============================================================
:FIND_DATA
echo [OK] Thu vien san sang.
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
REM ============================================================
echo.
echo [INFO] Bat dau xu ly...
echo.

"%PYTHON_EXE%" main.py --input "%INPUT_DIR%" --output "output" --date "%NGAY_DC%"

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
