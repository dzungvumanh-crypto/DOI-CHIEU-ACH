@echo off
chcp 65001 >nul
title DOI CHIEU ACH - Web UI

echo ============================================================
echo  DOI CHIEU ACH - Web UI (truy cap qua trinh duyet)
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
echo.
echo  Neu may nay khong co Internet:
echo    1. Lay goi cai dat tu may khac (da chay CHUAN_BI_CAN_INTERNET.bat)
echo    2. Chay INSTALL.bat mot lan
echo.
echo  Neu may nay CO Internet:
echo    - Tai Python tai: https://www.python.org/downloads/
echo    - Tick "Add Python to PATH" khi cai dat
echo.
pause & exit /b 1

REM ============================================================
REM BUOC 2: KIEM TRA THU VIEN
REM ============================================================
:CHECK_LIBS
"%PYTHON_EXE%" -c "import flask, flask_socketio, pandas, pyzipper, xlsxwriter, openpyxl, tqdm, python_calamine" >nul 2>&1
if not errorlevel 1 goto :START_SERVER

echo [INFO] Thu vien chua du, dang cai...
if exist "%~dp0_setup\packages" (
    REM Che do offline: dung goi da tai
    "%PYTHON_EXE%" -m pip install ^
        --find-links="%~dp0_setup\packages" ^
        --no-index ^
        -r requirements.txt --quiet
) else (
    REM Che do online: tai tu Internet
    "%PYTHON_EXE%" -m pip install -r requirements.txt --quiet
)
if errorlevel 1 (
    echo [LOI] Cai thu vien that bai!
    echo       Hay chay INSTALL.bat truoc.
    pause & exit /b 1
)

REM ============================================================
REM BUOC 3: KHOI DONG WEB SERVER
REM ============================================================
:START_SERVER
echo [OK] Thu vien san sang.
echo.
echo [INFO] Dang khoi dong Web UI...
echo [INFO] Giu cua so nay mo trong luc su dung.
echo [INFO] Nhan Ctrl+C de dung chuong trinh.
echo.

REM Doi 2 giay roi tu dong mo trinh duyet
start "" cmd /c "ping -n 3 127.0.0.1 >nul && start http://localhost:8080"

"%PYTHON_EXE%" web_app.py

echo.
echo [INFO] Web UI da dung.
pause
