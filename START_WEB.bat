@echo off
chcp 65001 >nul
title DOI CHIEU ACH - Web UI

echo ============================================================
echo  DOI CHIEU ACH - Web UI (truy cap qua trinh duyet)
echo ============================================================
echo.

cd /d "%~dp0"

REM ============================================================
REM BUOC 1: KIEM TRA PYTHON
REM ============================================================
python --version >nul 2>&1
if not errorlevel 1 goto :PYTHON_OK
echo [LOI] Chua co Python. Cai tai: https://www.python.org/downloads/
pause & exit /b 1
:PYTHON_OK
echo [OK] Python san sang.

REM ============================================================
REM BUOC 2: KIEM TRA THU VIEN
REM ============================================================
python -c "import flask, flask_socketio, pandas, pyzipper, xlsxwriter, openpyxl, tqdm, python_calamine" >nul 2>&1
if not errorlevel 1 goto :LIB_OK
echo [INFO] Dang cai thu vien...
python -m pip install -r requirements.txt --quiet
if errorlevel 1 ( echo [LOI] Cai thu vien that bai. & pause & exit /b 1 )
python -m pip install python-calamine --quiet >nul 2>&1
:LIB_OK
echo [OK] Thu vien san sang.

REM ============================================================
REM BUOC 3: CHAY WEB SERVER
REM ============================================================
echo.
echo [INFO] Dang khoi dong Web UI...
echo [INFO] Giu cua so nay mo trong luc su dung. Nhan Ctrl+C de dung.
echo.

REM Doi 2 giay roi tu dong mo trinh duyet
start "" cmd /c "ping -n 3 127.0.0.1 >nul && start http://localhost:8080"

python web_app.py

echo.
echo [INFO] Web UI da dung.
pause
