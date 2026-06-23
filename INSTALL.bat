@echo off
chcp 65001 >nul
title CAI DAT DOI CHIEU ACH (OFFLINE)
setlocal enabledelayedexpansion

echo ============================================================
echo  CAI DAT DOI CHIEU ACH - CHE DO OFFLINE
echo  (Chi can chay mot lan duy nhat)
echo ============================================================
echo.

cd /d "%~dp0"

REM ── Kiem tra da cai day du chua ──────────────────────────────
if exist "_python\python.exe" (
    "_python\python.exe" -c "import flask, pandas, pyzipper" >nul 2>&1
    if not errorlevel 1 (
        echo [OK] Chuong trinh da duoc cai dat truoc do.
        echo      Chay START_WEB.bat de su dung.
        pause & exit /b 0
    )
)

REM ── Kiem tra goi cai dat ─────────────────────────────────────
if not exist "_setup\python-embed.zip" (
    echo [LOI] Khong tim thay goi cai dat!
    echo.
    echo  Ban can chay CHUAN_BI_CAN_INTERNET.bat tren may CO Internet
    echo  de chuan bi goi cai dat truoc, roi copy sang may nay.
    echo.
    pause & exit /b 1
)

REM ── BUOC 1: Giai nen Python 3.12 ─────────────────────────────
echo [1/4] Dang giai nen Python 3.12 Embedded...
if exist "_python" rd /s /q "_python"
mkdir "_python"
powershell -NoProfile -Command ^
    "Expand-Archive -Path '_setup\python-embed.zip' -DestinationPath '_python' -Force"
if not exist "_python\python.exe" (
    echo [LOI] Giai nen Python that bai!
    pause & exit /b 1
)
echo [OK] Python 3.12 tai: %~dp0_python\

REM ── BUOC 2: Bat site-packages ─────────────────────────────────
echo [2/4] Kich hoat che do site-packages...
REM Uncomment dong "#import site" trong python312._pth
powershell -NoProfile -Command ^
    "(Get-Content '_python\python312._pth') -replace '#import site','import site' | Set-Content '_python\python312._pth' -Encoding ASCII"
REM Tao thu muc site-packages
if not exist "_python\Lib" mkdir "_python\Lib"
mkdir "_python\Lib\site-packages" 2>nul
echo [OK]

REM ── BUOC 3: Cai pip ──────────────────────────────────────────
echo [3/4] Cai pip vao Python cuc bo...
set "PIP_WHL="
for %%f in ("_setup\packages\pip-*.whl") do set "PIP_WHL=%%~f"

if defined PIP_WHL (
    echo     Giai nen: !PIP_WHL!
    powershell -NoProfile -Command ^
        "Expand-Archive -Path '!PIP_WHL!' -DestinationPath '_python\Lib\site-packages' -Force"
    echo [OK] pip da cai.
) else if exist "_setup\get-pip.py" (
    echo     Dung get-pip.py...
    "_python\python.exe" "_setup\get-pip.py" ^
        --no-index --find-links="_setup\packages" --quiet
    if errorlevel 1 (
        echo [LOI] Cai pip that bai!
        pause & exit /b 1
    )
    echo [OK] pip da cai.
) else (
    echo [LOI] Khong tim thay pip wheel trong _setup\packages\
    echo       Hay chay lai CHUAN_BI_CAN_INTERNET.bat tren may CO Internet.
    pause & exit /b 1
)

REM Kiem tra pip hoat dong
"_python\python.exe" -m pip --version >nul 2>&1
if errorlevel 1 (
    echo [LOI] pip khong hoat dong sau khi cai!
    pause & exit /b 1
)

REM ── BUOC 4: Cai thu vien ────────────────────────────────────
echo [4/4] Cai tat ca thu vien (flask, pandas, pyzipper...)...
"_python\python.exe" -m pip install ^
    --find-links="_setup\packages" ^
    --no-index ^
    -r requirements.txt ^
    --quiet
if errorlevel 1 (
    echo.
    echo [LOI] Cai thu vien that bai! Chi tiet:
    "_python\python.exe" -m pip install ^
        --find-links="_setup\packages" ^
        --no-index ^
        -r requirements.txt
    pause & exit /b 1
)
echo [OK] Tat ca thu vien da cai thanh cong.

REM ── Kiem tra lan cuoi ────────────────────────────────────────
"_python\python.exe" -c ^
    "import flask, flask_socketio, pandas, pyzipper, xlsxwriter, openpyxl, tqdm, python_calamine; print('Kiem tra OK')"
if errorlevel 1 (
    echo [CANH BAO] Co the mot so thu vien chua dung. Xem lo tren.
) else (
    echo [OK] Kiem tra toan bo thu vien - thanh cong.
)

echo.
echo ============================================================
echo  CAI DAT THANH CONG!
echo  Bay gio chay START_WEB.bat de su dung chuong trinh.
echo ============================================================
pause
