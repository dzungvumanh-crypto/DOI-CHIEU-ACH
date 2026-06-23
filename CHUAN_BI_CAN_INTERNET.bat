@echo off
chcp 65001 >nul
title CHUAN BI GOI OFFLINE - DOI CHIEU ACH
setlocal

echo ============================================================
echo  BUOC CHUAN BI (CHAY TREN MAY CO INTERNET)
echo  Tai Python 3.12 + tat ca thu vien vao thu muc _setup\
echo  Sau khi xong: copy TOAN BO thu muc du an sang may dich
echo ============================================================
echo.

cd /d "%~dp0"

REM ── Kiem tra Internet ────────────────────────────────────────
ping -n 1 -w 3000 pypi.org >nul 2>&1
if errorlevel 1 (
    echo [LOI] Khong co ket noi Internet!
    echo       File nay phai duoc chay tren may CO Internet.
    pause & exit /b 1
)

REM ── Kiem tra pip ─────────────────────────────────────────────
pip --version >nul 2>&1
if errorlevel 1 (
    echo [LOI] Khong tim thay pip. Hay cai Python truoc.
    pause & exit /b 1
)

REM ── Don dep goi cu ───────────────────────────────────────────
if exist "_setup" (
    echo [INFO] Xoa goi cu trong _setup\ ...
    rd /s /q "_setup"
)
mkdir "_setup"
mkdir "_setup\packages"

echo.
echo [1/3] Dang tai Python 3.12.10 Embedded AMD64 (~12 MB)...
echo       Vui long doi...
powershell -NoProfile -Command ^
    "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip' -OutFile '_setup\python-embed.zip' -UseBasicParsing"
if not exist "_setup\python-embed.zip" (
    echo [LOI] Tai Python that bai! Kiem tra lai ket noi Internet.
    pause & exit /b 1
)
echo [OK] Python 3.12.10 embedded.

echo.
echo [2/3] Dang tai tat ca thu vien Python (~80-150 MB)...
echo       Vui long doi (co the mat 3-10 phut)...

pip download -r requirements.txt -d "_setup\packages" --prefer-binary --quiet
if errorlevel 1 (
    echo [CANH BAO] Lan 1 that bai. Thu lai khong filter binary...
    pip download -r requirements.txt -d "_setup\packages" --quiet
    if errorlevel 1 (
        echo [LOI] Tai thu vien that bai!
        pause & exit /b 1
    )
)

REM Dam bao pip co trong packages de boot tren may dich
pip download pip setuptools -d "_setup\packages" --prefer-binary --quiet

echo [OK] Thu vien.

echo.
echo [3/3] Dang tai pip bootstrapper (du phong)...
powershell -NoProfile -Command ^
    "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '_setup\get-pip.py' -UseBasicParsing" >nul 2>&1
if exist "_setup\get-pip.py" (
    echo [OK] get-pip.py.
) else (
    echo [CANH BAO] Khong tai duoc get-pip.py - co the bo qua.
)

REM ── Thong ke ket qua ─────────────────────────────────────────
echo.
set WHL_COUNT=0
for %%f in ("_setup\packages\*.whl") do set /a WHL_COUNT+=1
for %%f in ("_setup\packages\*.tar.gz") do set /a WHL_COUNT+=1

echo ============================================================
echo  HOAN TAT! Da chuan bi %WHL_COUNT% goi thu vien.
echo.
echo  Buoc tiep theo:
echo   1. Copy TOAN BO thu muc nay sang may dich (USB / mang noi bo)
echo   2. Tren may dich: double-click INSTALL.bat (chi can 1 lan)
echo   3. Sau do: double-click START_WEB.bat moi ngay
echo ============================================================
pause
