@echo off
chcp 65001 > nul
echo ====================================
echo  DOI CHIEU ACH - WEB UI
echo ====================================
cd /d "%~dp0"
python web_app.py
pause
