@echo off
chcp 65001 >nul
title [대산포크 B2B] 배포 서버 상태 확인 및 재점검
cd /d "%~dp0"
start "" "C:\Users\limkyungsoon\AppData\Local\Programs\Python\Python312\pythonw.exe" "%~dp0deploy\server_supervisor.py"
timeout /t 3 /nobreak >nul
if exist "%~dp0외부접속_URL_안내.txt" (
    type "%~dp0외부접속_URL_안내.txt"
) else (
    echo [안내] 서버 및 터널 상태를 초기화하는 중입니다. 5초 후 다시 확인해 주세요.
)
echo.
echo [윈도우 작업 스케줄러 등록 현황]
schtasks /Query /TN "DaesanPork_Handon_Watchdog" 2>nul
schtasks /Query /TN "DaesanPork_Handon_Crawl_0900" 2>nul
schtasks /Query /TN "DaesanPork_Handon_Crawl_1600" 2>nul
echo.
pause
