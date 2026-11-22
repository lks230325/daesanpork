@echo off
chcp 65001 >nul
title [대산포크 B2B] 한돈 4대 플랫폼 상시 서버 & 외부 HTTPS 배포 기동
cd /d "%~dp0"
echo ==========================================================================
echo  [대산포크 B2B 인텔리전스] 한돈 4대 플랫폼 상시 서버 및 외부 HTTPS 터널 기동
echo  - 백그라운드 자동 복구 슈퍼바이저(FastAPI + Cloudflare HTTPS) 시작 중...
echo  - 정기 크롤링: 월~금 오전 09:00 / 오후 16:00 (자동 누적)
echo ==========================================================================
start "" "C:\Users\limkyungsoon\AppData\Local\Programs\Python\Python312\pythonw.exe" "%~dp0deploy\server_supervisor.py"
timeout /t 4 /nobreak >nul
if exist "%~dp0외부접속_URL_안내.txt" type "%~dp0외부접속_URL_안내.txt"
start "" "http://localhost:8000"
echo 창을 닫아도 서버와 외부 HTTPS 접속 주소, 평일 자동 크롤러는 백그라운드에서 계속 작동합니다.
timeout /t 4 >nul
