@echo off
setlocal

echo Starting Media Point Excel Processor...
docker compose up -d --build

REM Give the container a moment to start
timeout /t 3 >nul

start "" http://localhost:8000/
echo Opened http://localhost:8000/ in your browser.

endlocal

