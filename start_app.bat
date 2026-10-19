@echo off
cd /d "C:\presales-hub"

echo Checking if app is already running...
netstat -ano | findstr :8000 | findstr LISTENING >nul
if %errorlevel% equ 0 (
    echo App is already running on port 8000.
    start http://localhost:8000
    exit /b
)

echo Starting Presales Hub in the background...
powershell -WindowStyle Hidden -Command "Start-Process -FilePath 'C:\presales-hub\.venv\Scripts\python.exe' -ArgumentList '-m uvicorn main:app --host 0.0.0.0 --port 8000' -WindowStyle Hidden"

timeout /t 2 /nobreak >nul
start http://localhost:8000
exit