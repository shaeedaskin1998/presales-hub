@echo off
echo Looking for Presales Hub on port 8000...

set "PID="
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8000 ^| findstr LISTENING') do (
    set "PID=%%a"
)

if defined PID (
    echo Found process ID %PID%. Stopping server...
    taskkill /F /PID %PID% >nul 2>&1
    echo App successfully stopped.
) else (
    echo App is not currently running on port 8000.
)

timeout /t 2 /nobreak >nul
exit