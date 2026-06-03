@echo off
echo Stopping Delivery Monitor Server (Port 4000)...

REM Find the process listening on port 4000 and kill it
FOR /F "tokens=5" %%T IN ('netstat -ano ^| find "LISTENING" ^| findstr ":4000 "') DO (
    echo Found running server with PID %%T
    taskkill /F /PID %%T
)

echo.
echo Server stopped successfully.
pause
