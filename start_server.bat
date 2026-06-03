@echo off
echo ============================================
echo   Inventory / Delivery Monitor - Startup
echo ============================================
echo.

REM Install dependencies on first run (downloads the fast SQLite engine)
if not exist "node_modules" (
    echo Installing dependencies (first run only)...
    call npm install
    echo.
)

REM Build the SQLite database from your data if it does not exist yet.
REM This is safe to run every time - it skips if the database already has data.
echo Preparing database...
node migrate_to_sqlite.js
echo.

echo Starting server...
start "Inventory Monitor Backend" cmd /k "node server.js"
echo.
echo Server started. Keep the new window open while using the app.
echo Open the app at: http://localhost:4000/item_tracker.html
pause
