@echo off
echo Starting Delivery Monitor Server...
start "Delivery Monitor Backend" cmd /k "node server.js"
echo Server started successfully. Please do not close the new command window while using the app.
echo You can now access the app at: http://localhost:4000/item_tracker.html
pause
