@echo off
echo =====================================
echo  jscriptowork test runner
echo =====================================

SET mypath=%~dp0

echo.
echo --- polyfills ---
cscript.exe %mypath%launcher.js %mypath%test-polyfills.js

echo.
echo --- filesystem ---
cscript.exe %mypath%launcher.js %mypath%test-filesystem.js

echo.
echo --- http (requires network) ---
cscript.exe %mypath%launcher.js %mypath%test-http.js

echo.
echo --- ui (HTA windows will flash briefly) ---
cscript.exe %mypath%launcher.js %mypath%test-ui.js

pause
