@echo off
echo =====================================
echo  jscriptowork test runner
echo =====================================

SET mypath=%~dp0

echo.
echo ============ offline suites =========

echo.
echo --- core ---
cscript.exe %mypath%launcher.js %mypath%test-core.js

echo.
echo --- polyfills ---
cscript.exe %mypath%launcher.js %mypath%test-polyfills.js

echo.
echo --- minitest ---
cscript.exe %mypath%launcher.js %mypath%test-minitest.js

echo.
echo --- crypto ---
cscript.exe %mypath%launcher.js %mypath%test-crypto.js

echo.
echo --- minimist ---
cscript.exe %mypath%launcher.js %mypath%test-minimist.js

echo.
echo ============ disk suites ============

echo.
echo --- system ---
cscript.exe %mypath%launcher.js %mypath%test-system.js

echo.
echo --- filesystem ---
cscript.exe %mypath%launcher.js %mypath%test-filesystem.js

echo.
echo --- helpers (Excel/Access mocked, Office tests skipped) ---
cscript.exe %mypath%launcher.js %mypath%test-helpers.js

echo.
echo ============ network suite ==========

echo.
echo --- http (requires network) ---
cscript.exe %mypath%launcher.js %mypath%test-http.js

echo.
echo ============ desktop suite ==========

echo.
echo --- ui (HTA windows will flash briefly) ---
cscript.exe %mypath%launcher.js %mypath%test-ui.js

pause
