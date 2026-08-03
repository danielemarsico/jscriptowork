@echo off
echo =====================================
echo  jscriptowork test runner
echo =====================================

SET mypath=%~dp0
SET launcher=%mypath%..\launcher.js

echo.
echo ============ offline suites =========

echo.
echo --- core ---
cscript.exe %launcher% %mypath%test-core.js

echo.
echo --- ext ---
cscript.exe %launcher% %mypath%test-ext.js

echo.
echo --- polyfills ---
cscript.exe %launcher% %mypath%test-polyfills.js

echo.
echo --- minitest ---
cscript.exe %launcher% %mypath%test-minitest.js

echo.
echo --- crypto ---
cscript.exe %launcher% %mypath%test-crypto.js

echo.
echo --- minimist ---
cscript.exe %launcher% %mypath%test-minimist.js

echo.
echo ============ disk suites ============

echo.
echo --- system ---
cscript.exe %launcher% %mypath%test-system.js

echo.
echo --- filesystem ---
cscript.exe %launcher% %mypath%test-filesystem.js

echo.
echo --- helpers (Excel/Access mocked, Office tests skipped) ---
cscript.exe %launcher% %mypath%test-helpers.js

echo.
echo ============ network suite ==========

echo.
echo --- http (requires network) ---
cscript.exe %launcher% %mypath%test-http.js

echo.
echo ============ desktop suite ==========

echo.
echo --- ui (HTA windows will flash briefly) ---
cscript.exe %launcher% %mypath%test-ui.js

pause
