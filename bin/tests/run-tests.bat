@echo off
setlocal enabledelayedexpansion
echo =====================================
echo  jscriptowork test runner
echo =====================================

SET mypath=%~dp0
SET launcher=%mypath%..\launcher.js
SET OVERALL_EXIT=0

echo.
echo ============ offline suites =========

echo.
echo --- core ---
cscript.exe %launcher% %mypath%test-core.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- ext ---
cscript.exe %launcher% %mypath%test-ext.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- polyfills ---
cscript.exe %launcher% %mypath%test-polyfills.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- minitest ---
cscript.exe %launcher% %mypath%test-minitest.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- console ---
cscript.exe %launcher% %mypath%test-console.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- csv ---
cscript.exe %launcher% %mypath%test-csv.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- crypto ---
cscript.exe %launcher% %mypath%test-crypto.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- base64 ---
cscript.exe %launcher% %mypath%test-base64.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- minimist ---
cscript.exe %launcher% %mypath%test-minimist.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo ============ disk suites ============

echo.
echo --- system ---
cscript.exe %launcher% %mypath%test-system.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- filesystem ---
cscript.exe %launcher% %mypath%test-filesystem.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- log (levels captured, file tee on disk) ---
cscript.exe %launcher% %mypath%test-log.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- helpers (Excel/Access mocked, Office tests skipped) ---
cscript.exe %launcher% %mypath%test-helpers.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- win (HKCU registry sandbox, cmd.exe subprocesses) ---
cscript.exe %launcher% %mypath%test-win.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo --- build (spawns a real `cscript build.js`, regenerates dist/) ---
cscript.exe %launcher% %mypath%test-build.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo ============ network suite ==========

echo.
if "%CI%"=="true" (
    echo --- http - OFFLINE stub, CI=true ---
) else (
    echo --- http - requires network ---
)
cscript.exe %launcher% %mypath%test-http.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo ============ desktop suite ==========

echo.
if "%CI%"=="true" (
    echo --- ui - HEADLESS, CI=true: window tests skipped, parsing tests run ---
) else (
    echo --- ui - HTA windows will flash briefly ---
)
cscript.exe %launcher% %mypath%test-ui.js
if errorlevel 1 set OVERALL_EXIT=1

echo.
echo =====================================
if "%OVERALL_EXIT%"=="0" (
    echo  ALL SUITES PASSED
) else (
    echo  ONE OR MORE SUITES FAILED
)
echo =====================================

if not "%CI%"=="true" pause

exit /b %OVERALL_EXIT%
