@echo off
SET mypath=%~dp0
cscript.exe "%mypath%..\bin\launcher.js" "%mypath%%1"
pause
