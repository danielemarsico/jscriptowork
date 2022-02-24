@echo off
echo ***************************************************************
echo ******* Hello, friend
echo ******* You were wondering why you have to execute this script
echo ******* unfortunately shit happens


SET mypath=%~dp0
REM echo %mypath:~0,-1%


cscript.exe %mypath%%~n0".js"


pause