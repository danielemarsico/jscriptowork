@echo off
echo ***************************************************************
echo ******* Hello, friend
echo ******* You were wondering why you have to execute this script
echo ******* unfortunately shit happens


SET mypath=%~dp0

echo %mypath:~0,-1%
echo %~dp1

cscript.exe launcher.js %~dp0%1


pause