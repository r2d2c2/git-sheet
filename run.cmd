@echo off
setlocal
cd /d "%~dp0"
call gradlew.bat run
if errorlevel 1 pause
