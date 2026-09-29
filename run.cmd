@echo off
setlocal
cd /d "%~dp0"
call mvnw.cmd javafx:run
if errorlevel 1 pause
