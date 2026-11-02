@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem bootstrap_runtime.bat
rem Prepares the embedded Python runtime for FeeMan (Tkinter is pre-bundled
rem in the runtime folder, so nothing GUI-related is downloaded here).
rem Arguments are forwarded to bootstrap_runtime.ps1, e.g.
rem   bootstrap_runtime.bat -NoLaunch   prepare only, do not start the app
rem Keep this file ASCII-only: cmd.exe reads .bat in the OEM codepage and
rem non-ASCII comments can be parsed as commands.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0bootstrap_runtime.ps1" %*
set "RC=%ERRORLEVEL%"

rem launcher.exe passes -NoLaunch and waits for us; pausing there would hang startup
if /i not "%~1"=="-NoLaunch" pause
exit /b %RC%
