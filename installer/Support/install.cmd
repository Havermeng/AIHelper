@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-AIHelper.ps1" %*
exit /b %ERRORLEVEL%
