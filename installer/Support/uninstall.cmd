@echo off
setlocal
set "INSTALL_DIR=%~dp0"
if "%INSTALL_DIR:~-1%"=="\" set "INSTALL_DIR=%INSTALL_DIR:~0,-1%"
set "TEMP_SCRIPT=%TEMP%\AIHelper-Uninstall-%RANDOM%%RANDOM%.ps1"
copy "%~dp0Uninstall-AIHelper.ps1" "%TEMP_SCRIPT%" >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%TEMP_SCRIPT%" -InstallDir "%INSTALL_DIR%" %*
set "EXITCODE=%ERRORLEVEL%"
del "%TEMP_SCRIPT%" >nul 2>&1
if exist "%INSTALL_DIR%" start "" /min cmd.exe /c "ping 127.0.0.1 -n 3 >nul & rmdir /s /q ""%INSTALL_DIR%"""
exit /b %EXITCODE%
