@echo off
echo ========================================================
echo Uninstalling AI Speaker Notes Add-in from PowerPoint...
echo ========================================================
powershell.exe -ExecutionPolicy Bypass -File "%~dp0scripts\uninstall.ps1"
echo.
pause
