@echo off
echo ========================================================
echo Installing / Uploading AI Speaker Notes to PowerPoint...
echo ========================================================
powershell.exe -ExecutionPolicy Bypass -File "%~dp0scripts\install.ps1"
echo.
pause
