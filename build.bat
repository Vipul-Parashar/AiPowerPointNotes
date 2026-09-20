@echo off
echo ========================================================
echo Building AI Speaker Notes PowerPoint Add-in...
echo ========================================================
powershell.exe -ExecutionPolicy Bypass -File "%~dp0scripts\build.ps1"
if %ERRORLEVEL% equ 0 (
    echo.
    echo Build completed successfully!
) else (
    echo.
    echo Build failed with error code %ERRORLEVEL%.
)
pause
