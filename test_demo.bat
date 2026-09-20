@echo off
echo ========================================================
echo Running End-to-End Demo on Sample 5-Slide Presentation...
echo ========================================================
powershell.exe -ExecutionPolicy Bypass -File "%~dp0scripts\test_generator.ps1"
if %ERRORLEVEL% equ 0 (
    echo.
    echo Launching presentation in PowerPoint to inspect notes...
    powershell.exe -ExecutionPolicy Bypass -File "%~dp0scripts\open_presentation.ps1"
) else (
    echo.
    echo Demo encountered an error.
)
pause
