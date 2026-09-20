# scripts/uninstall.ps1 - Unregisters AiSpeakerNotes COM Add-in from HKCU
$ErrorActionPreference = "SilentlyContinue"

Write-Host "Unregistering AiSpeakerNotes COM Add-in..." -ForegroundColor Yellow

$connectProgId = "AiSpeakerNotes.Connect"
$connectClsid = "{E4B5C6D7-1A2B-3C4D-5E6F-7A8B9C0D1E2F}"

$paneProgId = "AiSpeakerNotes.TaskPaneControl"
$paneClsid = "{D8E5F2A1-9B3C-4E7F-8A12-34567890ABCD}"

# Remove PowerPoint Add-in registration
Remove-Item -Path "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$connectProgId" -Recurse -Force

# Remove COM Classes
Remove-Item -Path "HKCU:\Software\Classes\$connectProgId" -Recurse -Force
Remove-Item -Path "HKCU:\Software\Classes\CLSID\$connectClsid" -Recurse -Force

Remove-Item -Path "HKCU:\Software\Classes\$paneProgId" -Recurse -Force
Remove-Item -Path "HKCU:\Software\Classes\CLSID\$paneClsid" -Recurse -Force

Write-Host "✓ AiSpeakerNotes Add-in unregistered cleanly." -ForegroundColor Green
