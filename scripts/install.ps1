# scripts/install.ps1 - Registers AiSpeakerNotes COM Add-in for PowerPoint (Current User, No Admin Needed)
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$dllPath = Join-Path $projectRoot "bin\AiSpeakerNotes.dll"

if (-not (Test-Path $dllPath)) {
    Write-Host "DLL not found. Running build..." -ForegroundColor Cyan
    & "$PSScriptRoot\build.ps1"
}

$dllUri = "file:///" + ($dllPath -replace '\\', '/')
$assemblyName = "AiSpeakerNotes, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
$runtimeVersion = "v4.0.30319"

Write-Host "Registering AiSpeakerNotes COM Add-in in HKCU..." -ForegroundColor Cyan

# 1. Register Connect class (CLSID: {E4B5C6D7-1A2B-3C4D-5E6F-7A8B9C0D1E2F})
$connectProgId = "AiSpeakerNotes.Connect"
$connectClsid = "{E4B5C6D7-1A2B-3C4D-5E6F-7A8B9C0D1E2F}"

# HKCU\Software\Classes\AiSpeakerNotes.Connect
New-Item -Path "HKCU:\Software\Classes\$connectProgId" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Classes\$connectProgId" -Name "(Default)" -Value $connectProgId
New-Item -Path "HKCU:\Software\Classes\$connectProgId\CLSID" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Classes\$connectProgId\CLSID" -Name "(Default)" -Value $connectClsid

# HKCU\Software\Classes\CLSID\{...}
$clsidPath = "HKCU:\Software\Classes\CLSID\$connectClsid"
New-Item -Path $clsidPath -Force | Out-Null
Set-ItemProperty -Path $clsidPath -Name "(Default)" -Value $connectProgId

$inproc = "$clsidPath\InprocServer32"
New-Item -Path $inproc -Force | Out-Null
Set-ItemProperty -Path $inproc -Name "(Default)" -Value "mscoree.dll"
Set-ItemProperty -Path $inproc -Name "ThreadingModel" -Value "Both"
Set-ItemProperty -Path $inproc -Name "Class" -Value $connectProgId
Set-ItemProperty -Path $inproc -Name "Assembly" -Value $assemblyName
Set-ItemProperty -Path $inproc -Name "RuntimeVersion" -Value $runtimeVersion
Set-ItemProperty -Path $inproc -Name "CodeBase" -Value $dllUri

New-Item -Path "$clsidPath\ProgId" -Force | Out-Null
Set-ItemProperty -Path "$clsidPath\ProgId" -Name "(Default)" -Value $connectProgId

New-Item -Path "$clsidPath\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" -Force | Out-Null


# 2. Register TaskPaneControl class (CLSID: {D8E5F2A1-9B3C-4E7F-8A12-34567890ABCD})
$paneProgId = "AiSpeakerNotes.TaskPaneControl"
$paneClsid = "{D8E5F2A1-9B3C-4E7F-8A12-34567890ABCD}"

New-Item -Path "HKCU:\Software\Classes\$paneProgId" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Classes\$paneProgId" -Name "(Default)" -Value $paneProgId
New-Item -Path "HKCU:\Software\Classes\$paneProgId\CLSID" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Classes\$paneProgId\CLSID" -Name "(Default)" -Value $paneClsid

$paneClsidPath = "HKCU:\Software\Classes\CLSID\$paneClsid"
New-Item -Path $paneClsidPath -Force | Out-Null
Set-ItemProperty -Path $paneClsidPath -Name "(Default)" -Value $paneProgId

$paneInproc = "$paneClsidPath\InprocServer32"
New-Item -Path $paneInproc -Force | Out-Null
Set-ItemProperty -Path $paneInproc -Name "(Default)" -Value "mscoree.dll"
Set-ItemProperty -Path $paneInproc -Name "ThreadingModel" -Value "Both"
Set-ItemProperty -Path $paneInproc -Name "Class" -Value $paneProgId
Set-ItemProperty -Path $paneInproc -Name "Assembly" -Value $assemblyName
Set-ItemProperty -Path $paneInproc -Name "RuntimeVersion" -Value $runtimeVersion
Set-ItemProperty -Path $paneInproc -Name "CodeBase" -Value $dllUri

New-Item -Path "$paneClsidPath\ProgId" -Force | Out-Null
Set-ItemProperty -Path "$paneClsidPath\ProgId" -Name "(Default)" -Value $paneProgId
New-Item -Path "$paneClsidPath\Control" -Force | Out-Null
New-Item -Path "$paneClsidPath\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" -Force | Out-Null


# 3. Register Addin in PowerPoint
$pptAddinKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$connectProgId"
New-Item -Path $pptAddinKey -Force | Out-Null
Set-ItemProperty -Path $pptAddinKey -Name "FriendlyName" -Value "AI Speaker Notes Generator"
Set-ItemProperty -Path $pptAddinKey -Name "Description" -Value "Automatically writes plain-English spoken notes for Presenter View"
Set-ItemProperty -Path $pptAddinKey -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $pptAddinKey -Name "CommandLineSafe" -Value 0 -Type DWord

Write-Host ""
Write-Host "AiSpeakerNotes Add-in registered successfully in PowerPoint!" -ForegroundColor Green
Write-Host "Next steps:"
Write-Host "1. Launch PowerPoint."
Write-Host "2. You will see the 'AI Speaker Notes' tab in the top ribbon."
Write-Host "3. Click 'Notes Task Pane' or 'Generate Notes'."

