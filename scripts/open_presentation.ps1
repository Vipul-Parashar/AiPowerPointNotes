# scripts/open_presentation.ps1 - Opens the generated presentation in PowerPoint for review
$projectRoot = Split-Path -Parent $PSScriptRoot
$sampleOutput = Join-Path $projectRoot "sample_presentation_with_notes.pptx"

if (-not (Test-Path $sampleOutput)) {
    & "$PSScriptRoot\test_generator.ps1"
}

Write-Host "Opening presentation in PowerPoint: $sampleOutput" -ForegroundColor Cyan
Start-Process -FilePath "POWERPNT.EXE" -ArgumentList "`"$sampleOutput`""

Write-Host ""
Write-Host "Instructions to view notes in Presenter View:" -ForegroundColor Green
Write-Host "1. In PowerPoint, click on any slide and check the 'Notes' pane at the bottom."
Write-Host "2. Go to the top ribbon: 'Slide Show' > 'From Beginning' (or press F5)."
Write-Host "3. Right-click anywhere on the slide show and select 'Show Presenter View' (or press Alt + F5)."
Write-Host "4. You will see your full AI-generated speaker notes prominently on the right side!"
