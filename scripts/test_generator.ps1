# scripts/test_generator.ps1 - Automated End-to-End Test for AI Speaker Notes
param(
    [string]$Provider = "Mock",
    [string]$ApiKey = "",
    [string]$Model = ""
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$dllPath = Join-Path $projectRoot "bin\AiSpeakerNotes.dll"
$sampleInput = Join-Path $projectRoot "sample_presentation.pptx"
$sampleOutput = Join-Path $projectRoot "sample_presentation_with_notes.pptx"

if (-not (Test-Path $dllPath)) {
    & "$PSScriptRoot\build.ps1"
}

if (-not (Test-Path $sampleInput)) {
    & "$PSScriptRoot\create_sample_presentation.ps1"
}

Write-Host "Loading assembly: $dllPath" -ForegroundColor Cyan
Add-Type -Path $dllPath

# Configure settings
$settings = New-Object AiSpeakerNotes.Models.GenerationSettings
switch ($Provider.ToLower()) {
    "gemini" { $settings.Provider = [AiSpeakerNotes.Models.LlmProvider]::Gemini }
    "openai" { $settings.Provider = [AiSpeakerNotes.Models.LlmProvider]::OpenAI }
    "claude" { $settings.Provider = [AiSpeakerNotes.Models.LlmProvider]::Claude }
    default  { $settings.Provider = [AiSpeakerNotes.Models.LlmProvider]::Mock }
}

if ($ApiKey) { $settings.ApiKey = $ApiKey }
if ($Model) { $settings.Model = $Model }
$settings.LanguageLevel = "Normal"
$settings.Length = "Medium (5-6 sentences)"
$settings.Audience = "college students, computer science"
$settings.OverwriteExisting = $true

Write-Host "Settings configured:" -ForegroundColor Cyan
Write-Host "  Provider: $($settings.Provider)"
Write-Host "  Model: $($settings.Model)"
Write-Host "  Audience: $($settings.Audience)"
Write-Host "  Length: $($settings.Length)"

# Copy sample input to sample output
Copy-Item $sampleInput $sampleOutput -Force

$ppt = New-Object -ComObject PowerPoint.Application
$llmService = New-Object AiSpeakerNotes.Services.LlmService

try {
    $pres = $ppt.Presentations.Open($sampleOutput, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    $totalSlides = $pres.Slides.Count
    Write-Host ""
    Write-Host "Opened presentation with $totalSlides slides." -ForegroundColor Cyan

    # Pre-extract all slides
    $slidesData = @()
    for ($i = 1; $i -le $totalSlides; $i++) {
        $s = $pres.Slides.Item($i)
        $data = [AiSpeakerNotes.Services.SlideContentExtractor]::ExtractSlide($s)
        $slidesData += $data
    }

    # Generate notes for each slide
    for ($i = 1; $i -le $totalSlides; $i++) {
        $slideObj = $pres.Slides.Item($i)
        $slideData = $slidesData[$i - 1]

        # Populate context summaries
        if ($i -gt 1) {
            $slideData.PrevSlideSummary = [AiSpeakerNotes.Services.SlideContentExtractor]::SummarizeSlide($slidesData[$i - 2])
        }
        if ($i -lt $totalSlides) {
            $slideData.NextSlideSummary = [AiSpeakerNotes.Services.SlideContentExtractor]::SummarizeSlide($slidesData[$i])
        }

        Write-Host ""
        Write-Host "------------------------------------------------------------" -ForegroundColor Yellow
        Write-Host "Slide $i Content Analysis:" -ForegroundColor Yellow
        Write-Host "  Title: '$($slideData.Title)'"
        Write-Host "  Bullets Count: $($slideData.BulletPoints.Count)"
        Write-Host "  Tables Count: $($slideData.Tables.Count)"
        Write-Host "  Images/AltText Count: $($slideData.ImageAltTexts.Count)"
        if ($slideData.PrevSlideSummary) {
            Write-Host "  Context from Prev Slide: $($slideData.PrevSlideSummary)" -ForegroundColor DarkGray
        }

        Write-Host "  Generating speaker notes via $($settings.Provider)..." -ForegroundColor Cyan
        $notes = $llmService.GenerateNotesWithRetry($slideData, $settings, 3)

        # Write to NotesPage
        [AiSpeakerNotes.Services.NotesWriter]::SetNotes($slideObj, $notes)
        Write-Host "  [OK] Written to NotesPage!" -ForegroundColor Green

        # Verify read-back
        $readBack = [AiSpeakerNotes.Services.NotesWriter]::GetNotes($slideObj)
        if ($readBack -eq $notes) {
            Write-Host "  [OK] Verified read-back matches written notes exactly!" -ForegroundColor Green
        } else {
            Write-Warning "  Read-back mismatch!"
        }

        Write-Host ""
        Write-Host "[SPEAKER NOTES FOR PRESENTER VIEW - SLIDE $i]:" -ForegroundColor White
        Write-Host $notes -ForegroundColor Gray
    }

    $pres.Save()
    $pres.Close()
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "All slides processed successfully!" -ForegroundColor Green
    Write-Host "Saved presentation: $sampleOutput" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
}
finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
