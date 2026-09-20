# scripts/build.ps1 - Compiles AiSpeakerNotes.dll using 64-bit .NET Framework csc.exe
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$srcDir = Join-Path $projectRoot "src"
$binDir = Join-Path $projectRoot "bin"

if (-not (Test-Path $binDir)) {
    New-Item -ItemType Directory -Path $binDir -Force | Out-Null
}

$outputDll = Join-Path $binDir "AiSpeakerNotes.dll"

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) {
    $csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
}

$extDll = "C:\Windows\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll"
$msoDll = "C:\Windows\assembly\GAC_MSIL\Office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL"
$pptDll = "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.PowerPoint.dll"

$references = @(
    "System.dll",
    "System.Core.dll",
    "Microsoft.CSharp.dll",
    "System.Data.dll",
    "System.Drawing.dll",
    "System.Windows.Forms.dll",
    "System.Security.dll",
    "System.Web.Extensions.dll",
    "`"$extDll`"",
    "`"$msoDll`"",
    "`"$pptDll`""
)

$refArgs = $references | ForEach-Object { "/reference:$_" }

$csFiles = Get-ChildItem -Path $srcDir -Filter "*.cs" -Recurse | Select-Object -ExpandProperty FullName
$fileArgs = $csFiles | ForEach-Object { "`"$_`"" }

Write-Host "Compiling AiSpeakerNotes.dll..." -ForegroundColor Cyan

$allArgs = @(
    "/target:library",
    "/platform:anycpu",
    "/optimize+",
    "/out:`"$outputDll`""
) + $refArgs + $fileArgs

$proc = Start-Process -FilePath $csc -ArgumentList $allArgs -NoNewWindow -PassThru -Wait

if ($proc.ExitCode -eq 0) {
    Write-Host "Build succeeded! Output: $outputDll" -ForegroundColor Green
} else {
    Write-Error "Build failed with exit code $($proc.ExitCode)"
}
