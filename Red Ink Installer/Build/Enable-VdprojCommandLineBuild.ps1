$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$packagingRoot = Split-Path -Parent $PSScriptRoot
$vs = & (Join-Path $PSScriptRoot 'Find-VisualStudio2022.ps1')
$toolDir = Join-Path $vs.InstallationPath 'Common7\IDE\CommonExtensions\Microsoft\VSI\DisableOutOfProcBuild'
$tool = Join-Path $toolDir 'DisableOutOfProcBuild.exe'

if (-not (Test-Path -LiteralPath $tool)) {
    throw "Microsoft DisableOutOfProcBuild.exe was not found at '$tool'. Repair/install Microsoft Visual Studio Installer Projects 2022."
}

Write-Host 'Applying Microsoft Visual Studio Installer Projects command-line build workaround...' -ForegroundColor Cyan
Write-Host "Tool: $tool"
Write-Host 'This Microsoft tool changes a current-user Visual Studio registry setting.' -ForegroundColor Yellow

$exitCode = -1
Push-Location $toolDir
try {
    & $tool
    $exitCode = $LASTEXITCODE
} finally {
    Pop-Location
}

if ($exitCode -ne 0) {
    throw "DisableOutOfProcBuild.exe failed with exit code $exitCode."
}

Write-Host ''
Write-Host 'VDPROJ command-line build workaround applied successfully.' -ForegroundColor Green
Write-Host 'You can now run Check-Installer-Setup.cmd and Build-Preview-MSI.cmd.'
