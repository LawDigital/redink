$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$packagingRoot = Split-Path -Parent $PSScriptRoot
& (Join-Path $PSScriptRoot 'Test-PowerShellSyntax.ps1') | Out-Null
$solutionRoot = Split-Path -Parent $packagingRoot
$solution = Join-Path $solutionRoot 'Red Ink.sln'

Write-Host 'Red Ink MSI build preflight' -ForegroundColor Cyan
Write-Host "Installer folder: $packagingRoot"
Write-Host "Solution root:    $solutionRoot"

if (-not (Test-Path -LiteralPath $solution)) { throw "Missing solution: $solution" }
Write-Host 'Red Ink.sln: OK'

$vs = & (Join-Path $PSScriptRoot 'Find-VisualStudio2022.ps1')
Write-Host "Visual Studio 2022: $($vs.InstallationPath)"
Write-Host "MSBuild:            $($vs.MSBuild)"
Write-Host "devenv.com:         $($vs.Devenv)"

$disableOutOfProcTool = [string]$vs.DisableOutOfProcBuild
if ([string]::IsNullOrWhiteSpace($disableOutOfProcTool) -or -not (Test-Path -LiteralPath $disableOutOfProcTool -PathType Leaf)) {
    throw ("Microsoft DisableOutOfProcBuild.exe not found for Visual Studio installation '{0}'. Repair/install Microsoft Visual Studio Installer Projects 2022." -f $vs.InstallationPath)
}
Write-Host 'VDPROJ command-line workaround tool: available'

foreach ($channel in @('Preview','GA')) {
    $v = & (Join-Path $PSScriptRoot 'Get-ChannelVersion.ps1') -Channel $channel -SolutionRoot $solutionRoot
    Write-Host ("{0}: ApplicationVersion={1}; MSI={2}" -f $channel,$v.ApplicationVersion,$v.MsiVersion)
}

$projectsRoot = Join-Path $packagingRoot 'InstallerProjects\Projects'
$expected = foreach ($channel in @('Preview','GA')) {
    foreach ($app in @('Word','Excel','Outlook')) {
        foreach ($arch in @('x86','x64')) {
            Join-Path $projectsRoot "$channel-$app-$arch\RedInk-$channel-$app-$arch.vdproj"
        }
    }
}
$missing = @($expected | Where-Object { -not (Test-Path -LiteralPath $_) })
if ($missing.Count -gt 0) {
    Write-Host ''
    Write-Host "Installer projects not complete yet ($($missing.Count) missing):" -ForegroundColor Yellow
    $missing | ForEach-Object { Write-Host "  $_" }
    Write-Host ''
    Write-Host 'This is expected until you complete InstallerProjects\ONE-TIME-SETUP.md.'
} else {
    foreach ($channel in @('Preview','GA')) {
        foreach ($app in @('Word','Excel','Outlook')) {
            foreach ($arch in @('x86','x64')) {
                $vdproj = Join-Path $projectsRoot "$channel-$app-$arch\RedInk-$channel-$app-$arch.vdproj"
                & (Join-Path $PSScriptRoot 'Validate-Vdproj.ps1') -Channel $channel -App $app -Arch $arch -ProjectPath $vdproj | Out-Null
            }
        }
    }
    foreach ($identityName in @('ProductCode','PackageCode','UpgradeCode')) {
        $codes = @{}
        foreach ($vdproj in $expected) {
            $text = [System.IO.File]::ReadAllText($vdproj)
            $pattern = '\"' + $identityName + '\"\s*=\s*\"8:(\{[0-9A-Fa-f-]+\})\"'
            $matches = [regex]::Matches($text, $pattern)
            if ($matches.Count -ne 1) { throw ("{0} occurrence count is {1} in {2}; expected exactly one." -f $identityName, $matches.Count, $vdproj) }
            [string]$code = $matches[0].Groups[1].Value.ToUpperInvariant()
            if ($codes.ContainsKey($code)) { throw ("Duplicate {0} in {1} and {2}" -f $identityName, $vdproj, $codes[$code]) }
            $codes[$code] = $vdproj
        }
        Write-Host ("All 12 {0}s: unique and present" -f $identityName)
    }
    Write-Host 'All 12 installer projects: OK and validated'
}
