$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) {
    throw ("vswhere.exe was not found at '{0}'. Install/repair Visual Studio 2022 using Visual Studio Installer." -f $vswhere)
}

$json = & $vswhere -products * -version '[17.0,18.0)' -requires Microsoft.Component.MSBuild -format json
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace(($json -join "`n"))) {
    throw 'No Visual Studio 2022 installation with MSBuild was found.'
}

$instances = @((($json -join "`n") | ConvertFrom-Json))
if ($instances.Count -lt 1) {
    throw 'No Visual Studio 2022 installation with MSBuild was found.'
}

$rejections = New-Object 'System.Collections.Generic.List[string]'
foreach ($instance in $instances) {
    [string]$installationPath = [string]$instance.installationPath
    if ([string]::IsNullOrWhiteSpace($installationPath)) { continue }

    [string]$msbuild = Join-Path $installationPath 'MSBuild\Current\Bin\MSBuild.exe'
    [string]$devenv = Join-Path $installationPath 'Common7\IDE\devenv.com'
    [string]$disableOutOfProc = Join-Path $installationPath 'Common7\IDE\CommonExtensions\Microsoft\VSI\DisableOutOfProcBuild\DisableOutOfProcBuild.exe'

    $missing = New-Object 'System.Collections.Generic.List[string]'
    if (-not (Test-Path -LiteralPath $msbuild -PathType Leaf)) { [void]$missing.Add('MSBuild.exe') }
    if (-not (Test-Path -LiteralPath $devenv -PathType Leaf)) { [void]$missing.Add('devenv.com') }
    if (-not (Test-Path -LiteralPath $disableOutOfProc -PathType Leaf)) { [void]$missing.Add('Visual Studio Installer Projects tooling') }

    if ($missing.Count -eq 0) {
        return [pscustomobject]@{
            InstallationPath = $installationPath
            MSBuild = $msbuild
            Devenv = $devenv
            DisableOutOfProcBuild = $disableOutOfProc
        }
    }

    [void]$rejections.Add(("{0}: missing {1}" -f $installationPath, ([string]::Join(', ', $missing.ToArray()))))
}

throw ("No Visual Studio 2022 installation has all required MSI build tools. Checked:`r`n{0}" -f ([string]::Join("`r`n", $rejections.ToArray())))
