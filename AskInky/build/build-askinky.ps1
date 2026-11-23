# Part of "Red Ink" (Ask Inky). Copyright (c) LawDigital AG, Switzerland.
# P18 W367: builds AskInky.exe (self-contained, single file), the MSI and the portable zip.
#   build-askinky.ps1                 unsigned TEST build  -> dist\AskInky-<v>-TEST-UNSIGNED\
#   build-askinky.ps1 -Sign           signed release build -> dist\AskInky-<v>\ (VISCHER AG, automatic)
#   build-askinky.ps1 -PackageOnly -ExePath <signed AskInky.exe> [-Sign]   MSI + zip from an existing EXE
#   build-askinky.ps1 -Unsigned       release names, NOT signed (W382: delivered unsigned; Alles-signieren signs
#                                     the EXE and rebuilds + signs the MSI from it). Paths: everything lives in AskInky\.
# Needs the .NET 8 SDK (dotnet) and the WiX Toolset 5.0.2 build tool; the script installs WiX into
# .tools\wix (NuGet) when it is missing. WiX is a build tool only; nothing of it is shipped.
param([switch]$Sign, [switch]$Unsigned, [switch]$PackageOnly, [string]$ExePath = '', [string]$Dotnet = '', [string]$Wix = '')
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$project = Join-Path $root 'AskInky.vbproj'
[xml]$proj = Get-Content $project -Raw
$version = ($proj.Project.PropertyGroup | Where-Object { $_.Version } | Select-Object -First 1).Version
if (-not $version) { throw 'Version fehlt in AskInky.vbproj.' }
# W382: the .NET 8 SDK is needed to compile (and to fetch WiX once); packaging an existing EXE with a given -Wix needs none.
if (-not $Dotnet) { $Dotnet = (Get-Command dotnet -ErrorAction SilentlyContinue).Source }
if (-not $Dotnet -and -not ($PackageOnly -and $Wix)) { throw 'dotnet (.NET 8 SDK) wurde nicht gefunden.' }
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'; $env:DOTNET_NOLOGO = '1'
if ($Dotnet -and -not $env:DOTNET_ROOT) { $env:DOTNET_ROOT = Split-Path -Parent $Dotnet }

$label = if ($Sign -or $Unsigned) { "AskInky-$version" } else { "AskInky-$version-TEST-UNSIGNED" }
$dist = Join-Path $root ("dist\" + $label)
$stage = Join-Path $root 'obj\askinky-setup-stage'
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force $stage, $dist | Out-Null

if (-not $PackageOnly) {
    & $Dotnet publish $project -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o (Join-Path $root 'obj\publish')
    if ($LASTEXITCODE -ne 0) { throw 'dotnet publish fehlgeschlagen.' }
    $ExePath = Join-Path $root 'obj\publish\AskInky.exe'
}
if (-not (Test-Path $ExePath)) { throw "AskInky.exe nicht gefunden: $ExePath" }
Copy-Item $ExePath (Join-Path $stage 'AskInky.exe')
Copy-Item (Join-Path $root 'LICENSE.txt') (Join-Path $stage 'LICENSE.txt')
Copy-Item (Join-Path $root 'Assets\RedInk.ico') (Join-Path $stage 'RedInk.ico')
# Licence page of the installer: RTF made from LICENSE.txt (plain text, escaped, Unicode as \uN?).
$lic = [IO.File]::ReadAllText((Join-Path $root 'LICENSE.txt'))
$sb = New-Object Text.StringBuilder
[void]$sb.Append('{\rtf1\ansi\deff0{\fonttbl{\f0 Segoe UI;}}\fs16 ')
foreach ($ch in $lic.ToCharArray()) { $c = [int]$ch; if ($ch -eq '\' -or $ch -eq '{' -or $ch -eq '}') { [void]$sb.Append('\' + $ch) } elseif ($c -eq 10) { [void]$sb.Append('\par ') } elseif ($c -eq 13) { } elseif ($c -gt 127) { [void]$sb.Append('\u' + $(if ($c -gt 32767) { $c - 65536 } else { $c }) + '?') } else { [void]$sb.Append($ch) } }
[void]$sb.Append('}')
[IO.File]::WriteAllText((Join-Path $stage 'License.rtf'), $sb.ToString(), (New-Object Text.ASCIIEncoding))
if ($Sign) { & (Join-Path $PSScriptRoot 'sign-askinky.ps1') -Path (Join-Path $stage 'AskInky.exe'); if ($LASTEXITCODE -ne 0) { throw 'Signieren der EXE fehlgeschlagen.' } }

if (-not $Wix) {
    $Wix = Join-Path $root '.tools\wix\wix.exe'
    if (-not (Test-Path $Wix)) {
        & $Dotnet tool install wix --version 5.0.2 --tool-path (Join-Path $root '.tools\wix')
        if ($LASTEXITCODE -ne 0) { throw 'WiX 5.0.2 konnte nicht installiert werden.' }
    }
}
$msi = Join-Path $dist ("AskInky-$version-x64.msi")
Push-Location $root
try {
  # WiX UI + Util extensions 5.0.2 (build tools only; the util custom action that starts Ask Inky after the install is the only WiX code inside the MSI).
  & $Wix extension add WixToolset.UI.wixext/5.0.2 WixToolset.Util.wixext/5.0.2 | Out-Null
  & $Wix build (Join-Path $root 'Setup\Package.wxs') -ext WixToolset.UI.wixext -ext WixToolset.Util.wixext -culture de-DE -arch x64 -d "ProductVersion=$version" -d "SourceDir=$stage" -o $msi
if ($LASTEXITCODE -ne 0) { throw 'MSI-Build fehlgeschlagen.' }
} finally { Pop-Location }
if ($Sign) { & (Join-Path $PSScriptRoot 'sign-askinky.ps1') -Path $msi; if ($LASTEXITCODE -ne 0) { throw 'Signieren des MSI fehlgeschlagen.' } }

$zip = Join-Path $dist ("AskInky-$version-portable.zip")
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage 'AskInky.exe'), (Join-Path $stage 'LICENSE.txt') -DestinationPath $zip
Remove-Item (Join-Path $dist '*.wixpdb') -ErrorAction SilentlyContinue

$lines = foreach ($f in Get-ChildItem $dist -File | Where-Object { $_.Name -ne 'SHA256SUMS.txt' -and $_.Name -ne 'build-info.json' }) {
    '{0}  {1}' -f (Get-FileHash $f.FullName -Algorithm SHA256).Hash.ToLower(), $f.Name
}
[IO.File]::WriteAllLines((Join-Path $dist 'SHA256SUMS.txt'), $lines)
$info = [ordered]@{ product = 'askinky-windows'; version = $version; signed = [bool]$Sign; builtAt = (Get-Date).ToString('s');
    files = @(Get-ChildItem $dist -File | Where-Object { $_.Name -like 'AskInky-*' } | ForEach-Object { [ordered]@{ name = $_.Name; size = $_.Length; sha256 = (Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLower() } }) }
[IO.File]::WriteAllText((Join-Path $dist 'build-info.json'), ($info | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding($false)))
Write-Host "Fertig: $dist"
