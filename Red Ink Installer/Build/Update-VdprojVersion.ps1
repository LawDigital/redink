param(
    [Parameter(Mandatory=$true)][string]$ProjectPath,
    [Parameter(Mandatory=$true)][ValidatePattern('^\d+\.\d+\.\d+$')][string]$MsiVersion
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$packagingRoot = Split-Path -Parent $PSScriptRoot
[string]$packagingRootFull = [System.IO.Path]::GetFullPath($packagingRoot)
if (-not $packagingRootFull.EndsWith([System.IO.Path]::DirectorySeparatorChar.ToString())) {
    $packagingRootFull += [System.IO.Path]::DirectorySeparatorChar
}
[string]$projectPathFull = [System.IO.Path]::GetFullPath($ProjectPath)
if (-not $projectPathFull.StartsWith($packagingRootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw ("Safety check failed: Update-VdprojVersion.ps1 may modify only files under Red Ink Installer: {0}" -f $projectPathFull)
}

if (-not (Test-Path -LiteralPath $ProjectPath)) {
    throw ("Installer project not found: {0}" -f $ProjectPath)
}

$versionParts = @($MsiVersion -split '\.')
[int]$versionMajor = [int]$versionParts[0]
[int]$versionMinor = [int]$versionParts[1]
[int]$versionBuild = [int]$versionParts[2]
if ($versionMajor -gt 255 -or $versionMinor -gt 255 -or $versionBuild -gt 65535) {
    throw ("MSI ProductVersion {0} exceeds Windows Installer limits (255.255.65535)." -f $MsiVersion)
}

$content = [System.IO.File]::ReadAllText($ProjectPath)

$productVersionPattern = '(?m)^\s*"ProductVersion"\s*=\s*"8:([^"]+)"\s*$'
$productCodePattern = '(?m)^\s*"ProductCode"\s*=\s*"8:(\{[0-9A-Fa-f-]+\})"\s*$'
$packageCodePattern = '(?m)^\s*"PackageCode"\s*=\s*"8:(\{[0-9A-Fa-f-]+\})"\s*$'
$upgradeCodePattern = '(?m)^\s*"UpgradeCode"\s*=\s*"8:(\{[0-9A-Fa-f-]+\})"\s*$'

$productVersionRegex = New-Object System.Text.RegularExpressions.Regex($productVersionPattern)
$productCodeRegex = New-Object System.Text.RegularExpressions.Regex($productCodePattern)
$packageCodeRegex = New-Object System.Text.RegularExpressions.Regex($packageCodePattern)
$upgradeCodeRegex = New-Object System.Text.RegularExpressions.Regex($upgradeCodePattern)

$productVersionMatches = $productVersionRegex.Matches($content)
$productCodeMatches = $productCodeRegex.Matches($content)
$packageCodeMatches = $packageCodeRegex.Matches($content)
$upgradeCodeMatches = $upgradeCodeRegex.Matches($content)

if ($productVersionMatches.Count -ne 1) { throw ("Expected exactly one ProductVersion in {0}; found {1}." -f $ProjectPath, $productVersionMatches.Count) }
if ($productCodeMatches.Count -ne 1) { throw ("Expected exactly one ProductCode in {0}; found {1}." -f $ProjectPath, $productCodeMatches.Count) }
if ($packageCodeMatches.Count -ne 1) { throw ("Expected exactly one PackageCode in {0}; found {1}." -f $ProjectPath, $packageCodeMatches.Count) }
if ($upgradeCodeMatches.Count -ne 1) { throw ("Expected exactly one UpgradeCode in {0}; found {1}." -f $ProjectPath, $upgradeCodeMatches.Count) }

[string]$currentVersion = $productVersionMatches[0].Groups[1].Value
[string]$previousProductCode = $productCodeMatches[0].Groups[1].Value
[string]$previousPackageCode = $packageCodeMatches[0].Groups[1].Value
[string]$upgradeCode = $upgradeCodeMatches[0].Groups[1].Value
[bool]$versionChanged = $currentVersion -ne $MsiVersion

[string]$newProductCode = $previousProductCode
if ($versionChanged) {
    $newProductCode = '{' + ([System.Guid]::NewGuid().ToString().ToUpperInvariant()) + '}'
    $content = $productCodeRegex.Replace($content, ('        "ProductCode" = "8:{0}"' -f $newProductCode), 1)
}

# PackageCode identifies this physical MSI package and must be unique for each rebuild.
[string]$newPackageCode = '{' + ([System.Guid]::NewGuid().ToString().ToUpperInvariant()) + '}'
$content = $packageCodeRegex.Replace($content, ('        "PackageCode" = "8:{0}"' -f $newPackageCode), 1)
$content = $productVersionRegex.Replace($content, ('        "ProductVersion" = "8:{0}"' -f $MsiVersion), 1)

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($ProjectPath, $content, $utf8Bom)

# Verify the write, including the invariant that UpgradeCode is never changed by this helper.
$verifyContent = [System.IO.File]::ReadAllText($ProjectPath)
$verifyVersion = $productVersionRegex.Match($verifyContent)
$verifyProductCode = $productCodeRegex.Match($verifyContent)
$verifyPackageCode = $packageCodeRegex.Match($verifyContent)
$verifyUpgradeCode = $upgradeCodeRegex.Match($verifyContent)
if (-not $verifyVersion.Success -or $verifyVersion.Groups[1].Value -ne $MsiVersion) { throw ("ProductVersion verification failed after writing {0}." -f $ProjectPath) }
if (-not $verifyProductCode.Success -or $verifyProductCode.Groups[1].Value -ne $newProductCode) { throw ("ProductCode verification failed after writing {0}." -f $ProjectPath) }
if (-not $verifyPackageCode.Success -or $verifyPackageCode.Groups[1].Value -ne $newPackageCode) { throw ("PackageCode verification failed after writing {0}." -f $ProjectPath) }
if (-not $verifyUpgradeCode.Success -or $verifyUpgradeCode.Groups[1].Value -ne $upgradeCode) { throw ("UpgradeCode changed unexpectedly in {0}." -f $ProjectPath) }
if ($newPackageCode -eq $previousPackageCode) { throw ("PackageCode was not refreshed in {0}." -f $ProjectPath) }
if ($versionChanged -and $newProductCode -eq $previousProductCode) { throw ("ProductCode did not change with ProductVersion in {0}." -f $ProjectPath) }
if ((-not $versionChanged) -and $newProductCode -ne $previousProductCode) { throw ("ProductCode changed even though ProductVersion did not change in {0}." -f $ProjectPath) }

[pscustomobject]@{
    ProjectPath = $ProjectPath
    PreviousVersion = $currentVersion
    ProductVersion = $MsiVersion
    ProductCodeChanged = $versionChanged
    PackageCodeChanged = $true
}
