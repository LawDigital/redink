param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$packagingRoot = Split-Path -Parent $PSScriptRoot
$outputRoot = Join-Path $packagingRoot 'Output'
if (-not (Test-Path -LiteralPath $outputRoot -PathType Container)) {
    throw ("Output folder not found: {0}" -f $outputRoot)
}

$msiFiles = @(Get-ChildItem -LiteralPath $outputRoot -Recurse -File -Filter '*.msi' -ErrorAction Stop | Sort-Object FullName)
if ($msiFiles.Count -eq 0) {
    throw ("No MSI files found under: {0}" -f $outputRoot)
}

$paths = New-Object 'System.Collections.Generic.List[string]'
foreach ($msiFile in $msiFiles) { [void]$paths.Add($msiFile.FullName) }
& (Join-Path $PSScriptRoot 'Sign-MsiFiles.ps1') -Path $paths.ToArray()

# Signing changes MSI bytes. Refresh SHA256SUMS.txt in every release directory that contains an MSI.
$releaseDirectories = @($msiFiles | ForEach-Object { $_.Directory.FullName } | Sort-Object -Unique)
foreach ($releaseDirectory in $releaseDirectories) {
    $hashFile = Join-Path $releaseDirectory 'SHA256SUMS.txt'
    $hashLines = New-Object 'System.Collections.Generic.List[string]'
    $releaseFiles = @(Get-ChildItem -LiteralPath $releaseDirectory -File | Where-Object { $_.Extension -in '.msi','.exe' } | Sort-Object Name)
    foreach ($releaseFile in $releaseFiles) {
        $fileHash = Get-FileHash -LiteralPath $releaseFile.FullName -Algorithm SHA256
        [void]$hashLines.Add(("{0}  {1}" -f $fileHash.Hash, $releaseFile.Name))
    }
    [System.IO.File]::WriteAllLines($hashFile, $hashLines.ToArray(), [System.Text.Encoding]::ASCII)
    Write-Host ("Refreshed hashes: {0}" -f $hashFile)
}

Write-Host ''
Write-Host ("SIGNED AND VERIFIED {0} MSI FILE(S)." -f $msiFiles.Count)
