param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$InstallerRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))

if (-not (Test-Path -LiteralPath $InstallerRoot -PathType Container)) {
    throw ("Installer root not found: {0}" -f $InstallerRoot)
}

$scriptFiles = @(Get-ChildItem -LiteralPath $InstallerRoot -Recurse -File -Filter '*.ps1' -ErrorAction Stop | Sort-Object FullName)
if ($scriptFiles.Count -eq 0) {
    throw ("No PowerShell scripts found under installer root: {0}" -f $InstallerRoot)
}

$allParseErrors = New-Object 'System.Collections.Generic.List[string]'
foreach ($scriptFile in $scriptFiles) {
    $tokens = $null
    $parseErrors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($scriptFile.FullName, [ref]$tokens, [ref]$parseErrors)
    if ($null -eq $parseErrors) { continue }
    foreach ($parseError in @($parseErrors)) {
        [int]$line = $parseError.Extent.StartLineNumber
        [int]$column = $parseError.Extent.StartColumnNumber
        [void]$allParseErrors.Add(("{0} ({1},{2}): {3}" -f $scriptFile.FullName, $line, $column, $parseError.Message))
    }
}

if ($allParseErrors.Count -gt 0) {
    $message = [string]::Join("`r`n", $allParseErrors.ToArray())
    throw ("PowerShell syntax audit failed:`r`n{0}" -f $message)
}

# Regression guards for PowerShell 5.1 / earlier failures in this installer toolchain.
$forbiddenChecks = @(
    @{ Name = 'Path.GetRelativePath (not available on Windows PowerShell 5.1/.NET Framework)'; Pattern = ('[System.IO.Path]::Get' + 'RelativePath') },
    @{ Name = 'obsolete local VDPROJ workaround marker'; Pattern = ('.vdproj-commandline-' + 'workaround-applied') }
)

$forbiddenHits = New-Object 'System.Collections.Generic.List[string]'
foreach ($scriptFile in $scriptFiles) {
    [string]$scriptText = [System.IO.File]::ReadAllText($scriptFile.FullName)
    foreach ($forbiddenCheck in $forbiddenChecks) {
        if ($scriptText.IndexOf([string]$forbiddenCheck.Pattern, [System.StringComparison]::Ordinal) -ge 0) {
            [void]$forbiddenHits.Add(("{0}: {1}" -f $scriptFile.FullName, $forbiddenCheck.Name))
        }
    }
}

if ($forbiddenHits.Count -gt 0) {
    $message = [string]::Join("`r`n", $forbiddenHits.ToArray())
    throw ("PowerShell regression audit failed:`r`n{0}" -f $message)
}

Write-Output ([pscustomobject]@{
    ScriptsParsed = $scriptFiles.Count
    Valid = $true
})
