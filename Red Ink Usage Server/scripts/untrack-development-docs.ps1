#requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$serverRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$allowed = @(
    'docs/ADMINISTRATOR_GUIDE.md',
    'docs/USER_GUIDE.md',
    'docs/THIRD_PARTY_LICENSES.md'
)
$git = Get-Command 'git.exe' -ErrorAction Stop
# -z preserves whitespace and non-ASCII paths. --cached removes index entries only.
$raw = & $git.Source -C $serverRoot ls-files -z -- docs
if ($LASTEXITCODE -ne 0) { throw 'Unable to enumerate tracked server documents.' }
$found = @($raw -split "`0" | Where-Object { $_ -ne '' })
foreach ($name in $found) {
    $normalized = $name.Replace('\', '/')
    if ($normalized -in $allowed) { continue }
    if ($PSCmdlet.ShouldProcess($normalized, 'Remove from Git index; keep the local file')) {
        & $git.Source -C $serverRoot rm --cached -- $normalized
        if ($LASTEXITCODE -ne 0) { throw "Could not remove tracked document from Git index: $normalized" }
    }
}
Write-Host 'Finished. Inspect git status before committing. Local document files were not deleted.'
