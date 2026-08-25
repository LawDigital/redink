param(
    [Parameter(Mandatory=$true)][ValidateSet('Preview','GA')][string]$Channel,
    [Parameter(Mandatory=$true)][string]$SolutionRoot
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$projects = @(
    (Join-Path $SolutionRoot 'Red Ink for Word\Red Ink for Word.vbproj')
    (Join-Path $SolutionRoot 'Red Ink for Excel\Red Ink for Excel.vbproj')
    (Join-Path $SolutionRoot 'Red Ink for Outlook\Red Ink for Outlook.vbproj')
)

$versions = New-Object 'System.Collections.Generic.List[string]'

foreach ($project in $projects) {
    if (-not (Test-Path -LiteralPath $project)) {
        throw "Project not found: $project"
    }

    [xml]$xml = Get-Content -LiteralPath $project -Raw

    # MSBuild .vbproj files use a default XML namespace. local-name() keeps this
    # lookup namespace-independent and also works if the project format changes.
    $groups = @($xml.SelectNodes("/*[local-name()='Project']/*[local-name()='PropertyGroup'][@Condition]"))

    $applicationVersion = $null
    foreach ($group in $groups) {
        $conditionAttribute = $group.Attributes['Condition']
        if ($null -eq $conditionAttribute) { continue }

        [string]$condition = $conditionAttribute.Value
        if ([string]::IsNullOrWhiteSpace($condition)) { continue }

        $channelPattern = "RedInkEnvironment.*'" + [System.Text.RegularExpressions.Regex]::Escape($Channel) + "'"
        if ($condition -notmatch $channelPattern) { continue }

        $versionNode = $group.SelectSingleNode("*[local-name()='ApplicationVersion']")
        if ($null -eq $versionNode) { continue }

        [string]$candidate = $versionNode.InnerText
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }

        $applicationVersion = $candidate.Trim()
        break
    }

    if ([string]::IsNullOrWhiteSpace([string]$applicationVersion)) {
        throw "No $Channel ApplicationVersion found in $project"
    }

    $versions.Add([string]$applicationVersion)
}

$unique = @($versions | Sort-Object -Unique)
if ($unique.Count -ne 1) {
    throw "$Channel versions do not match. Word=$($versions[0]), Excel=$($versions[1]), Outlook=$($versions[2])"
}

[string]$applicationVersion = [string]$unique[0]
$parts = @($applicationVersion -split '\.')
if ($parts.Count -lt 3) {
    throw "Invalid ApplicationVersion: $applicationVersion"
}

if (@($parts[0..2] | Where-Object { $_ -notmatch '^\d+$' }).Count -gt 0) {
    throw "Invalid numeric ApplicationVersion: $applicationVersion"
}

[int]$major = [int]$parts[0]
[int]$minor = [int]$parts[1]
[int]$build = [int]$parts[2]
if ($major -gt 255 -or $minor -gt 255 -or $build -gt 65535) {
    throw ("ApplicationVersion {0} cannot be represented as a Windows Installer ProductVersion. Limits are major<=255, minor<=255, build<=65535." -f $applicationVersion)
}

$msiVersion = '{0}.{1}.{2}' -f $major, $minor, $build

[pscustomobject]@{
    ApplicationVersion = $applicationVersion
    MsiVersion = $msiVersion
}
