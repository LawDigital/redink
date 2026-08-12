param(
    [Parameter(Mandatory=$true)][ValidateSet('Preview','GA')][string]$Channel,
    [Parameter(Mandatory=$true)][ValidateSet('Word','Excel','Outlook')][string]$App,
    [Parameter(Mandatory=$true)][ValidateSet('x86','x64')][string]$Arch,
    [Parameter(Mandatory=$true)][string]$ProjectPath
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

if (-not (Test-Path -LiteralPath $ProjectPath)) { throw "Installer project not found: $ProjectPath" }
$content = [System.IO.File]::ReadAllText($ProjectPath)

$appGuid = @{
    Word    = '{78C8EDF7-8504-4E46-AACA-C9099800605A}'
    Excel   = '{3B6621E2-218B-4807-A9E6-F241124D8287}'
    Outlook = '{113F0BEC-AA7D-497D-906A-2BC3E8FEF889}'
}[$App]

$productName = if ($Channel -eq 'Preview') { "Red Ink for $App (Preview)" } else { "Red Ink for $App" }
$targetPlatform = if ($Arch -eq 'x64') { '1' } else { '0' }
$appFolder = if ($Channel -eq 'Preview') { "[ProgramFilesFolder]\\Red Ink\\Preview\\$App\\" } else { "[ProgramFilesFolder]\\Red Ink\\$App\\" }
$opposite = if ($Channel -eq 'Preview') { 'GA' } else { 'Preview' }
$oppositeProperty = ("{0}{1}INSTALLED" -f $opposite.ToUpperInvariant(), $App.ToUpperInvariant())
$oppositePath = "Software\\Red Ink\\MSI\\$opposite\\$App"
$sourcePart = if ($Arch -eq 'x64') { "Red Ink for $App\\bin\\x64\\Release" } else { "Red Ink for $App\\bin\\Release" }
$releaseOutput = "Release\\$Channel-$App-$Arch.msi"

# Extract only the Registry section. .vdproj stores registry paths as a nested tree,
# not as one literal Software\\Microsoft\\Office... string.
$registryMatch = [regex]::Match($content, '(?ms)^\s*"Registry"\s*\{(?<body>.*?)^\s*"Sequences"\s*\{')
if (-not $registryMatch.Success) { throw "Installer validation failed [Registry section missing]: $ProjectPath" }
$registry = [string]$registryMatch.Groups['body'].Value

function Test-OrderedNames {
    param([string]$Text,[string[]]$Names)
    [int]$pos = 0
    foreach ($name in $Names) {
        $needle = '"Name" = "8:' + $name + '"'
        $next = $Text.IndexOf($needle, $pos, [System.StringComparison]::Ordinal)
        if ($next -lt 0) { return $false }
        $pos = $next + $needle.Length
    }
    return $true
}

$checks = @(
    @{ Name='Manufacturer'; Ok=$content.Contains('"Manufacturer" = "8:LawDigital Ltd."') },
    @{ Name='InstallAllUsers'; Ok=$content.Contains('"InstallAllUsers" = "11:TRUE"') },
    @{ Name='RemovePreviousVersions'; Ok=$content.Contains('"RemovePreviousVersions" = "11:TRUE"') },
    @{ Name='DetectNewerInstalledVersion'; Ok=$content.Contains('"DetectNewerInstalledVersion" = "11:TRUE"') },
    @{ Name='ProductName'; Ok=$content.Contains('"ProductName" = "8:' + $productName + '"') },
    @{ Name='TargetPlatform'; Ok=$content.Contains('"TargetPlatform" = "3:' + $targetPlatform + '"') },
    @{ Name='Application Folder'; Ok=$content.Contains('"DefaultLocation" = "8:' + $appFolder + '"') },
    @{ Name='Office registry tree'; Ok=(Test-OrderedNames -Text $registry -Names @('Microsoft','Office',$App,'Addins',"Red Ink for $App")) },
    @{ Name='Office Manifest value'; Ok=$registry.Contains('"Value" = "8:file:///[TARGETDIR]Red Ink for ' + $App + '.vsto|vstolocal"') },
    @{ Name='Office FriendlyName'; Ok=$registry.Contains('"Name" = "8:FriendlyName"') -and $registry.Contains('"Value" = "8:' + $productName + '"') },
    @{ Name='Office Description'; Ok=$registry.Contains('"Name" = "8:Description"') -and $registry.Contains('"Value" = "8:' + $productName + '"') },
    @{ Name='Office LoadBehavior'; Ok=$registry.Contains('"Name" = "8:LoadBehavior"') -and $registry.Contains('"Value" = "3:3"') },
    @{ Name='Channel marker tree'; Ok=(Test-OrderedNames -Text $registry -Names @('Red Ink','MSI',$Channel,$App,'Installed')) -and $registry.Contains('"Value" = "8:1"') },
    @{ Name='Opposite-channel registry search'; Ok=$content.Contains('"RegKey" = "8:' + $oppositePath + '"') },
    @{ Name='Opposite-channel property'; Ok=$content.Contains('"Property" = "8:' + $oppositeProperty + '"') },
    @{ Name='Maintenance-safe channel condition'; Ok=$content.Contains('"Condition" = "8:Installed OR NOT ' + $oppositeProperty + '"') },
    @{ Name='VSTO/manifest source architecture'; Ok=$content.Contains($sourcePart) },
    @{ Name='VSTO Runtime launch condition'; Ok=$content.Contains('VSTORUNTIMEREDIST') -and $content.Contains('OFFICERUNTIME') },
    @{ Name='Release MSI output'; Ok=$content.Contains('"OutputFilename" = "8:' + $releaseOutput + '"') },
    @{ Name='No detected-dependency assembly objects'; Ok=(-not $content.Contains('"{9F6F8455-1EF1-4B85-886A-4223BCC8E7F7}:_')) },
    @{ Name='No Primary Output objects'; Ok=(-not $content.Contains('"{5259A561-127C-4D43-A0A1-72F10C7B3BF8}:_')) }
)

foreach ($check in $checks) {
    if (-not [bool]$check.Ok) {
        throw ("Installer validation failed [{0}]: {1}" -f $check.Name, $ProjectPath)
    }
}

# The Hierarchy section is a graph over payload objects. Earlier installer revisions removed
# dependency objects but left hundreds of stale graph references behind. Require a deterministic
# one-to-one mapping: every current plain File object has exactly one hierarchy node, and no
# hierarchy node owns or references an object that does not exist.
$fileSectionMatch = [regex]::Match($content, '(?ms)^\s*"File"\s*\{(?<body>.*?)^\s*"FileType"\s*\{')
if (-not $fileSectionMatch.Success) { throw ("Installer validation failed [File section missing]: {0}" -f $ProjectPath) }
$fileSection = [string]$fileSectionMatch.Groups['body'].Value
$fileIdMatches = [regex]::Matches($fileSection, '(?m)^\s*"\{1FB2D0AE-D3B9-43D4-B9DD-F88EC61E35DE\}:(_[A-Fa-f0-9]+)"\s*$')
$fileIds = @{}
foreach ($fileIdMatch in $fileIdMatches) {
    [string]$fileId = $fileIdMatch.Groups[1].Value
    if ($fileIds.ContainsKey($fileId)) { throw ("Installer validation failed [duplicate plain file ID {0}]: {1}" -f $fileId, $ProjectPath) }
    $fileIds[$fileId] = $true
}
if ($fileIds.Count -lt 2) { throw ("Installer validation failed [expected at least .vsto and .dll.manifest plain files]: {0}" -f $ProjectPath) }

$hierarchyMatch = [regex]::Match($content, '(?ms)^\s*"Hierarchy"\s*\{(?<body>.*?)^\s*"Configurations"\s*\{')
if (-not $hierarchyMatch.Success) { throw ("Installer validation failed [Hierarchy section missing]: {0}" -f $ProjectPath) }
$hierarchy = [string]$hierarchyMatch.Groups['body'].Value
$hierarchyKeyMatches = [regex]::Matches($hierarchy, '(?m)^\s*"MsmKey"\s*=\s*"8:([^"]+)"\s*$')
$hierarchyOwnerMatches = [regex]::Matches($hierarchy, '(?m)^\s*"OwnerKey"\s*=\s*"8:([^"]+)"\s*$')
$hierarchySigMatches = [regex]::Matches($hierarchy, '(?m)^\s*"MsmSig"\s*=\s*"8:([^"]+)"\s*$')
if ($hierarchyKeyMatches.Count -ne $fileIds.Count -or $hierarchyOwnerMatches.Count -ne $fileIds.Count -or $hierarchySigMatches.Count -ne $fileIds.Count) {
    throw ("Installer validation failed [Hierarchy count mismatch: files={0}, keys={1}, owners={2}, sigs={3}]: {4}" -f $fileIds.Count, $hierarchyKeyMatches.Count, $hierarchyOwnerMatches.Count, $hierarchySigMatches.Count, $ProjectPath)
}
$hierarchyKeys = @{}
foreach ($hierarchyKeyMatch in $hierarchyKeyMatches) {
    [string]$hierarchyKey = $hierarchyKeyMatch.Groups[1].Value
    if (-not $fileIds.ContainsKey($hierarchyKey)) { throw ("Installer validation failed [dangling Hierarchy MsmKey {0}]: {1}" -f $hierarchyKey, $ProjectPath) }
    if ($hierarchyKeys.ContainsKey($hierarchyKey)) { throw ("Installer validation failed [duplicate Hierarchy MsmKey {0}]: {1}" -f $hierarchyKey, $ProjectPath) }
    $hierarchyKeys[$hierarchyKey] = $true
}
foreach ($fileId in @($fileIds.Keys)) {
    if (-not $hierarchyKeys.ContainsKey($fileId)) { throw ("Installer validation failed [plain file {0} missing from Hierarchy]: {1}" -f $fileId, $ProjectPath) }
}
foreach ($hierarchyOwnerMatch in $hierarchyOwnerMatches) {
    if ($hierarchyOwnerMatch.Groups[1].Value -ne '_UNDEFINED') { throw ("Installer validation failed [unexpected Hierarchy OwnerKey {0}]: {1}" -f $hierarchyOwnerMatch.Groups[1].Value, $ProjectPath) }
}
foreach ($hierarchySigMatch in $hierarchySigMatches) {
    if ($hierarchySigMatch.Groups[1].Value -ne '_UNDEFINED') { throw ("Installer validation failed [unexpected Hierarchy MsmSig {0}]: {1}" -f $hierarchySigMatch.Groups[1].Value, $ProjectPath) }
}

# Core MSI identity fields must be singular and syntactically valid.
foreach ($identityName in @('ProductCode','PackageCode','UpgradeCode')) {
    $identityMatches = [regex]::Matches($content, ('(?m)^\s*"{0}"\s*=\s*"8:(\{{[0-9A-Fa-f-]+\}})"\s*$' -f $identityName))
    if ($identityMatches.Count -ne 1) { throw ("Installer validation failed [{0} occurrence count {1}]: {2}" -f $identityName, $identityMatches.Count, $ProjectPath) }
    $parsedGuid = [System.Guid]::Empty
    if (-not [System.Guid]::TryParse($identityMatches[0].Groups[1].Value, [ref]$parsedGuid)) { throw ("Installer validation failed [{0} invalid GUID]: {1}" -f $identityName, $ProjectPath) }
}
$productVersionMatches = [regex]::Matches($content, '(?m)^\s*"ProductVersion"\s*=\s*"8:(\d+)\.(\d+)\.(\d+)"\s*$')
if ($productVersionMatches.Count -ne 1) { throw ("Installer validation failed [ProductVersion occurrence/format]: {0}" -f $ProjectPath) }
[int]$pvMajor = [int]$productVersionMatches[0].Groups[1].Value
[int]$pvMinor = [int]$productVersionMatches[0].Groups[2].Value
[int]$pvBuild = [int]$productVersionMatches[0].Groups[3].Value
if ($pvMajor -gt 255 -or $pvMinor -gt 255 -or $pvBuild -gt 65535) { throw ("Installer validation failed [ProductVersion MSI bounds]: {0}" -f $ProjectPath) }

# Prevent known cross-host cloning regressions specifically in the registry tree.
foreach ($otherApp in @('Word','Excel','Outlook') | Where-Object { $_ -ne $App }) {
    if (Test-OrderedNames -Text $registry -Names @('Microsoft','Office',$otherApp,'Addins',"Red Ink for $otherApp")) {
        throw "Installer validation failed: $ProjectPath still contains $otherApp Office registration."
    }
}

# Word packages intentionally retain the VC++ v14 prerequisite/launch condition.
if ($App -eq 'Word' -and -not ($content -match '(?i)(Visual C\+\+|VC\+\+|VCRUNTIME|VCREDIST|14\.0)')) {
    Write-Warning "Word installer does not expose an obvious VC++ v14 marker in vdproj text. Verify prerequisite manually if Setup.exe bootstrapper is used."
}

Write-Output ([pscustomobject]@{ ProjectPath=$ProjectPath; Valid=$true })
