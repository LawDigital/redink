param(
    [Parameter(Mandatory=$true)][ValidateSet('Word','Excel','Outlook')][string]$App,
    [Parameter(Mandatory=$true)][ValidateSet('x86','x64')][string]$Arch,
    [Parameter(Mandatory=$true)][string]$ProjectPath,
    [Parameter(Mandatory=$true)][string]$SolutionRoot
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$fileObjectType = '{1FB2D0AE-D3B9-43D4-B9DD-F88EC61E35DE}'
$assemblyObjectType = '{9F6F8455-1EF1-4B85-886A-4223BCC8E7F7}'
$projectOutputObjectType = '{5259A561-127C-4D43-A0A1-72F10C7B3BF8}'
$customFolderObjectType = '{9EF0B969-E518-4E46-987F-47570745A589}'
$generatedTag = 'REDINK_AUTOPAYLOAD'

$packagingRoot = Split-Path -Parent $PSScriptRoot
[string]$packagingRootFull = [System.IO.Path]::GetFullPath($packagingRoot)
if (-not $packagingRootFull.EndsWith([System.IO.Path]::DirectorySeparatorChar.ToString())) {
    $packagingRootFull += [System.IO.Path]::DirectorySeparatorChar
}
[string]$projectPathFull = [System.IO.Path]::GetFullPath($ProjectPath)
if (-not $projectPathFull.StartsWith($packagingRootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw ("Safety check failed: Repair-VdprojDependencies.ps1 may modify only files under Red Ink Installer: {0}" -f $projectPathFull)
}

function Get-RelativePathPs51 {
    param([Parameter(Mandatory=$true)][string]$BaseDirectory,[Parameter(Mandatory=$true)][string]$TargetPath)
    [string]$baseFull = [System.IO.Path]::GetFullPath($BaseDirectory)
    [string]$targetFull = [System.IO.Path]::GetFullPath($TargetPath)
    if (-not $baseFull.EndsWith([System.IO.Path]::DirectorySeparatorChar.ToString())) {
        $baseFull += [System.IO.Path]::DirectorySeparatorChar
    }
    $baseUri = New-Object System.Uri($baseFull)
    $targetUri = New-Object System.Uri($targetFull)
    [string]$relativePath = [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString())
    return $relativePath.Replace('/', [System.IO.Path]::DirectorySeparatorChar)
}

function Get-BlockEnd {
    param(
        [Parameter(Mandatory=$true)][System.Collections.Generic.List[string]]$Lines,
        [Parameter(Mandatory=$true)][int]$HeaderIndex,
        [Parameter(Mandatory=$true)][string]$Description
    )
    if (($HeaderIndex + 1) -ge $Lines.Count -or $Lines[$HeaderIndex + 1].Trim() -ne '{') {
        throw ("Malformed {0} at line {1} in {2}" -f $Description, ($HeaderIndex + 1), $ProjectPath)
    }
    [int]$depth = 0
    for ([int]$scanIndex = $HeaderIndex + 1; $scanIndex -lt $Lines.Count; $scanIndex++) {
        [string]$trimmedLine = $Lines[$scanIndex].Trim()
        if ($trimmedLine -eq '{') { $depth++ }
        elseif ($trimmedLine -eq '}') {
            $depth--
            if ($depth -eq 0) { return $scanIndex }
        }
    }
    throw ("Unterminated {0} at line {1} in {2}" -f $Description, ($HeaderIndex + 1), $ProjectPath)
}

function Get-NamedSectionBounds {
    param(
        [Parameter(Mandatory=$true)][System.Collections.Generic.List[string]]$Lines,
        [Parameter(Mandatory=$true)][string]$SectionName
    )
    [int]$headerIndex = -1
    for ([int]$lineIndex = 0; $lineIndex -lt $Lines.Count; $lineIndex++) {
        if ($Lines[$lineIndex].Trim() -eq ('"' + $SectionName + '"')) {
            $headerIndex = $lineIndex
            break
        }
    }
    if ($headerIndex -lt 0) {
        throw ("VDPROJ section '{0}' was not found in {1}." -f $SectionName, $ProjectPath)
    }
    [int]$endIndex = Get-BlockEnd -Lines $Lines -HeaderIndex $headerIndex -Description ("VDPROJ section '{0}'" -f $SectionName)
    return [pscustomobject]@{ Header = $headerIndex; End = $endIndex }
}

function Get-StableObjectId {
    param([Parameter(Mandatory=$true)][string]$Seed)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        [byte[]]$bytes = [System.Text.Encoding]::UTF8.GetBytes($Seed)
        [byte[]]$hashBytes = $sha256.ComputeHash($bytes)
        [System.Text.StringBuilder]$builder = New-Object System.Text.StringBuilder
        foreach ($hashByte in $hashBytes) {
            [void]$builder.Append($hashByte.ToString('X2'))
        }
        return '_' + $builder.ToString().Substring(0, 32)
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-CopyToOutputFileNames {
    param([Parameter(Mandatory=$true)][string]$ProjectFile)
    [xml]$projectXml = Get-Content -LiteralPath $ProjectFile -Raw
    $names = New-Object 'System.Collections.Generic.List[string]'
    $itemNodes = @($projectXml.SelectNodes("/*[local-name()='Project']/*[local-name()='ItemGroup']/*"))
    foreach ($itemNode in $itemNodes) {
        $copyNode = $itemNode.SelectSingleNode("*[local-name()='CopyToOutputDirectory']")
        if ($null -eq $copyNode) { continue }
        [string]$copyMode = $copyNode.InnerText.Trim()
        if ([string]::IsNullOrWhiteSpace($copyMode) -or $copyMode -eq 'Never') { continue }
        $includeAttribute = $itemNode.Attributes['Include']
        if ($null -eq $includeAttribute -or [string]::IsNullOrWhiteSpace([string]$includeAttribute.Value)) { continue }
        [string]$fileName = [System.IO.Path]::GetFileName([string]$includeAttribute.Value)
        if (-not [string]::IsNullOrWhiteSpace($fileName) -and -not $names.Contains($fileName)) {
            [void]$names.Add($fileName)
        }
    }
    return $names.ToArray()
}

function Get-PlainFileObjects {
    param([Parameter(Mandatory=$true)][System.Collections.Generic.List[string]]$Lines)
    $fileBounds = Get-NamedSectionBounds -Lines $Lines -SectionName 'File'
    $result = New-Object 'System.Collections.Generic.List[object]'
    for ([int]$lineIndex = $fileBounds.Header + 2; $lineIndex -lt $fileBounds.End;) {
        [string]$header = $Lines[$lineIndex].Trim()
        if ($header.StartsWith('"' + $fileObjectType + ':_', [System.StringComparison]::Ordinal)) {
            [int]$objectEnd = Get-BlockEnd -Lines $Lines -HeaderIndex $lineIndex -Description 'VDPROJ file object'
            $idMatch = [regex]::Match($header, '^"\{1FB2D0AE-D3B9-43D4-B9DD-F88EC61E35DE\}:(_[A-Fa-f0-9]+)"$')
            if (-not $idMatch.Success) {
                throw ("Could not parse plain file object ID in {0}: {1}" -f $ProjectPath, $header)
            }
            [string]$targetName = ''
            [string]$folderId = ''
            [string]$tag = ''
            [string]$sourcePath = ''
            for ([int]$propertyIndex = $lineIndex + 1; $propertyIndex -le $objectEnd; $propertyIndex++) {
                $targetMatch = [regex]::Match($Lines[$propertyIndex], '^\s*"TargetName"\s*=\s*"8:([^"]+)"\s*$')
                if ($targetMatch.Success) { $targetName = $targetMatch.Groups[1].Value }
                $folderMatch = [regex]::Match($Lines[$propertyIndex], '^\s*"Folder"\s*=\s*"8:([^"]+)"\s*$')
                if ($folderMatch.Success) { $folderId = $folderMatch.Groups[1].Value }
                $tagMatch = [regex]::Match($Lines[$propertyIndex], '^\s*"Tag"\s*=\s*"8:([^"]*)"\s*$')
                if ($tagMatch.Success) { $tag = $tagMatch.Groups[1].Value }
                $sourceMatch = [regex]::Match($Lines[$propertyIndex], '^\s*"SourcePath"\s*=\s*"8:([^"]+)"\s*$')
                if ($sourceMatch.Success) { $sourcePath = $sourceMatch.Groups[1].Value }
            }
            [void]$result.Add([pscustomobject]@{
                Id = $idMatch.Groups[1].Value
                TargetName = $targetName
                FolderId = $folderId
                Tag = $tag
                SourcePath = $sourcePath
                Header = $lineIndex
                End = $objectEnd
            })
            $lineIndex = $objectEnd + 1
            continue
        }
        $lineIndex++
    }
    return $result.ToArray()
}

function Get-ApplicationFolderInfo {
    param([Parameter(Mandatory=$true)][System.Collections.Generic.List[string]]$Lines)
    $folderBounds = Get-NamedSectionBounds -Lines $Lines -SectionName 'Folder'
    for ([int]$lineIndex = $folderBounds.Header + 2; $lineIndex -lt $folderBounds.End;) {
        [string]$header = $Lines[$lineIndex].Trim()
        if ($header -match '^"\{[0-9A-Fa-f-]+\}:(_[A-Fa-f0-9]+)"$') {
            [int]$objectEnd = Get-BlockEnd -Lines $Lines -HeaderIndex $lineIndex -Description 'VDPROJ folder object'
            [bool]$isTargetDir = $false
            [int]$foldersHeader = -1
            [int]$foldersEnd = -1
            for ([int]$propertyIndex = $lineIndex + 1; $propertyIndex -le $objectEnd; $propertyIndex++) {
                if ($Lines[$propertyIndex].Trim() -eq '"Property" = "8:TARGETDIR"') { $isTargetDir = $true }
                if ($Lines[$propertyIndex].Trim() -eq '"Folders"') {
                    $foldersHeader = $propertyIndex
                    $foldersEnd = Get-BlockEnd -Lines $Lines -HeaderIndex $foldersHeader -Description 'Application Folder child Folders block'
                    break
                }
            }
            if ($isTargetDir) {
                if ($foldersHeader -lt 0 -or $foldersEnd -lt 0) {
                    throw ("Application Folder in {0} has no Folders block." -f $ProjectPath)
                }
                $idMatch = [regex]::Match($header, ':(_[A-Fa-f0-9]+)"$')
                return [pscustomobject]@{
                    Id = $idMatch.Groups[1].Value
                    Header = $lineIndex
                    End = $objectEnd
                    FoldersHeader = $foldersHeader
                    FoldersEnd = $foldersEnd
                }
            }
            $lineIndex = $objectEnd + 1
            continue
        }
        $lineIndex++
    }
    throw ("Application Folder (Property TARGETDIR) was not found in {0}." -f $ProjectPath)
}

function New-GeneratedFolderLines {
    param(
        [Parameter(Mandatory=$true)][AllowEmptyString()][string]$ParentPath,
        [Parameter(Mandatory=$true)][string]$Indent,
        [Parameter(Mandatory=$true)][hashtable]$FolderIdByPath,
        [Parameter(Mandatory=$true)][hashtable]$FolderPropertyByPath,
        [Parameter(Mandatory=$true)][string[]]$AllFolderPaths
    )
    $result = New-Object 'System.Collections.Generic.List[string]'
    $children = New-Object 'System.Collections.Generic.List[string]'
    foreach ($folderPath in $AllFolderPaths) {
        [string]$parent = [System.IO.Path]::GetDirectoryName($folderPath)
        if ($null -eq $parent) { $parent = '' }
        if ($parent -eq $ParentPath) { [void]$children.Add($folderPath) }
    }
    $childArray = @($children.ToArray() | Sort-Object)
    foreach ($childPath in $childArray) {
        [string]$leafName = [System.IO.Path]::GetFileName($childPath)
        [string]$folderId = [string]$FolderIdByPath[$childPath.ToLowerInvariant()]
        [string]$propertyId = [string]$FolderPropertyByPath[$childPath.ToLowerInvariant()]
        [void]$result.Add(($Indent + '"' + $customFolderObjectType + ':' + $folderId + '"'))
        [void]$result.Add(($Indent + '{'))
        [void]$result.Add(($Indent + '"Name" = "8:' + $leafName + '"'))
        [void]$result.Add(($Indent + '"AlwaysCreate" = "11:FALSE"'))
        [void]$result.Add(($Indent + '"Condition" = "8:"'))
        [void]$result.Add(($Indent + '"Transitive" = "11:FALSE"'))
        [void]$result.Add(($Indent + '"Property" = "8:' + $propertyId + '"'))
        [void]$result.Add(($Indent + '    "Folders"'))
        [void]$result.Add(($Indent + '    {'))
        $grandChildren = New-GeneratedFolderLines -ParentPath $childPath -Indent ($Indent + '        ') -FolderIdByPath $FolderIdByPath -FolderPropertyByPath $FolderPropertyByPath -AllFolderPaths $AllFolderPaths
        foreach ($grandChild in @($grandChildren)) { [void]$result.Add($grandChild) }
        [void]$result.Add(($Indent + '    }'))
        [void]$result.Add(($Indent + '}'))
    }
    return $result.ToArray()
}

$appProject = Join-Path $SolutionRoot ("Red Ink for {0}\Red Ink for {0}.vbproj" -f $App)
$sharedProject = Join-Path $SolutionRoot 'SharedLibrary\SharedLibrary.vbproj'
if (-not (Test-Path -LiteralPath $appProject -PathType Leaf)) { throw ("Source project not found: {0}" -f $appProject) }
if (-not (Test-Path -LiteralPath $sharedProject -PathType Leaf)) { throw ("SharedLibrary project not found: {0}" -f $sharedProject) }
if (-not (Test-Path -LiteralPath $ProjectPath -PathType Leaf)) { throw ("Installer project not found: {0}" -f $ProjectPath) }

$configurationPart = if ($Arch -eq 'x64') { 'bin\x64\Release' } else { 'bin\Release' }
$appOutput = Join-Path (Split-Path -Parent $appProject) $configurationPart
$sharedOutput = Join-Path (Split-Path -Parent $sharedProject) $configurationPart
$projectDir = Split-Path -Parent $ProjectPath
if (-not (Test-Path -LiteralPath $appOutput -PathType Container)) { throw ("Application output folder not found: {0}" -f $appOutput) }
if (-not (Test-Path -LiteralPath $sharedOutput -PathType Container)) { throw ("SharedLibrary output folder not found: {0}" -f $sharedOutput) }

# Build the payload from the actual output trees. Relative paths are preserved. App output is
# authoritative when both output trees contain the same relative path; SharedLibrary fills gaps.
$payloadByRelativePath = @{}
$outputFolders = @($appOutput, $sharedOutput)
foreach ($outputFolder in $outputFolders) {
    $outputFiles = @(Get-ChildItem -LiteralPath $outputFolder -Recurse -File -ErrorAction Stop)
    foreach ($payloadCandidate in $outputFiles) {
        if ($payloadCandidate.Extension.ToLowerInvariant() -eq '.pdb') { continue }
        [string]$relativePayloadPath = Get-RelativePathPs51 -BaseDirectory $outputFolder -TargetPath $payloadCandidate.FullName
        $relativePayloadPath = $relativePayloadPath.TrimStart('\')
        if ([string]::IsNullOrWhiteSpace($relativePayloadPath) -or $relativePayloadPath.StartsWith('..\', [System.StringComparison]::Ordinal)) {
            throw ("Built payload path escaped its output root: {0}" -f $payloadCandidate.FullName)
        }
        [string]$payloadPathKey = $relativePayloadPath.ToLowerInvariant()
        if (-not $payloadByRelativePath.ContainsKey($payloadPathKey)) {
            $payloadByRelativePath.Add($payloadPathKey, [pscustomobject]@{
                RelativePath = $relativePayloadPath
                FullName = $payloadCandidate.FullName
            })
        }
    }
}

$requiredRootPaths = @(
    ("Red Ink for {0}.dll" -f $App),
    ("Red Ink for {0}.vsto" -f $App),
    ("Red Ink for {0}.dll.manifest" -f $App),
    'SharedLibrary.dll'
)
foreach ($requiredRootPath in $requiredRootPaths) {
    if (-not $payloadByRelativePath.ContainsKey($requiredRootPath.ToLowerInvariant())) {
        throw ("Required built payload '{0}' was not found at the output root for {1} {2}." -f $requiredRootPath, $App, $Arch)
    }
}

# Anything explicitly marked CopyToOutputDirectory must exist somewhere in the built payload.
$declaredCopyFiles = New-Object 'System.Collections.Generic.List[string]'
foreach ($copyProject in @($appProject, $sharedProject)) {
    foreach ($declaredCopyFile in @(Get-CopyToOutputFileNames -ProjectFile $copyProject)) {
        if (-not $declaredCopyFiles.Contains($declaredCopyFile)) { [void]$declaredCopyFiles.Add($declaredCopyFile) }
    }
}
foreach ($declaredCopyFile in $declaredCopyFiles) {
    [bool]$foundDeclaredFile = $false
    foreach ($payloadEntry in @($payloadByRelativePath.Values)) {
        if ([System.IO.Path]::GetFileName([string]$payloadEntry.RelativePath).Equals($declaredCopyFile, [System.StringComparison]::OrdinalIgnoreCase)) {
            $foundDeclaredFile = $true
            break
        }
    }
    if (-not $foundDeclaredFile) {
        throw ("CopyToOutputDirectory file '{0}' is declared by the source projects but was not produced for {1} {2}." -f $declaredCopyFile, $App, $Arch)
    }
}

$lineArray = [System.IO.File]::ReadAllLines($ProjectPath)
$lines = New-Object 'System.Collections.Generic.List[string]'
foreach ($sourceLine in $lineArray) { [void]$lines.Add($sourceLine) }

# Remove every detected-dependency assembly object and every previously generated payload object.
# Hierarchy and generated application subfolders are rebuilt later, so no stale graph survives reruns.
$cleanedLines = New-Object 'System.Collections.Generic.List[string]'
for ([int]$lineIndex = 0; $lineIndex -lt $lines.Count;) {
    [string]$header = $lines[$lineIndex].Trim()
    [bool]$isAssembly = $header.StartsWith('"' + $assemblyObjectType + ':_', [System.StringComparison]::Ordinal)
    [bool]$isGeneratedFile = $false
    if ($header.StartsWith('"' + $fileObjectType + ':_', [System.StringComparison]::Ordinal)) {
        [int]$objectEnd = Get-BlockEnd -Lines $lines -HeaderIndex $lineIndex -Description 'VDPROJ file object'
        [string]$blockText = [string]::Join("`n", $lines.GetRange($lineIndex, ($objectEnd - $lineIndex + 1)).ToArray())
        $isGeneratedFile = $blockText.Contains('"Tag" = "8:' + $generatedTag + '"')
        if ($isGeneratedFile) {
            $lineIndex = $objectEnd + 1
            continue
        }
    }
    if ($isAssembly) {
        [int]$objectEnd = Get-BlockEnd -Lines $lines -HeaderIndex $lineIndex -Description 'VDPROJ detected-dependency object'
        $lineIndex = $objectEnd + 1
        continue
    }
    [void]$cleanedLines.Add($lines[$lineIndex])
    $lineIndex++
}
$lines = $cleanedLines

for ([int]$lineIndex = 0; $lineIndex -lt $lines.Count; $lineIndex++) {
    if ($lines[$lineIndex].Trim().StartsWith('"' + $projectOutputObjectType + ':_', [System.StringComparison]::Ordinal)) {
        throw ("Legacy Primary Output object remains in {0}." -f $ProjectPath)
    }
}

$appFolder = Get-ApplicationFolderInfo -Lines $lines
[string]$applicationFolderId = [string]$appFolder.Id

# Existing user-authored plain files (.vsto and .dll.manifest) must stay in the application root.
$plainFilesBeforeGeneration = @(Get-PlainFileObjects -Lines $lines)
if ($plainFilesBeforeGeneration.Count -lt 2) {
    throw ("Expected at least the user-authored .vsto and .dll.manifest file objects in {0}." -f $ProjectPath)
}
$existingRootNames = @{}
foreach ($plainFile in $plainFilesBeforeGeneration) {
    if ([string]::IsNullOrWhiteSpace([string]$plainFile.TargetName)) {
        throw ("Plain file object {0} has no TargetName in {1}." -f $plainFile.Id, $ProjectPath)
    }
    if ([string]$plainFile.FolderId -ne $applicationFolderId) {
        throw ("Unexpected pre-existing plain file outside the Application Folder in {0}: {1}. Generated subfolders are rebuilt automatically; manually authored nested file objects are not supported." -f $ProjectPath, $plainFile.TargetName)
    }
    [string]$rootNameKey = ([string]$plainFile.TargetName).ToLowerInvariant()
    if ($existingRootNames.ContainsKey($rootNameKey)) {
        throw ("Duplicate pre-existing Application Folder TargetName '{0}' in {1}." -f $plainFile.TargetName, $ProjectPath)
    }
    $existingRootNames[$rootNameKey] = $true
}

# Build a deterministic directory tree for every nested payload path.
$folderPaths = New-Object 'System.Collections.Generic.List[string]'
foreach ($payloadEntry in @($payloadByRelativePath.Values)) {
    [string]$relativePath = [string]$payloadEntry.RelativePath
    [string]$directory = [System.IO.Path]::GetDirectoryName($relativePath)
    while (-not [string]::IsNullOrWhiteSpace($directory)) {
        if (-not $folderPaths.Contains($directory)) { [void]$folderPaths.Add($directory) }
        $directory = [System.IO.Path]::GetDirectoryName($directory)
    }
}
$allFolderPaths = @($folderPaths.ToArray() | Sort-Object { ($_ -split '[\\/]').Count }, { $_ })
$folderIdByPath = @{}
$folderPropertyByPath = @{}
foreach ($folderPath in $allFolderPaths) {
    [string]$folderKey = $folderPath.ToLowerInvariant()
    $folderIdByPath[$folderKey] = Get-StableObjectId -Seed (($ProjectPath.ToLowerInvariant()) + '|folder|' + $folderKey)
    $folderPropertyByPath[$folderKey] = Get-StableObjectId -Seed (($ProjectPath.ToLowerInvariant()) + '|folder-property|' + $folderKey)
}

# Replace the Application Folder's child-folder tree on every run. This removes stale generated
# subfolders and recreates exactly the directory layout emitted by the current source build.
$generatedFolderLines = New-Object 'System.Collections.Generic.List[string]'
$folderTreeLines = New-GeneratedFolderLines -ParentPath '' -Indent '                        ' -FolderIdByPath $folderIdByPath -FolderPropertyByPath $folderPropertyByPath -AllFolderPaths $allFolderPaths
foreach ($folderTreeLine in @($folderTreeLines)) { [void]$generatedFolderLines.Add($folderTreeLine) }
$newLines = New-Object 'System.Collections.Generic.List[string]'
for ([int]$lineIndex = 0; $lineIndex -lt ($appFolder.FoldersHeader + 2); $lineIndex++) { [void]$newLines.Add($lines[$lineIndex]) }
foreach ($folderLine in $generatedFolderLines) { [void]$newLines.Add($folderLine) }
for ([int]$lineIndex = $appFolder.FoldersEnd; $lineIndex -lt $lines.Count; $lineIndex++) { [void]$newLines.Add($lines[$lineIndex]) }
$lines = $newLines

# Add every built payload path not already represented by a user-authored root file.
$fileBounds = Get-NamedSectionBounds -Lines $lines -SectionName 'File'
$generatedLines = New-Object 'System.Collections.Generic.List[string]'
[int]$generatedCount = 0
$payloadKeys = @($payloadByRelativePath.Keys | Sort-Object)
foreach ($payloadKey in $payloadKeys) {
    $payloadEntry = $payloadByRelativePath[$payloadKey]
    [string]$relativePayloadPath = [string]$payloadEntry.RelativePath
    [string]$sourcePath = [string]$payloadEntry.FullName
    [string]$targetName = [System.IO.Path]::GetFileName($relativePayloadPath)
    [string]$targetDirectory = [System.IO.Path]::GetDirectoryName($relativePayloadPath)
    if ($null -eq $targetDirectory) { $targetDirectory = '' }

    if ([string]::IsNullOrWhiteSpace($targetDirectory) -and $existingRootNames.ContainsKey($targetName.ToLowerInvariant())) {
        continue
    }

    [string]$targetFolderId = $applicationFolderId
    if (-not [string]::IsNullOrWhiteSpace($targetDirectory)) {
        [string]$folderKey = $targetDirectory.ToLowerInvariant()
        if (-not $folderIdByPath.ContainsKey($folderKey)) {
            throw ("No generated installer folder exists for payload path '{0}' in {1}." -f $relativePayloadPath, $ProjectPath)
        }
        $targetFolderId = [string]$folderIdByPath[$folderKey]
    }

    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw ("Built payload disappeared before packaging: {0}" -f $sourcePath)
    }
    [string]$relativeSource = Get-RelativePathPs51 -BaseDirectory $projectDir -TargetPath $sourcePath
    [string]$vdprojSource = $relativeSource.Replace('\','\\')
    [string]$objectId = Get-StableObjectId -Seed (($ProjectPath.ToLowerInvariant()) + '|file|' + $payloadKey)

    [void]$generatedLines.Add(('            "{0}:{1}"' -f $fileObjectType, $objectId))
    [void]$generatedLines.Add('            {')
    [void]$generatedLines.Add(('            "SourcePath" = "8:{0}"' -f $vdprojSource))
    [void]$generatedLines.Add(('            "TargetName" = "8:{0}"' -f $targetName))
    [void]$generatedLines.Add(('            "Tag" = "8:{0}"' -f $generatedTag))
    [void]$generatedLines.Add(('            "Folder" = "8:{0}"' -f $targetFolderId))
    [void]$generatedLines.Add('            "Condition" = "8:"')
    [void]$generatedLines.Add('            "Transitive" = "11:FALSE"')
    [void]$generatedLines.Add('            "Vital" = "11:TRUE"')
    [void]$generatedLines.Add('            "ReadOnly" = "11:FALSE"')
    [void]$generatedLines.Add('            "Hidden" = "11:FALSE"')
    [void]$generatedLines.Add('            "System" = "11:FALSE"')
    [void]$generatedLines.Add('            "Permanent" = "11:FALSE"')
    [void]$generatedLines.Add('            "SharedLegacy" = "11:FALSE"')
    [void]$generatedLines.Add('            "PackageAs" = "3:1"')
    [void]$generatedLines.Add('            "Register" = "3:1"')
    [void]$generatedLines.Add('            "Exclude" = "11:FALSE"')
    [void]$generatedLines.Add('            "IsDependency" = "11:FALSE"')
    [void]$generatedLines.Add('            "IsolateTo" = "8:"')
    [void]$generatedLines.Add('            }')
    $generatedCount++
}

if ($generatedLines.Count -gt 0) {
    $newLines = New-Object 'System.Collections.Generic.List[string]'
    for ([int]$lineIndex = 0; $lineIndex -lt $fileBounds.End; $lineIndex++) { [void]$newLines.Add($lines[$lineIndex]) }
    foreach ($generatedLine in $generatedLines) { [void]$newLines.Add($generatedLine) }
    for ([int]$lineIndex = $fileBounds.End; $lineIndex -lt $lines.Count; $lineIndex++) { [void]$newLines.Add($lines[$lineIndex]) }
    $lines = $newLines
}

# Rebuild Hierarchy from the CURRENT plain file objects. No removed dependency or generated-file
# object can remain referenced after this point.
$plainFiles = @(Get-PlainFileObjects -Lines $lines)
if ($plainFiles.Count -lt 2) {
    throw ("Unexpectedly few plain file objects in {0}: {1}." -f $ProjectPath, $plainFiles.Count)
}
$seenIds = @{}
$seenFolderNamePairs = @{}
foreach ($plainFile in $plainFiles) {
    [string]$plainId = [string]$plainFile.Id
    [string]$plainName = [string]$plainFile.TargetName
    if ($seenIds.ContainsKey($plainId)) { throw ("Duplicate plain file object ID {0} in {1}." -f $plainId, $ProjectPath) }
    $seenIds[$plainId] = $true
    [string]$pairKey = (([string]$plainFile.FolderId) + '|' + $plainName).ToLowerInvariant()
    if ($seenFolderNamePairs.ContainsKey($pairKey)) { throw ("Duplicate installer file '{0}' in folder {1} in {2}." -f $plainName, $plainFile.FolderId, $ProjectPath) }
    $seenFolderNamePairs[$pairKey] = $true
}

$hierarchyBounds = Get-NamedSectionBounds -Lines $lines -SectionName 'Hierarchy'
$hierarchyLines = New-Object 'System.Collections.Generic.List[string]'
foreach ($plainFile in $plainFiles) {
    [void]$hierarchyLines.Add('        "Entry"')
    [void]$hierarchyLines.Add('        {')
    [void]$hierarchyLines.Add(('        "MsmKey" = "8:{0}"' -f $plainFile.Id))
    [void]$hierarchyLines.Add('        "OwnerKey" = "8:_UNDEFINED"')
    [void]$hierarchyLines.Add('        "MsmSig" = "8:_UNDEFINED"')
    [void]$hierarchyLines.Add('        }')
}
$newLines = New-Object 'System.Collections.Generic.List[string]'
for ([int]$lineIndex = 0; $lineIndex -lt ($hierarchyBounds.Header + 2); $lineIndex++) { [void]$newLines.Add($lines[$lineIndex]) }
foreach ($hierarchyLine in $hierarchyLines) { [void]$newLines.Add($hierarchyLine) }
for ([int]$lineIndex = $hierarchyBounds.End; $lineIndex -lt $lines.Count; $lineIndex++) { [void]$newLines.Add($lines[$lineIndex]) }
$lines = $newLines

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($ProjectPath, $lines.ToArray(), $utf8Bom)

# Post-write assertions.
[string]$content = [System.IO.File]::ReadAllText($ProjectPath)
if ($content.Contains('"' + $assemblyObjectType + ':_')) { throw ("Detected-dependency assembly objects remain in {0}." -f $ProjectPath) }
if ($content.Contains('"' + $projectOutputObjectType + ':_')) { throw ("Primary Output objects remain in {0}." -f $ProjectPath) }

$verifyArray = [System.IO.File]::ReadAllLines($ProjectPath)
$verifyLines = New-Object 'System.Collections.Generic.List[string]'
foreach ($verifyLine in $verifyArray) { [void]$verifyLines.Add($verifyLine) }
$verifyPlainFiles = @(Get-PlainFileObjects -Lines $verifyLines)

foreach ($requiredRootPath in $requiredRootPaths) {
    [int]$requiredCount = 0
    foreach ($plainFile in $verifyPlainFiles) {
        if (([string]$plainFile.FolderId -eq $applicationFolderId) -and ([string]$plainFile.TargetName).Equals($requiredRootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            $requiredCount++
        }
    }
    if ($requiredCount -ne 1) {
        throw ("Required root payload '{0}' is present {1} time(s) in {2}; expected exactly one." -f $requiredRootPath, $requiredCount, $ProjectPath)
    }
}

foreach ($payloadKey in $payloadKeys) {
    $payloadEntry = $payloadByRelativePath[$payloadKey]
    [string]$relativePayloadPath = [string]$payloadEntry.RelativePath
    [string]$targetName = [System.IO.Path]::GetFileName($relativePayloadPath)
    [string]$targetDirectory = [System.IO.Path]::GetDirectoryName($relativePayloadPath)
    if ($null -eq $targetDirectory) { $targetDirectory = '' }
    [string]$expectedFolderId = $applicationFolderId
    if (-not [string]::IsNullOrWhiteSpace($targetDirectory)) { $expectedFolderId = [string]$folderIdByPath[$targetDirectory.ToLowerInvariant()] }
    [int]$matchCount = 0
    foreach ($plainFile in $verifyPlainFiles) {
        if (([string]$plainFile.FolderId -eq $expectedFolderId) -and ([string]$plainFile.TargetName).Equals($targetName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $matchCount++
        }
    }
    if ($matchCount -ne 1) {
        throw ("Built payload '{0}' is present {1} time(s) in its expected installer folder in {2}; expected exactly one." -f $relativePayloadPath, $matchCount, $ProjectPath)
    }
}

$verifyHierarchy = Get-NamedSectionBounds -Lines $verifyLines -SectionName 'Hierarchy'
$hierarchyKeys = New-Object 'System.Collections.Generic.List[string]'
$hierarchyOwners = New-Object 'System.Collections.Generic.List[string]'
$hierarchySigs = New-Object 'System.Collections.Generic.List[string]'
for ([int]$lineIndex = $verifyHierarchy.Header + 2; $lineIndex -lt $verifyHierarchy.End; $lineIndex++) {
    $msmKeyMatch = [regex]::Match($verifyLines[$lineIndex], '^\s*"MsmKey"\s*=\s*"8:([^"]+)"\s*$')
    if ($msmKeyMatch.Success) { [void]$hierarchyKeys.Add($msmKeyMatch.Groups[1].Value) }
    $ownerMatch = [regex]::Match($verifyLines[$lineIndex], '^\s*"OwnerKey"\s*=\s*"8:([^"]+)"\s*$')
    if ($ownerMatch.Success) { [void]$hierarchyOwners.Add($ownerMatch.Groups[1].Value) }
    $sigMatch = [regex]::Match($verifyLines[$lineIndex], '^\s*"MsmSig"\s*=\s*"8:([^"]+)"\s*$')
    if ($sigMatch.Success) { [void]$hierarchySigs.Add($sigMatch.Groups[1].Value) }
}
if ($hierarchyKeys.Count -ne $verifyPlainFiles.Count -or $hierarchyOwners.Count -ne $verifyPlainFiles.Count -or $hierarchySigs.Count -ne $verifyPlainFiles.Count) {
    throw ("Hierarchy/file count mismatch in {0}: files={1}, keys={2}, owners={3}, sigs={4}." -f $ProjectPath, $verifyPlainFiles.Count, $hierarchyKeys.Count, $hierarchyOwners.Count, $hierarchySigs.Count)
}
foreach ($plainFile in $verifyPlainFiles) {
    [int]$keyCount = 0
    foreach ($hierarchyKey in $hierarchyKeys) { if ($hierarchyKey -eq $plainFile.Id) { $keyCount++ } }
    if ($keyCount -ne 1) { throw ("Plain file object {0} has {1} hierarchy nodes in {2}." -f $plainFile.Id, $keyCount, $ProjectPath) }
}
foreach ($hierarchyOwner in $hierarchyOwners) {
    if ($hierarchyOwner -ne '_UNDEFINED') { throw ("Unexpected non-undefined Hierarchy OwnerKey '{0}' in {1}." -f $hierarchyOwner, $ProjectPath) }
}
foreach ($hierarchySig in $hierarchySigs) {
    if ($hierarchySig -ne '_UNDEFINED') { throw ("Unexpected non-undefined Hierarchy MsmSig '{0}' in {1}." -f $hierarchySig, $ProjectPath) }
}

[pscustomobject]@{
    ProjectPath = $ProjectPath
    ExplicitPayloadFilesAdded = $generatedCount
    TotalBuiltPayloadFiles = $payloadByRelativePath.Count
    GeneratedSubfolders = $allFolderPaths.Count
    HierarchyEntries = $hierarchyKeys.Count
}
