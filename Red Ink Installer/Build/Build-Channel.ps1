param(
    [Parameter(Mandatory=$true)][ValidateSet('Preview','GA')][string]$Channel
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$packagingRoot = Split-Path -Parent $PSScriptRoot
& (Join-Path $PSScriptRoot 'Test-PowerShellSyntax.ps1') | Out-Null
$solutionRoot = Split-Path -Parent $packagingRoot
$solution = Join-Path $solutionRoot 'Red Ink.sln'
if (-not (Test-Path -LiteralPath $solution)) {
    throw "Red Ink.sln not found at '$solution'. Put the 'Red Ink Installer' folder directly beside Red Ink.sln."
}

function Assert-InPackagingRoot {
    param([Parameter(Mandatory=$true)][string]$PathToCheck)
    [string]$root = [System.IO.Path]::GetFullPath($packagingRoot)
    [string]$candidate = [System.IO.Path]::GetFullPath($PathToCheck)
    if (-not $root.EndsWith([System.IO.Path]::DirectorySeparatorChar.ToString())) {
        $root += [System.IO.Path]::DirectorySeparatorChar
    }
    if (-not $candidate.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Safety check failed: write/delete target is outside Red Ink Installer: $candidate"
    }
}

$ver = & (Join-Path $PSScriptRoot 'Get-ChannelVersion.ps1') -Channel $Channel -SolutionRoot $solutionRoot
if ($null -eq $ver -or [string]::IsNullOrWhiteSpace([string]$ver.MsiVersion)) {
    throw 'Could not determine the VSTO/MSI version.'
}
Write-Host "$Channel VSTO ApplicationVersion: $($ver.ApplicationVersion)"
Write-Host "MSI ProductVersion: $($ver.MsiVersion)"

$vs = & (Join-Path $PSScriptRoot 'Find-VisualStudio2022.ps1')
Write-Host "Visual Studio: $($vs.InstallationPath)"

$disableOutOfProcTool = [string]$vs.DisableOutOfProcBuild
if ([string]::IsNullOrWhiteSpace($disableOutOfProcTool) -or -not (Test-Path -LiteralPath $disableOutOfProcTool -PathType Leaf)) {
    throw ("Microsoft DisableOutOfProcBuild.exe not found for Visual Studio installation '{0}'." -f $vs.InstallationPath)
}

# Build only the four VSTO-related projects. Do not build unrelated projects in Red Ink.sln.
# Remove only compiler-generated Release bin/obj folders first. This prevents obsolete payload
# files from an older build from being swept into the MSI. No source/config/project file is deleted.
$generatedReleaseDirs = @(
    (Join-Path $solutionRoot 'SharedLibrary\bin\Release'),
    (Join-Path $solutionRoot 'SharedLibrary\obj\Release'),
    (Join-Path $solutionRoot 'SharedLibrary\bin\x64\Release'),
    (Join-Path $solutionRoot 'SharedLibrary\obj\x64\Release'),
    (Join-Path $solutionRoot 'Red Ink for Word\bin\Release'),
    (Join-Path $solutionRoot 'Red Ink for Word\obj\Release'),
    (Join-Path $solutionRoot 'Red Ink for Word\bin\x64\Release'),
    (Join-Path $solutionRoot 'Red Ink for Word\obj\x64\Release'),
    (Join-Path $solutionRoot 'Red Ink for Excel\bin\Release'),
    (Join-Path $solutionRoot 'Red Ink for Excel\obj\Release'),
    (Join-Path $solutionRoot 'Red Ink for Excel\bin\x64\Release'),
    (Join-Path $solutionRoot 'Red Ink for Excel\obj\x64\Release'),
    (Join-Path $solutionRoot 'Red Ink for Outlook\bin\Release'),
    (Join-Path $solutionRoot 'Red Ink for Outlook\obj\Release'),
    (Join-Path $solutionRoot 'Red Ink for Outlook\bin\x64\Release'),
    (Join-Path $solutionRoot 'Red Ink for Outlook\obj\x64\Release')
)
foreach ($generatedReleaseDir in $generatedReleaseDirs) {
    if (Test-Path -LiteralPath $generatedReleaseDir -PathType Container) {
        Remove-Item -LiteralPath $generatedReleaseDir -Recurse -Force
    }
}
$sourceProjects = @(
    (Join-Path $solutionRoot 'SharedLibrary\SharedLibrary.vbproj'),
    (Join-Path $solutionRoot 'Red Ink for Word\Red Ink for Word.vbproj'),
    (Join-Path $solutionRoot 'Red Ink for Excel\Red Ink for Excel.vbproj'),
    (Join-Path $solutionRoot 'Red Ink for Outlook\Red Ink for Outlook.vbproj')
)
$builds = @(
    @{ Platform = 'AnyCPU'; Label = 'Any CPU' },
    @{ Platform = 'x64'; Label = 'x64' }
)
foreach ($build in $builds) {
    foreach ($sourceProject in $sourceProjects) {
        if (-not (Test-Path -LiteralPath $sourceProject)) { throw "Source project not found: $sourceProject" }
        Write-Host "Rebuilding $([System.IO.Path]::GetFileNameWithoutExtension($sourceProject)): Release|$($build.Label), RedInkEnvironment=$Channel"
        & $vs.MSBuild $sourceProject '/t:Rebuild' '/p:Configuration=Release' ("/p:Platform={0}" -f $build.Platform) ("/p:RedInkEnvironment={0}" -f $Channel) '/m'
        if ($LASTEXITCODE -ne 0) { throw "$sourceProject build failed with exit code $LASTEXITCODE." }
    }
}

$matrix = @(
    @{ App='Word'; Arch='x86' }, @{ App='Word'; Arch='x64' },
    @{ App='Excel'; Arch='x86' }, @{ App='Excel'; Arch='x64' },
    @{ App='Outlook'; Arch='x86' }, @{ App='Outlook'; Arch='x64' }
)
$projectsRoot = Join-Path $packagingRoot 'InstallerProjects\Projects'

# Validate every installer project before modifying/building any of them.
$items = @()
foreach ($item in $matrix) {
    $name = "$($item.App)-$($item.Arch)"
    $projectDir = Join-Path $projectsRoot "$Channel-$name"
    $vdproj = Join-Path $projectDir "RedInk-$Channel-$name.vdproj"
    if (-not (Test-Path -LiteralPath $vdproj)) {
        throw "Missing installer project: $vdproj`nComplete InstallerProjects\ONE-TIME-SETUP.md first."
    }
    Assert-InPackagingRoot -PathToCheck $vdproj
    $repair = & (Join-Path $PSScriptRoot 'Repair-VdprojDependencies.ps1') -App $item.App -Arch $item.Arch -ProjectPath $vdproj -SolutionRoot $solutionRoot
    Write-Host "Prepared $name payload: $($repair.TotalBuiltPayloadFiles) built runtime file(s); $($repair.ExplicitPayloadFilesAdded) added as ordinary installer files; no Primary Output / detected-dependency objects."
    & (Join-Path $PSScriptRoot 'Validate-Vdproj.ps1') -Channel $Channel -App $item.App -Arch $item.Arch -ProjectPath $vdproj | Out-Null
    $items += [pscustomobject]@{ App=$item.App; Arch=$item.Arch; Name=$name; ProjectDir=$projectDir; Vdproj=$vdproj }
}

$out = Join-Path $packagingRoot ("Output\{0}\{1}" -f $Channel.ToLowerInvariant(), $ver.MsiVersion)
Assert-InPackagingRoot -PathToCheck $out
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Recurse -Force }
New-Item -ItemType Directory -Force -Path $out | Out-Null

foreach ($item in $items) {
    $update = & (Join-Path $PSScriptRoot 'Update-VdprojVersion.ps1') -ProjectPath $item.Vdproj -MsiVersion $ver.MsiVersion
    if ($update.ProductCodeChanged) {
        Write-Host "Updated $($item.Name): $($update.PreviousVersion) -> $($update.ProductVersion); generated new ProductCode."
    } else {
        Write-Host "$($item.Name): ProductVersion already $($update.ProductVersion); ProductCode retained."
    }
    # Revalidate after version/code mutation, not only before it.
    & (Join-Path $PSScriptRoot 'Validate-Vdproj.ps1') -Channel $Channel -App $item.App -Arch $item.Arch -ProjectPath $item.Vdproj | Out-Null

    # Build the setup project directly. A .vdproj has Release/Debug configurations but no
    # CPU platform configuration of its own. Building a synthetic setup-only solution with
    # an invented Any CPU/x64 solution platform caused devenv to reject the configuration.
    # devenv accepts a project file as its first argument and creates/uses an implicit solution.

    # .vdproj OutputFilename is relative, and Visual Studio may resolve it relative to
    # the solution/build context rather than the .vdproj directory. Do not guess the
    # physical output directory. Remove any stale copy of the exact declared MSI name
    # anywhere under Red Ink Installer, then locate the fresh file by exact name after
    # a successful build.
    $declaredMsiName = "{0}-{1}-{2}.msi" -f $Channel, $item.App, $item.Arch
    $staleMsiFiles = @(Get-ChildItem -LiteralPath $packagingRoot -Recurse -File -Filter $declaredMsiName -ErrorAction SilentlyContinue)
    foreach ($staleMsi in $staleMsiFiles) {
        Assert-InPackagingRoot -PathToCheck $staleMsi.FullName
        Remove-Item -LiteralPath $staleMsi.FullName -Force
    }
    # setup.exe has the same generic name for every setup project and Visual Studio may place it
    # outside the .vdproj directory. Clear stale generated copies everywhere under the installer
    # tree EXCEPT the release Output tree, where already-collected bootstrapper files use different
    # names and must remain intact.
    $outputRoot = Join-Path $packagingRoot 'Output'
    $staleSetupFiles = @(Get-ChildItem -LiteralPath $packagingRoot -Recurse -File -Filter 'setup.exe' -ErrorAction SilentlyContinue)
    foreach ($staleSetup in $staleSetupFiles) {
        if ($staleSetup.FullName.StartsWith($outputRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        Assert-InPackagingRoot -PathToCheck $staleSetup.FullName
        Remove-Item -LiteralPath $staleSetup.FullName -Force
    }

    $logsDir = Join-Path $packagingRoot 'Build\InstallerLogs'
    Assert-InPackagingRoot -PathToCheck $logsDir
    New-Item -ItemType Directory -Force -Path $logsDir | Out-Null
    $devenvLog = Join-Path $logsDir ("{0}-{1}-{2}.log" -f $Channel, $item.App, $item.Arch)
    if (Test-Path -LiteralPath $devenvLog) { Remove-Item -LiteralPath $devenvLog -Force }

    Write-Host "Rebuilding setup project directly: $($item.Vdproj)"
    & $vs.Devenv $item.Vdproj '/Rebuild' 'Release' '/Out' $devenvLog
    $devenvExitCode = $LASTEXITCODE
    if ($devenvExitCode -ne 0) {
        throw "Installer build failed: $($item.Vdproj) (exit code $devenvExitCode). Visual Studio log: $devenvLog"
    }

    # A direct setup-project build must not rebuild any source project. If it does, a
    # Primary Output/project reference has been reintroduced and packaging is no longer isolated.
    if (Test-Path -LiteralPath $devenvLog) {
        $unexpectedSourceBuild = @(Select-String -LiteralPath $devenvLog -Pattern 'Build started: Project: SharedLibrary|Build started: Project: Red Ink for Word|Build started: Project: Red Ink for Excel|Build started: Project: Red Ink for Outlook' -ErrorAction SilentlyContinue)
        if ($unexpectedSourceBuild.Count -gt 0) {
            throw "Direct setup-project build unexpectedly rebuilt a source project for $Channel-$($item.App)-$($item.Arch). Visual Studio log: $devenvLog"
        }
    }

    # This package intentionally contains no Primary Output or detected-dependency objects.
    # Any dependency-analysis warning therefore indicates that Visual Studio mutated or
    # misread the setup project; fail immediately with the exact log text.
    if (Test-Path -LiteralPath $devenvLog) {
        $dependencyWarnings = @(Select-String -LiteralPath $devenvLog -Pattern "WARNING: Unable to update the dependencies of the project|The dependencies for the object '([^']+)' cannot be determined" -ErrorAction SilentlyContinue)
        if ($dependencyWarnings.Count -gt 0) {
            $warningLines = New-Object 'System.Collections.Generic.List[string]'
            foreach ($dependencyWarning in $dependencyWarnings) {
                [string]$warningLine = $dependencyWarning.Line.Trim()
                if (-not $warningLines.Contains($warningLine)) { [void]$warningLines.Add($warningLine) }
            }
            $warningText = [string]::Join("`n", $warningLines.ToArray())
            throw "Visual Studio did not complete VDPROJ dependency analysis for $Channel-$($item.App)-$($item.Arch):`n$warningText`nVisual Studio log: $devenvLog"
        }
    }

    $freshMsiFiles = @(Get-ChildItem -LiteralPath $packagingRoot -Recurse -File -Filter $declaredMsiName -ErrorAction SilentlyContinue)
    if ($freshMsiFiles.Count -eq 0) {
        throw "Installer build succeeded, but Visual Studio did not produce '$declaredMsiName' anywhere under '$packagingRoot'. Visual Studio log: $devenvLog"
    }
    if ($freshMsiFiles.Count -gt 1) {
        $pathLines = New-Object 'System.Collections.Generic.List[string]'
        foreach ($duplicateMsi in $freshMsiFiles) { [void]$pathLines.Add($duplicateMsi.FullName) }
        $paths = [string]::Join("`n  ", $pathLines.ToArray())
        throw "Installer build produced more than one '$declaredMsiName'. Refusing to guess which file is correct:`n  $paths`nVisual Studio log: $devenvLog"
    }
    $msiPath = $freshMsiFiles[0].FullName
    Write-Host "Installer output found: $msiPath"

    $channelPart = if ($Channel -eq 'Preview') { '-Preview' } else { '' }
    $targetName = "RedInk-$($item.App)$channelPart-$($ver.MsiVersion)-$($item.Arch).msi"
    $finalMsiPath = Join-Path $out $targetName
    Copy-Item -LiteralPath $msiPath -Destination $finalMsiPath -Force

    # Sign the final release MSI, then verify it before any release hashes are generated.
    # Keep LawDigital Ltd. as MSI Manufacturer; the Authenticode publisher is the
    # subject of the installed VISCHER AG code-signing certificate.
    & (Join-Path $PSScriptRoot 'Sign-MsiFiles.ps1') -Path $finalMsiPath | Out-Null

    $setupExePath = Join-Path (Split-Path -Parent $msiPath) 'setup.exe'
    if (Test-Path -LiteralPath $setupExePath) {
        $setupName = "RedInk-$($item.App)$channelPart-$($ver.MsiVersion)-$($item.Arch)-Setup.exe"
        Copy-Item -LiteralPath $setupExePath -Destination (Join-Path $out $setupName) -Force
    }
}

Copy-Item -LiteralPath (Join-Path $packagingRoot 'CustomerDocs\CUSTOMER-DEPLOYMENT-GUIDE.md') -Destination $out -Force
Copy-Item -LiteralPath (Join-Path $packagingRoot 'CustomerDocs\CUSTOMER-PREREQUISITES.md') -Destination $out -Force
Copy-Item -LiteralPath (Join-Path $packagingRoot 'CustomerDocs\CUSTOMER-SILENT-COMMANDS.md') -Destination $out -Force
Copy-Item -LiteralPath (Join-Path $packagingRoot 'Scripts\Detect-RedInk-Prerequisites.ps1') -Destination $out -Force

$hashFile = Join-Path $out 'SHA256SUMS.txt'
$hashLines = New-Object 'System.Collections.Generic.List[string]'
$releaseFiles = @(Get-ChildItem -LiteralPath $out -File | Where-Object { $_.Extension -in '.msi','.exe' } | Sort-Object Name)
foreach ($releaseFile in $releaseFiles) {
    $fileHash = Get-FileHash -LiteralPath $releaseFile.FullName -Algorithm SHA256
    [void]$hashLines.Add(("{0}  {1}" -f $fileHash.Hash, $releaseFile.Name))
}
[System.IO.File]::WriteAllLines($hashFile, $hashLines.ToArray(), [System.Text.Encoding]::ASCII)

Write-Host ''
Write-Host "Release created: $out"
Write-Host "ApplicationVersion $($ver.ApplicationVersion) -> MSI $($ver.MsiVersion)"
