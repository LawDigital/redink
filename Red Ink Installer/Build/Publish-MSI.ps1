param(
    [Parameter(Mandatory=$true)][ValidateSet('Preview','GA')][string]$Channel,
    [switch]$Build,
    [string]$PublishRoot = 'D:\Clickonce\redink'
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$installerRoot = Split-Path -Parent $PSScriptRoot
$solutionRoot = Split-Path -Parent $installerRoot
$channelSlug = $Channel.ToLowerInvariant()
$expectedBranch = if ($Channel -eq 'Preview') { 'preview' } else { 'main' }
$destination = Join-Path $PublishRoot "apps\$channelSlug"
$backupDirectory = Join-Path $PublishRoot "msi-backup\$channelSlug"
$workingDirectory = Join-Path $PublishRoot ('.msi-publish-' + $channelSlug + '-' + [guid]::NewGuid().ToString('N'))
$apps = @('Excel','Word','Outlook')
$published = $false
$createdWork = $false
$publishLock = $null

function Assert-Branch {
    $git = Get-Command git.exe -ErrorAction SilentlyContinue
    if ($null -eq $git) { throw 'git.exe is required for release branch validation.' }
    $branch = (& $git.Source -C $solutionRoot branch --show-current 2>&1)
    if ($LASTEXITCODE -ne 0 -or @($branch).Count -ne 1 -or ([string]$branch).Trim() -cne $expectedBranch) {
        throw "Expected Git branch '$expectedBranch' in '$solutionRoot'; found '$branch'."
    }
}
function Assert-Hash {
    param([string]$Path,[string]$Expected)
    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actual -cne $Expected.ToUpperInvariant()) { throw "SHA256 mismatch: $Path" }
}
function Get-MsiVersion {
    param([string]$MsiPath)
    # Windows Installer holds the MSI open while its COM database/view RCWs exist.
    # Dispose them explicitly in reverse order, including on query failures.
    $wi = $null
    $database = $null
    $view = $null
    $record = $null
    try {
        $wi = New-Object -ComObject WindowsInstaller.Installer
        $database = $wi.GetType().InvokeMember('OpenDatabase','InvokeMethod',$null,$wi,@($MsiPath,0))
        $view = $database.GetType().InvokeMember('OpenView','InvokeMethod',$null,$database,@("SELECT ``Value`` FROM ``Property`` WHERE ``Property``='ProductVersion'"))
        $view.GetType().InvokeMember('Execute','InvokeMethod',$null,$view,$null) | Out-Null
        $record = $view.GetType().InvokeMember('Fetch','InvokeMethod',$null,$view,$null)
        if ($null -eq $record) { throw "MSI ProductVersion missing: $MsiPath" }
        $version = [string]$record.GetType().InvokeMember('StringData','GetProperty',$null,$record,@(1))
        if ($version -notmatch '^\d+\.\d+\.\d+$') { throw "Invalid MSI ProductVersion '$version': $MsiPath" }
        return $version
    } finally {
        if ($null -ne $view) {
            try { $view.GetType().InvokeMember('Close','InvokeMethod',$null,$view,$null) | Out-Null }
            catch { Write-Warning "Unable to close Windows Installer MSI view for '$MsiPath': $_" }
        }
        foreach ($comObject in @($record,$view,$database,$wi)) {
            if ($null -ne $comObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
            }
        }
    }
}
function Remove-TemporaryMsi {
    param([string]$Path)
    # A scanner can briefly open the file after Windows Installer releases it.
    # Never silently ignore a failure to remove an extracted MSI.
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        if (-not (Test-Path -LiteralPath $Path)) { return }
        try {
            Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
            return
        } catch {
            if ($attempt -eq 5) { throw "Unable to remove extracted MSI '$Path' after 5 attempts: $_" }
            Start-Sleep -Milliseconds (200 * $attempt)
        }
    }
}
function Get-ZipMsiVersion {
    param([string]$Archive,[string]$ExpectedApp)
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Archive)
    try {
        $entries = @($zip.Entries | Where-Object { $_.Name -match '(?i)\.msi$' })
        if ($entries.Count -ne 1) { throw "Expected exactly one MSI in previous ZIP: $Archive" }
        $entry = $entries[0]
        if ($entry.Name -notmatch ("(?i)^RedInk-" + $ExpectedApp + "(?:-Preview)?-\d+\.\d+\.\d+-x64\.msi$")) {
            throw "Unexpected MSI in previous $ExpectedApp ZIP: $($entry.Name)"
        }
        $temp = Join-Path $workingDirectory ([guid]::NewGuid().ToString('N') + '.msi')
        try {
            $stream = $entry.Open()
            $file = [System.IO.File]::Create($temp)
            try { $stream.CopyTo($file) } finally { $file.Dispose(); $stream.Dispose() }
            return Get-MsiVersion -MsiPath $temp
        } finally { Remove-TemporaryMsi -Path $temp }
    } finally { $zip.Dispose() }
}
function Assert-MsiSignature {
    param([string]$Path,[string]$SignTool)
    & $SignTool verify /pa /v $Path
    if ($LASTEXITCODE -ne 0) { throw "Authenticode verification failed: $Path" }
}
function Find-SignTool {
    $tool = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($null -ne $tool) { return $tool.Source }
    $kits = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'
    if (Test-Path -LiteralPath $kits) {
        foreach ($version in @(Get-ChildItem -LiteralPath $kits -Directory | Sort-Object Name -Descending)) {
            foreach ($arch in @('x64','x86')) {
                $candidate = Join-Path $version.FullName "$arch\signtool.exe"
                if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
            }
        }
    }
    throw 'SignTool not found; cannot verify MSI signatures.'
}
function Write-Ascii {
    param([string]$Path,[string[]]$Lines)
    [System.IO.File]::WriteAllLines($Path,$Lines,[System.Text.Encoding]::ASCII)
}
function Get-ZipHash {
    param([string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}
function Verify-PublishedPair {
    param([string]$Zip,[string]$Sidecar)
    if (-not (Test-Path -LiteralPath $Zip -PathType Leaf) -or -not (Test-Path -LiteralPath $Sidecar -PathType Leaf)) {
        throw "Incomplete published pair: $Zip"
    }
    $line = ([System.IO.File]::ReadAllText($Sidecar)).Trim()
    $escaped = [regex]::Escape([System.IO.Path]::GetFileName($Zip))
    if ($line -notmatch ('^(?<hash>[A-Fa-f0-9]{64})\s+\*?' + $escaped + '$')) { throw "Invalid SHA256 sidecar: $Sidecar" }
    Assert-Hash -Path $Zip -Expected $Matches['hash']
}

Assert-Branch
& (Join-Path $PSScriptRoot 'Test-PowerShellSyntax.ps1') | Out-Null
if ($Build) {
    & (Join-Path $PSScriptRoot 'Build-Channel.ps1') -Channel $Channel
    if (-not $?) { throw 'Build-Channel failed.' }
}
$version = & (Join-Path $PSScriptRoot 'Get-ChannelVersion.ps1') -Channel $Channel -SolutionRoot $solutionRoot
$releaseDir = Join-Path $installerRoot "Output\$channelSlug\$($version.MsiVersion)"
if (-not (Test-Path -LiteralPath $releaseDir -PathType Container)) { throw "Release output missing: $releaseDir" }
$signTool = Find-SignTool
$hashManifest = Join-Path $releaseDir 'SHA256SUMS.txt'
if (-not (Test-Path -LiteralPath $hashManifest -PathType Leaf)) { throw "Missing release hashes: $hashManifest" }
$manifest = @{}
foreach ($line in @(Get-Content -LiteralPath $hashManifest)) {
    if ($line -notmatch '^([0-9A-Fa-f]{64})  (.+)$') { throw "Invalid release hash line: $line" }
    if ($manifest.ContainsKey($Matches[2])) { throw "Duplicate release hash: $($Matches[2])" }
    $manifest.Add($Matches[2],$Matches[1])
}
# All six architectures must exist, be signed, and match the source manifest.
foreach ($app in $apps) {
    foreach ($arch in @('x86','x64')) {
        $suffix = if ($Channel -eq 'Preview') { '-Preview' } else { '' }
        $msiName = "RedInk-$app$suffix-$($version.MsiVersion)-$arch.msi"
        $msiFile = Join-Path $releaseDir $msiName
        if (-not (Test-Path -LiteralPath $msiFile -PathType Leaf) -or -not $manifest.ContainsKey($msiName)) { throw "Missing MSI or SHA256 entry: $msiName" }
        Assert-Hash -Path $msiFile -Expected $manifest[$msiName]
        Assert-MsiSignature -Path $msiFile -SignTool $signTool
        if ((Get-MsiVersion -MsiPath $msiFile) -cne $version.MsiVersion) { throw "MSI ProductVersion mismatch: $msiFile" }
    }
}
# Every reference to a release file must be present; do not publish incomplete documentation.
$documents = @(
    'CUSTOMER-DEPLOYMENT-GUIDE.md','CUSTOMER-PREREQUISITES.md','CUSTOMER-SILENT-COMMANDS.md',
    'Detect-RedInk-Prerequisites.ps1'
)
foreach ($doc in $documents) { if (-not (Test-Path -LiteralPath (Join-Path $releaseDir $doc) -PathType Leaf)) { throw "Release documentation missing: $doc" } }
$licenseFile = Join-Path $solutionRoot 'LICENSE.txt'
if (-not (Test-Path -LiteralPath $licenseFile -PathType Leaf)) { $licenseFile = Join-Path $installerRoot 'LICENSE.txt' }
if (-not (Test-Path -LiteralPath $licenseFile -PathType Leaf)) { throw 'LICENSE.txt missing at solution root or installer root; publication aborted rather than inventing license text.' }

# Serialize publishers sharing the same channel; keep the lock outside temporary work.
New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
$lockPath = Join-Path $backupDirectory '.publish.lock'
try {
    $publishLock = New-Object System.IO.FileStream($lockPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
} catch { throw "Another publisher may be active for $Channel (lock: $lockPath): $_" }
New-Item -ItemType Directory -Path $workingDirectory -Force | Out-Null
$createdWork = $true
$jobs = @()
try {
    foreach ($app in $apps) {
        $appSlug = $app.ToLowerInvariant()
        $suffix = if ($Channel -eq 'Preview') { '-preview' } else { '' }
        $name = "redink-$appSlug$suffix-msi.zip"
        $base = Join-Path $workingDirectory $appSlug
        New-Item -ItemType Directory -Path $base -Force | Out-Null
        $msiName = "RedInk-$app$(if ($Channel -eq 'Preview') { '-Preview' } else { '' })-$($version.MsiVersion)-x64.msi"
        Copy-Item -LiteralPath (Join-Path $releaseDir $msiName) -Destination $base
        foreach ($doc in $documents) { Copy-Item -LiteralPath (Join-Path $releaseDir $doc) -Destination $base }
        Copy-Item -LiteralPath $licenseFile -Destination $base
        Write-Ascii -Path (Join-Path $base 'VERSION.txt') -Lines @($version.MsiVersion)
        Write-Ascii -Path (Join-Path $base 'RELEASE-INFO.txt') -Lines @("Application: $app", "Channel: $Channel", "MSI ProductVersion: $($version.MsiVersion)", 'Office architecture: x64')
        Write-Ascii -Path (Join-Path $base 'INSTALL.txt') -Lines @('Install the included MSI with msiexec. Read CUSTOMER-DEPLOYMENT-GUIDE.md and CUSTOMER-PREREQUISITES.md first.',"msiexec.exe /i `"$msiName`" /qn /norestart")
        $hashes = @()
        foreach ($file in @(Get-ChildItem -LiteralPath $base -File | Sort-Object Name)) { $hashes += "$(Get-ZipHash -Path $file.FullName)  $($file.Name)" }
        Write-Ascii -Path (Join-Path $base 'SHA256SUMS.txt') -Lines $hashes
        $zip = Join-Path $workingDirectory $name
        [System.IO.Compression.ZipFile]::CreateFromDirectory($base,$zip,[System.IO.Compression.CompressionLevel]::Optimal,$false)
        $sidecar = "$zip.sha256"
        Write-Ascii -Path $sidecar -Lines @("$(Get-ZipHash -Path $zip)  $name")
        Verify-PublishedPair -Zip $zip -Sidecar $sidecar
        # Verify all ZIP members against its internal manifest, not only the outer sidecar.
        $archive = [System.IO.Compression.ZipFile]::OpenRead($zip)
        try {
            $members = @($archive.Entries | Where-Object { $_.Name -ne 'SHA256SUMS.txt' })
            if ($members.Count -ne $hashes.Count) { throw "ZIP entry count mismatch: $zip" }
            foreach ($entry in $members) {
                if ($entry.FullName -cne $entry.Name) { throw "Unexpected nested ZIP entry: $($entry.FullName)" }
                $expected = @($hashes | Where-Object { $_ -like "*  $($entry.Name)" })
                if ($expected.Count -ne 1) { throw "Missing/duplicate internal ZIP hash: $($entry.Name)" }
                $stream = $entry.Open()
                $sha = [System.Security.Cryptography.SHA256]::Create()
                try { $actual = [System.BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','') } finally { $sha.Dispose(); $stream.Dispose() }
                if ($actual -cne $expected[0].Substring(0,64)) { throw "ZIP member hash mismatch: $($entry.Name)" }
            }
        } finally { $archive.Dispose() }
        $jobs += [pscustomobject]@{ App=$app; Name=$name; Zip=$zip; Sidecar=$sidecar; LiveZip=(Join-Path $destination $name); LiveSidecar=(Join-Path $destination "$name.sha256"); OldZip=$false; OldSidecar=$false; Archive=$null }
    }

    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
    # Capture every prior file before changing any live URL. Unusable old ZIPs fail closed.
    foreach ($job in $jobs) {
        $job.OldZip = Test-Path -LiteralPath $job.LiveZip -PathType Leaf
        $job.OldSidecar = Test-Path -LiteralPath $job.LiveSidecar -PathType Leaf
        if ($job.OldSidecar -and -not $job.OldZip) { throw "Orphan checksum sidecar: $($job.LiveSidecar)" }
        if ($job.OldZip) {
            $oldVersion = Get-ZipMsiVersion -Archive $job.LiveZip -ExpectedApp $job.App
            $oldHash = Get-ZipHash -Path $job.LiveZip
            $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
            $uuid = [guid]::NewGuid().ToString('N').Substring(0,12)
            $archiveName = "redink-$($job.App.ToLowerInvariant())$(if ($Channel -eq 'Preview') { '-preview' } else { '' })-msi-$oldVersion-$stamp-$uuid.zip"
            $archivePath = Join-Path $backupDirectory $archiveName
            Copy-Item -LiteralPath $job.LiveZip -Destination $archivePath -ErrorAction Stop
            Assert-Hash -Path $archivePath -Expected $oldHash
            $job.Archive = $archivePath
        }
        if ($job.OldZip -and $job.OldSidecar) { Verify-PublishedPair -Zip $job.LiveZip -Sidecar $job.LiveSidecar }
        if ($job.OldSidecar) {
            Copy-Item -LiteralPath $job.LiveSidecar -Destination (Join-Path $workingDirectory ($job.Name + '.old.sha256')) -ErrorAction Stop
        }
    }
    # A destination-local staging directory permits rename on the same volume.
    $stage = Join-Path $destination ('.stage-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    foreach ($job in $jobs) {
        Copy-Item -LiteralPath $job.Zip -Destination (Join-Path $stage $job.Name)
        Copy-Item -LiteralPath $job.Sidecar -Destination (Join-Path $stage ($job.Name + '.sha256'))
        Verify-PublishedPair -Zip (Join-Path $stage $job.Name) -Sidecar (Join-Path $stage ($job.Name + '.sha256'))
    }
    try {
        foreach ($job in $jobs) {
            Move-Item -LiteralPath (Join-Path $stage $job.Name) -Destination $job.LiveZip -Force
            Move-Item -LiteralPath (Join-Path $stage ($job.Name + '.sha256')) -Destination $job.LiveSidecar -Force
            Verify-PublishedPair -Zip $job.LiveZip -Sidecar $job.LiveSidecar
        }
        $published = $true
    } catch {
        $cause = $_
        $rollbackErrors = @()
        foreach ($job in $jobs) {
            try {
                if ($job.OldZip) {
                    if (-not $job.Archive) { throw 'Old ZIP lacks backup.' }
                    Copy-Item -LiteralPath $job.Archive -Destination $job.LiveZip -Force
                } elseif (Test-Path -LiteralPath $job.LiveZip) { Remove-Item -LiteralPath $job.LiveZip -Force }
                if ($job.OldSidecar) {
                    Copy-Item -LiteralPath (Join-Path $workingDirectory ($job.Name + '.old.sha256')) -Destination $job.LiveSidecar -Force
                } elseif (Test-Path -LiteralPath $job.LiveSidecar) { Remove-Item -LiteralPath $job.LiveSidecar -Force }
                if ($job.OldZip -and $job.OldSidecar) { Verify-PublishedPair -Zip $job.LiveZip -Sidecar $job.LiveSidecar }
            } catch { $rollbackErrors += "$($job.Name): $_" }
        }
        if ($rollbackErrors.Count -gt 0) { throw "Publish failed: $cause. ROLLBACK INCOMPLETE: $($rollbackErrors -join '; '). Backup ZIPs: $backupDirectory" }
        throw "Publish failed, previous files restored: $cause"
    } finally {
        if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue }
    }
    Write-Host "PUBLISHED $Channel $($version.MsiVersion) to $destination; previous ZIPs archived in $backupDirectory"
} finally {
    if ($createdWork -and (Test-Path -LiteralPath $workingDirectory)) {
        Remove-Item -LiteralPath $workingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($null -ne $publishLock) { $publishLock.Dispose() }
}
