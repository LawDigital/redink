#requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('onedir', 'onefile')][string]$Mode = 'onedir',
    [string]$PythonExe = '',
    [string]$PythonVersion = '3.12',
    [string]$CertificateSubject = 'VISCHER AG',
    [string]$CertificateThumbprint = '',
    [switch]$MachineCertificateStore,
    [switch]$SkipSigning,
    [string]$PfxFile = '',
    [string]$PfxPasswordEnvironmentVariable = 'REDINK_SIGN_PFX_PASSWORD',
    [string]$SignToolPath = '',
    [string]$TimestampServer = 'http://timestamp.digicert.com'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if ($env:OS -ne 'Windows_NT') { throw 'The Windows executable must be built on Windows.' }
if ($CertificateThumbprint -and $PfxFile) { throw 'Choose a certificate-store thumbprint or a PFX file, not both.' }
if ($SkipSigning -and ($PfxFile -or $CertificateThumbprint)) { throw '-SkipSigning cannot be combined with explicit signing credentials.' }
if ($MachineCertificateStore -and -not $CertificateThumbprint) { throw '-MachineCertificateStore requires -CertificateThumbprint.' }
$ProjectRoot = $PSScriptRoot
$VenvRoot = Join-Path $ProjectRoot '.venv-build'
$BuildPython = Join-Path $VenvRoot 'Scripts\python.exe'
$BuildWork = Join-Path $ProjectRoot ('build\pyinstaller-' + $Mode)
$OutputRoot = Join-Path $ProjectRoot ('dist\windows-' + $Mode)
$ReleaseRoot = Join-Path $OutputRoot 'RedInkUsageServer'
$Version = '1.4.2'

function Invoke-Checked {
    param([Parameter(Mandatory=$true)][string]$Command,
          [Parameter(Mandatory=$true)][AllowEmptyCollection()][string[]]$Arguments)
    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Command failed with exit code $LASTEXITCODE`: $Command" }
}

function Find-SignTool {
    if ($SignToolPath) {
        if (-not (Test-Path -LiteralPath $SignToolPath -PathType Leaf)) { throw 'SignToolPath does not point to a file.' }
        return (Resolve-Path -LiteralPath $SignToolPath).Path
    }
    $onPath = Get-Command 'signtool.exe' -ErrorAction SilentlyContinue
    if ($null -ne $onPath) { return $onPath.Source }
    $kitsRoot = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'
    if (Test-Path -LiteralPath $kitsRoot) {
        $tools = @(Get-ChildItem -LiteralPath $kitsRoot -Directory |
            Sort-Object Name -Descending |
            ForEach-Object { Join-Path $_.FullName 'x64\signtool.exe' } |
            Where-Object { Test-Path -LiteralPath $_ -PathType Leaf })
        if ($tools.Count -gt 0) { return $tools[0] }
    }
    throw 'SignTool was not found. Install the Windows SDK or pass -SignToolPath.'
}

function Select-CodeSigningCertificate {
    param([Parameter(Mandatory=$true)][string]$Owner)
    # Resolve from the Windows stores, never rely on signtool /a's ambiguous choice.
    # Match an exact CN= or O= distinguished-name attribute, not arbitrary substrings.
    $escapedOwner = [Regex]::Escape($Owner.Trim())
    if ([String]::IsNullOrWhiteSpace($Owner)) { throw 'CertificateSubject must not be empty.' }
    $subjectPattern = '(?:^|,)\s*(?:CN|O)\s*=\s*' + $escapedOwner + '\s*(?:,|$)'
    $now = [DateTime]::UtcNow
    $candidates = @()
    foreach ($storeName in @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My')) {
        foreach ($cert in @(Get-ChildItem -Path $storeName -ErrorAction Stop)) {
            if (-not $cert.HasPrivateKey -or $cert.NotBefore.ToUniversalTime() -gt $now -or $cert.NotAfter.ToUniversalTime() -le $now) { continue }
            if ($cert.Subject -notmatch $subjectPattern) { continue }
            $codeSigning = $false
            foreach ($usage in @($cert.EnhancedKeyUsageList)) {
                if ($usage.ObjectId -eq '1.3.6.1.5.5.7.3.3') { $codeSigning = $true; break }
            }
            if ($codeSigning) {
                $candidates += [PSCustomObject]@{ Certificate = $cert; Store = $storeName; Thumbprint = $cert.Thumbprint }
            }
        }
    }
    # Same certificate in CurrentUser and LocalMachine is not an ambiguity.
    $unique = @($candidates | Group-Object Thumbprint)
    if ($unique.Count -eq 0) { throw "No currently valid code-signing certificate with a private key and CN/O '$Owner' was found in CurrentUser/My or LocalMachine/My. The EXE was not built or released unsigned. Use -SkipSigning only for a deliberate test build." }
    if ($unique.Count -gt 1) {
        $details = ($unique | ForEach-Object { $_.Name + ' (' + $_.Group[0].Certificate.NotAfter.ToString('yyyy-MM-dd') + ')' }) -join ', '
        throw "Multiple valid code-signing certificates match '$Owner': $details. Remove the ambiguity or supply -CertificateThumbprint explicitly."
    }
    $selection = @($unique[0].Group | Where-Object { $_.Store -eq 'Cert:\CurrentUser\My' })
    if ($selection.Count -gt 0) { return $selection[0] }
    return $unique[0].Group[0]
}

function Assert-NoDeploymentData {
    param([Parameter(Mandatory=$true)][string]$Directory)
    if (-not (Test-Path -LiteralPath $Directory)) { return }
    # The convenience launcher can initialize data next to the executable.
    # Refuse to clean output containing any deployment state, even if the
    # operator started that deployment directly inside dist by mistake.
    # Enumeration failures also stop the build rather than bypass this guard.
    $state = Get-ChildItem -LiteralPath $Directory -Recurse -Force -ErrorAction Stop |
        Where-Object {
            ($_.PSIsContainer -and $_.Name -ieq 'data') -or
            (-not $_.PSIsContainer -and (
                $_.Name -in @('secrets.json', 'server.toml', 'server.lock') -or
                $_.Name -like '*.sqlite3' -or
                $_.Name -like '*.sqlite3-wal' -or
                $_.Name -like '*.sqlite3-shm'
            ))
        } | Select-Object -First 1
    if ($null -ne $state) {
        throw "Build output contains deployment data: $($state.FullName). Nothing was deleted. Stop the server, back up and move the deployment/data outside dist, then retry."
    }
}

# Select and verify signing capability BEFORE the potentially lengthy build.
$AutomaticCertificate = $null
if (-not $SkipSigning -and -not $PfxFile -and -not $CertificateThumbprint) {
    $AutomaticCertificate = Select-CodeSigningCertificate -Owner $CertificateSubject
    Write-Host "Signing identity: $($AutomaticCertificate.Certificate.Subject) [$($AutomaticCertificate.Store)]"
}
if (-not $SkipSigning) { $null = Find-SignTool }

$previousMode = [Environment]::GetEnvironmentVariable('REDINK_BUILD_MODE', 'Process')
Push-Location -LiteralPath $ProjectRoot
try {
    if (-not (Test-Path -LiteralPath $BuildPython)) {
        Write-Host 'Creating an isolated build environment...'
        if ($PythonExe) { Invoke-Checked $PythonExe @('-m', 'venv', $VenvRoot) }
        elseif (Get-Command 'py.exe' -ErrorAction SilentlyContinue) { Invoke-Checked 'py.exe' @("-$PythonVersion", '-m', 'venv', $VenvRoot) }
        else { Invoke-Checked 'python.exe' @('-m', 'venv', $VenvRoot) }
    }
    Invoke-Checked $BuildPython @('scripts/build_support.py', 'check-python')
    Invoke-Checked $BuildPython @('-m', 'pip', 'install', '--disable-pip-version-check', '-r', 'requirements-build.txt')
    Invoke-Checked $BuildPython @('-m', 'pip', 'check')
    Write-Host 'Running regression tests...'
    Invoke-Checked $BuildPython @('-m', 'unittest', 'discover', '-s', 'tests', '-v')
    Invoke-Checked $BuildPython @('-m', 'compileall', '-q', 'redink_server', 'run_server.py')

    # Check immediately before cleanup; never delete an accidentally initialized deployment.
    Assert-NoDeploymentData -Directory $OutputRoot
    if (Test-Path -LiteralPath $OutputRoot) { Remove-Item -LiteralPath $OutputRoot -Recurse -Force }
    New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
    $env:REDINK_BUILD_MODE = $Mode
    Invoke-Checked $BuildPython @('-m', 'PyInstaller', '--noconfirm', '--clean', '--workpath', $BuildWork, '--distpath', $OutputRoot, 'RedInkUsageServer.spec')
    if ($Mode -eq 'onefile') {
        New-Item -ItemType Directory -Path $ReleaseRoot -Force | Out-Null
        Move-Item -LiteralPath (Join-Path $OutputRoot 'RedInkUsageServer.exe') -Destination $ReleaseRoot
    }
    $Executable = Join-Path $ReleaseRoot 'RedInkUsageServer.exe'
    if (-not (Test-Path -LiteralPath $Executable -PathType Leaf)) { throw 'PyInstaller did not produce the expected executable.' }

    Write-Host 'Checking the frozen CLI against isolated temporary data...'
    $SmokeData = Join-Path ([IO.Path]::GetTempPath()) ('RedInkUsageServer-build-' + [Guid]::NewGuid().ToString('N'))
    try {
        Invoke-Checked $Executable @('--help')
        # Suppress one-time credentials for this disposable test deployment.
        & $Executable init --data-dir $SmokeData *> $null
        if ($LASTEXITCODE -ne 0) { throw 'The frozen initialization smoke test failed.' }
        Invoke-Checked $Executable @('check', '--data-dir', $SmokeData)
    } finally {
        if (Test-Path -LiteralPath $SmokeData) { Remove-Item -LiteralPath $SmokeData -Recurse -Force }
    }

    $SigningStatus = 'unsigned'
    if (-not $SkipSigning) {
        $Tool = Find-SignTool
        $SigningArguments = @('sign', '/fd', 'SHA256', '/tr', $TimestampServer, '/td', 'SHA256')
        if ($AutomaticCertificate -or $CertificateThumbprint) {
            if ($AutomaticCertificate) {
                $CertificateThumbprint = $AutomaticCertificate.Thumbprint
                $MachineCertificateStore = $AutomaticCertificate.Store -eq 'Cert:\LocalMachine\My'
            }
            $thumbprint = ($CertificateThumbprint -replace '\s', '').ToUpperInvariant()
            if ($thumbprint -notmatch '^[A-F0-9]{40}$') { throw 'CertificateThumbprint must be a 40-digit certificate thumbprint.' }
            $SigningArguments += @('/sha1', $thumbprint, '/s', 'My')
            if ($MachineCertificateStore) { $SigningArguments += '/sm' }
        } else {
            if (-not (Test-Path -LiteralPath $PfxFile -PathType Leaf)) { throw 'PfxFile was not found.' }
            $pfxPassword = [Environment]::GetEnvironmentVariable($PfxPasswordEnvironmentVariable)
            if ([String]::IsNullOrEmpty($pfxPassword)) { throw "Set the $PfxPasswordEnvironmentVariable environment variable for PFX signing." }
            $SigningArguments += @('/f', (Resolve-Path -LiteralPath $PfxFile).Path, '/p', $pfxPassword)
        }
        $SigningArguments += $Executable
        try { Invoke-Checked $Tool $SigningArguments }
        finally { $pfxPassword = $null; $SigningArguments = $null }
        Invoke-Checked $Tool @('verify', '/pa', '/all', '/v', $Executable)
        $signature = Get-AuthenticodeSignature -LiteralPath $Executable
        if ($signature.Status -ne 'Valid') { throw "Authenticode validation failed: $($signature.Status)" }
        $SigningStatus = 'signed-and-verified'
    }

    # Publish only approved runtime entry points and administrator/user documentation.
    # Never recursively include docs, examples, config, scripts or source trees.
    Invoke-Checked $BuildPython @('scripts/release_payload.py', 'stage', $ProjectRoot, $ReleaseRoot)
    $LicenseDir = Join-Path $ReleaseRoot 'third-party-licenses'
    Invoke-Checked $BuildPython @('scripts/collect-third-party-notices.py', $LicenseDir)
    # Third-party license texts are legally required; they are not build telemetry.
    # A failed release audit stops packaging before any ZIP is produced.
    Invoke-Checked $BuildPython @('scripts/release_payload.py', 'verify', $ProjectRoot, $ReleaseRoot)
    $ReleaseZip = Join-Path $OutputRoot ("RedInkUsageServer-$Version-windows-$Mode-$SigningStatus.zip")
    # Python's ZIP writer includes dotfiles and does not inherit Compress-Archive's 2 GB file limit.
    Invoke-Checked $BuildPython @('scripts/build_support.py', 'archive', $ReleaseRoot, $ReleaseZip)
    $ZipHash = (Get-FileHash -LiteralPath $ReleaseZip -Algorithm SHA256).Hash.ToLowerInvariant()
    ($ZipHash + '  ' + [IO.Path]::GetFileName($ReleaseZip)) | Set-Content -LiteralPath ($ReleaseZip + '.sha256') -Encoding ASCII
    Write-Host "Build complete ($SigningStatus): $ReleaseRoot"
    Write-Host "Distribution archive: $ReleaseZip"
} finally {
    [Environment]::SetEnvironmentVariable('REDINK_BUILD_MODE', $previousMode, 'Process')
    Pop-Location
}
