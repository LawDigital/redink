param(
    [Parameter(Mandatory=$true)][string[]]$Path,
    [string]$CertificateSubject = 'VISCHER AG',
    [string]$CertificateThumbprint = $env:REDINK_SIGNING_CERT_THUMBPRINT,
    [string]$TimestampUrl = $env:REDINK_TIMESTAMP_URL
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

if ([string]::IsNullOrWhiteSpace($TimestampUrl)) {
    $TimestampUrl = 'http://timestamp.digicert.com'
}

function Find-SignTool {
    $command = Get-Command 'signtool.exe' -ErrorAction SilentlyContinue
    if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace([string]$command.Source)) {
        return [string]$command.Source
    }

    $kitsRoot = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'
    if (Test-Path -LiteralPath $kitsRoot -PathType Container) {
        $versionDirs = @(Get-ChildItem -LiteralPath $kitsRoot -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
        foreach ($versionDir in $versionDirs) {
            foreach ($architecture in @('x64','x86')) {
                $candidate = Join-Path $versionDir.FullName ("{0}\signtool.exe" -f $architecture)
                if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
            }
        }
    }

    throw 'signtool.exe was not found. Install a Windows SDK that includes SignTool, or put signtool.exe on PATH.'
}

function Get-CodeSigningCertificate {
    param(
        [Parameter(Mandatory=$true)][string]$SubjectText,
        [AllowEmptyString()][string]$Thumbprint
    )

    $now = Get-Date
    $matches = New-Object 'System.Collections.Generic.List[object]'
    foreach ($storeInfo in @(
        @{ Location = 'CurrentUser'; Path = 'Cert:\CurrentUser\My'; SignToolMachineSwitch = $false },
        @{ Location = 'LocalMachine'; Path = 'Cert:\LocalMachine\My'; SignToolMachineSwitch = $true }
    )) {
        $certificates = @(Get-ChildItem -LiteralPath $storeInfo.Path -ErrorAction SilentlyContinue)
        foreach ($certificate in $certificates) {
            if (-not $certificate.HasPrivateKey) { continue }
            if ($certificate.NotBefore -gt $now -or $certificate.NotAfter -lt $now) { continue }

            $normalizedCertificateThumbprint = ([string]$certificate.Thumbprint).Replace(' ', '').ToUpperInvariant()
            if (-not [string]::IsNullOrWhiteSpace($Thumbprint)) {
                $normalizedRequestedThumbprint = $Thumbprint.Replace(' ', '').ToUpperInvariant()
                if ($normalizedCertificateThumbprint -ne $normalizedRequestedThumbprint) { continue }
            } elseif ($certificate.Subject.IndexOf($SubjectText, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                continue
            }

            $hasCodeSigningEku = $false
            try {
                foreach ($eku in @($certificate.EnhancedKeyUsageList)) {
                    if ([string]$eku.ObjectId.Value -eq '1.3.6.1.5.5.7.3.3') {
                        $hasCodeSigningEku = $true
                        break
                    }
                }
            } catch {
                # SignTool will perform the authoritative usage check below.
                $hasCodeSigningEku = $true
            }
            if (-not $hasCodeSigningEku) { continue }

            [void]$matches.Add([pscustomobject]@{
                Certificate = $certificate
                StoreLocation = [string]$storeInfo.Location
                UseMachineStore = [bool]$storeInfo.SignToolMachineSwitch
            })
        }
    }

    if ($matches.Count -eq 0) {
        if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
            throw ("No currently valid code-signing certificate with a private key and subject containing '{0}' was found in CurrentUser\\My or LocalMachine\\My." -f $SubjectText)
        }
        throw ("No currently valid code-signing certificate with private key and thumbprint '{0}' was found in CurrentUser\\My or LocalMachine\\My." -f $Thumbprint)
    }
    if ($matches.Count -gt 1) {
        $descriptions = New-Object 'System.Collections.Generic.List[string]'
        foreach ($match in $matches) {
            [void]$descriptions.Add(("{0} | {1} | {2}" -f $match.StoreLocation, $match.Certificate.Thumbprint, $match.Certificate.Subject))
        }
        throw ("More than one matching code-signing certificate was found. Set REDINK_SIGNING_CERT_THUMBPRINT to the intended certificate thumbprint:`r`n{0}" -f [string]::Join("`r`n", $descriptions.ToArray()))
    }

    return $matches[0]
}

$signTool = Find-SignTool
$selected = Get-CodeSigningCertificate -SubjectText $CertificateSubject -Thumbprint $CertificateThumbprint
$certificate = $selected.Certificate
Write-Host ("Signing certificate: {0}" -f $certificate.Subject)
Write-Host ("Certificate thumbprint: {0}" -f $certificate.Thumbprint)
Write-Host ("Certificate store: {0}\\My" -f $selected.StoreLocation)
Write-Host ("Timestamp server: {0}" -f $TimestampUrl)

foreach ($inputPath in $Path) {
    $resolvedPath = (Resolve-Path -LiteralPath $inputPath -ErrorAction Stop).Path
    if ([System.IO.Path]::GetExtension($resolvedPath) -ine '.msi') {
        throw ("Refusing to sign non-MSI file: {0}" -f $resolvedPath)
    }

    $arguments = New-Object 'System.Collections.Generic.List[string]'
    [void]$arguments.Add('sign')
    [void]$arguments.Add('/v')
    [void]$arguments.Add('/fd')
    [void]$arguments.Add('SHA256')
    [void]$arguments.Add('/sha1')
    [void]$arguments.Add(([string]$certificate.Thumbprint).Replace(' ', ''))
    if ($selected.UseMachineStore) { [void]$arguments.Add('/sm') }
    [void]$arguments.Add('/tr')
    [void]$arguments.Add($TimestampUrl)
    [void]$arguments.Add('/td')
    [void]$arguments.Add('SHA256')
    [void]$arguments.Add($resolvedPath)

    Write-Host ("Signing MSI: {0}" -f $resolvedPath)
    & $signTool $arguments.ToArray()
    if ($LASTEXITCODE -ne 0) {
        throw ("SignTool failed to sign '{0}' with exit code {1}." -f $resolvedPath, $LASTEXITCODE)
    }

    & $signTool 'verify' '/pa' '/v' $resolvedPath
    if ($LASTEXITCODE -ne 0) {
        throw ("SignTool verification failed for '{0}' with exit code {1}." -f $resolvedPath, $LASTEXITCODE)
    }

    # SignTool verification is authoritative for MSI signatures.
    Write-Host ("Verified signed MSI: {0}" -f $resolvedPath)
}
