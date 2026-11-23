# Part of "Red Ink" (Ask Inky). Copyright (c) LawDigital AG, Switzerland.
# P18 W367: Authenticode signing with the VISCHER AG or LawDigital AG code-signing certificate (equally valid,
# DARO 2026-11-22), selected automatically
# (code-signing EKU, private key present, organisation VISCHER AG, latest expiry; CurrentUser before
# LocalMachine). No thumbprint is stored anywhere. Exit code 0 = signed and verified.
param([Parameter(Mandatory = $true)][string[]]$Path, [string[]]$Organizations = @('VISCHER AG', 'LawDigital AG'))
$ErrorActionPreference = 'Stop'

function Find-SignTool {
    $cmd = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $roots = @("${env:ProgramFiles(x86)}\Windows Kits\10\bin", "$env:ProgramFiles\Windows Kits\10\bin")
    $found = foreach ($r in $roots) { if (Test-Path $r) { Get-ChildItem $r -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue | Where-Object { $_.FullName -match '\\x64\\' } } }
    $best = $found | Sort-Object { $_.Directory.Parent.Name } -Descending | Select-Object -First 1
    if (-not $best) { throw 'SignTool.exe wurde nicht gefunden (Windows SDK installieren).' }
    return $best.FullName
}

function Find-Certificate {
  foreach ($Organization in $Organizations) {
    foreach ($store in 'Cert:\CurrentUser\My', 'Cert:\LocalMachine\My') {
        $candidates = Get-ChildItem $store -ErrorAction SilentlyContinue | Where-Object {
            $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) -and
            ($_.EnhancedKeyUsageList | Where-Object { $_.ObjectId -eq '1.3.6.1.5.5.7.3.3' }) -and
            ($_.Subject -match ('(^|, )O="?' + [regex]::Escape($Organization) + '"?(,|$)'))
        } | Sort-Object NotAfter -Descending
        if ($candidates) { return @{ Cert = $candidates[0]; Store = $store } }
    }
  }
  throw ("Kein gültiges Code-Signing-Zertifikat gefunden (" + ($Organizations -join ' oder ') + ").")
}

$signtool = Find-SignTool
$selected = Find-Certificate
$machine = if ($selected.Store -like 'Cert:\LocalMachine*') { @('/sm') } else { @() }
Write-Host ("Zertifikat: {0} (gültig bis {1:dd.MM.yyyy})" -f $selected.Cert.Subject, $selected.Cert.NotAfter)
foreach ($file in $Path) {
    $signed = $false
    foreach ($ts in 'http://timestamp.digicert.com', 'http://timestamp.sectigo.com') {
        & $signtool sign @machine /sha1 $selected.Cert.Thumbprint /fd SHA256 /tr $ts /td SHA256 /d 'Ask Inky' /du 'https://redink.ai' $file
        if ($LASTEXITCODE -eq 0) { $signed = $true; break }
    }
    if (-not $signed) { throw "Signieren fehlgeschlagen: $file" }
    & $signtool verify /pa /tw $file
    if ($LASTEXITCODE -ne 0) { throw "Signatur nicht gültig: $file" }
    Write-Host "Signiert: $file"
}
exit 0
