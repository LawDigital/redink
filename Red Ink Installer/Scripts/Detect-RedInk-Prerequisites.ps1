$ErrorActionPreference = 'SilentlyContinue'

function Get-DotNet48Status {
    $p = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction SilentlyContinue
    $release = if ($p) { [int]$p.Release } else { 0 }
    [pscustomobject]@{ Name='.NET Framework 4.8+'; Present=($release -ge 528040); Detail="Release=$release" }
}

function Get-VstoStatus {
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R',
        'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4'
    )
    foreach ($path in $paths) {
        $p = Get-ItemProperty $path -ErrorAction SilentlyContinue
        if ($p) { return [pscustomobject]@{ Name='VSTO Runtime'; Present=$true; Detail=("{0} {1}" -f $path,$p.Version) } }
    }
    [pscustomobject]@{ Name='VSTO Runtime'; Present=$false; Detail='Not detected' }
}

function Get-VcX64Status {
    $p = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64' -ErrorAction SilentlyContinue
    $ok = $p -and ($p.Installed -eq 1)
    [pscustomobject]@{ Name='Visual C++ v14 x64 Runtime'; Present=[bool]$ok; Detail=if($p){$p.Version}else{'Not detected'} }
}

function Get-OfficeBitness {
    $c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($c2r -and $c2r.Platform) { return $c2r.Platform }
    $x64 = Test-Path 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook'
    $x86 = Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Outlook'
    if ($x64 -and -not $x86) { return 'x64 (registry inference)' }
    if ($x86 -and -not $x64) { return 'x86 (registry inference)' }
    return 'Unknown - verify Office About dialog / deployment inventory'
}

Write-Host "Red Ink prerequisite report" -ForegroundColor Cyan
Write-Host "Office bitness: $(Get-OfficeBitness)"
Write-Host ""
@(
    Get-DotNet48Status
    Get-VstoStatus
    Get-VcX64Status
) | Format-Table -AutoSize

Write-Host ""
Write-Host 'Note: VC++ x64 is currently relevant to the Word package because the existing Word ClickOnce project declares that prerequisite.'
