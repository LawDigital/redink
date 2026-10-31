#requires -Version 5.1
<#
Optional launcher for an encrypted TLS key and unattended Windows operation.
Create PasswordFile under the runtime account on this computer with:
  Read-Host 'PEM key password' -AsSecureString |
    ConvertFrom-SecureString | Set-Content -LiteralPath 'C:\Private\tls-password.dpapi'
The password is decrypted only for this process and never put in task arguments.
The existing install-startup-task.ps1 remains the direct-executable launcher;
see docs/ADMINISTRATOR_GUIDE.md for the required account and TLS settings.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Executable,
    [Parameter(Mandatory=$true)][string]$DataDir,
    [Parameter(Mandatory=$true)][string]$PasswordFile,
    [ValidatePattern('^[A-Za-z_][A-Za-z0-9_]{0,127}$')]
    [string]$PasswordEnvironmentVariable = 'REDINK_TLS_KEY_PASSWORD',
    [switch]$Check
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if ($env:OS -ne 'Windows_NT') { throw 'This DPAPI launcher requires Windows.' }
$Executable = (Resolve-Path -LiteralPath $Executable -ErrorAction Stop).Path
$DataDir = (Resolve-Path -LiteralPath $DataDir -ErrorAction Stop).Path
$PasswordFile = (Resolve-Path -LiteralPath $PasswordFile -ErrorAction Stop).Path
if ([IO.Path]::GetFileName($Executable) -ne 'RedInkUsageServer.exe') { throw 'Executable must point to RedInkUsageServer.exe.' }
if ((Get-Item -LiteralPath $PasswordFile).Length -gt 65536) { throw 'The protected password file is unexpectedly large.' }
$redInkPreviousPassword = [Environment]::GetEnvironmentVariable($PasswordEnvironmentVariable, 'Process')
$redInkPasswordPointer = [IntPtr]::Zero
$redInkSecurePassword = $null
$redInkExitCode = 1
try {
    try {
        $redInkSecurePassword = (Get-Content -LiteralPath $PasswordFile -Raw).Trim() | ConvertTo-SecureString -ErrorAction Stop
        $redInkPasswordPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($redInkSecurePassword)
    } catch {
        throw 'Cannot decrypt the TLS password. The file must have been created on this computer by the same Windows account that runs this task.'
    }
    [Environment]::SetEnvironmentVariable($PasswordEnvironmentVariable,
        [Runtime.InteropServices.Marshal]::PtrToStringBSTR($redInkPasswordPointer), 'Process')
    $redInkCommand = if ($Check) { 'check' } else { 'serve' }
    $arguments = @($redInkCommand, '--data-dir', $DataDir)
    if (-not $Check) { $arguments += '--no-browser' }
    & $Executable @arguments
    $redInkExitCode = $LASTEXITCODE
} finally {
    [Environment]::SetEnvironmentVariable($PasswordEnvironmentVariable, $redInkPreviousPassword, 'Process')
    if ($redInkPasswordPointer -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($redInkPasswordPointer)
    }
    if ($null -ne $redInkSecurePassword) { $redInkSecurePassword.Dispose() }
}
exit $redInkExitCode
