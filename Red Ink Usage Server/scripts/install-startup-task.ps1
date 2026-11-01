#requires -Version 5.1
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [Parameter(Mandatory=$true)][string]$Executable,
    [Parameter(Mandatory=$true)][string]$DataDir,
    [Parameter(Mandatory=$true)][ValidateSet('CurrentUser', 'System', 'Credential')][string]$AccountMode,
    [System.Management.Automation.PSCredential]$Credential,
    [string]$TaskName = 'Red Ink Usage Server',
    [switch]$Replace,
    [switch]$StartNow
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Import-Module ScheduledTasks -ErrorAction Stop
$Executable = (Resolve-Path -LiteralPath $Executable -ErrorAction Stop).Path
$DataDir = (Resolve-Path -LiteralPath $DataDir -ErrorAction Stop).Path
if ([IO.Path]::GetFileName($Executable) -ne 'RedInkUsageServer.exe') { throw 'Executable must point to RedInkUsageServer.exe.' }
if ($DataDir.Contains('"')) { throw 'The data-directory path contains an invalid quote character.' }
& $Executable check --data-dir $DataDir
if ($LASTEXITCODE -ne 0) { throw 'Server configuration validation failed. No task was installed.' }
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -ne $existing -and -not $Replace) { throw "Task '$TaskName' already exists. Use -Replace only after reviewing it." }
if (-not $PSCmdlet.ShouldProcess($TaskName, "Register automatic server startup under $AccountMode")) { return }
$action = New-ScheduledTaskAction -Execute $Executable -Argument ('serve --no-browser --data-dir "' + $DataDir + '"') -WorkingDirectory ([IO.Path]::GetDirectoryName($Executable))
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$description = 'Runs the Red Ink Usage Server. Configuration and evidence remain in the explicit data directory.'
if ($AccountMode -eq 'CurrentUser') {
    $user = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $user
    $principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description $description
    Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force:$Replace | Out-Null
} elseif ($AccountMode -eq 'System') {
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId 'NT AUTHORITY\SYSTEM' -LogonType ServiceAccount
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description $description
    Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force:$Replace | Out-Null
} else {
    if ($null -eq $Credential) { $Credential = Get-Credential -Message 'Account for Red Ink Usage Server startup and import-directory access' }
    if ($null -eq $Credential) { throw 'No account credentials were supplied. No task was installed.' }
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Description $description
    $password = $Credential.GetNetworkCredential().Password
    try { Register-ScheduledTask -TaskName $TaskName -InputObject $task -User $Credential.UserName -Password $password -Force:$Replace | Out-Null }
    finally { $password = $null }
}
if ($StartNow) { Start-ScheduledTask -TaskName $TaskName }
Write-Host "Installed '$TaskName'. Data directory: $DataDir"
Write-Host 'The selected account must have write access to the data directory and read access to every configured import source.'
Write-Host 'This is a Task Scheduler task, not a native Windows Service.'
