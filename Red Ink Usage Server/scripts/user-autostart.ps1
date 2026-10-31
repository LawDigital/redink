#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('status','enable','disable')][string]$Action,
    [Parameter(Mandatory=$true)][string]$Executable,
    [Parameter(Mandatory=$true)][string]$DataDir,
    [string]$ScriptPath = ''
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Import-Module ScheduledTasks -ErrorAction Stop
$taskName = 'RedInkUsageServer-CurrentUser'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$user = $identity.Name
if ($identity.IsSystem -or $identity.User.Value -in @('S-1-5-19','S-1-5-20')) { throw 'Current-user autostart requires an interactive Windows user, not a service account.' }
$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($null -ne $existing) {
    if ($existing.Principal.UserId -ne $user) { throw 'Autostart task belongs to a different account.' }
    if (@($existing.Actions).Count -ne 1) { throw 'Autostart task has unexpected actions.' }
    if ($existing.Description -ne 'Red Ink Usage Server at Windows logon; headless') {
        throw 'The existing named task is not the Usage Server current-user task. It was not modified.'
    }
}
if ($Action -eq 'status') {
    if ($null -ne $existing -and $existing.State -ne 'Disabled') { 'ENABLED' } else { 'DISABLED' }
    exit 0
}
if ($Action -eq 'disable') {
    if ($null -ne $existing) { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop }
    'DISABLED'
    exit 0
}
# Ensure only this installation and data directory are ever registered.
$exe = (Resolve-Path -LiteralPath $Executable -ErrorAction Stop).Path
$data = (Resolve-Path -LiteralPath $DataDir -ErrorAction Stop).Path
$arguments = ''
if ($ScriptPath -ne '') {
    $entry = (Resolve-Path -LiteralPath $ScriptPath -ErrorAction Stop).Path
    if ([IO.Path]::GetFileName($entry) -ne 'run_server.py') { throw 'Unexpected Python entrypoint.' }
    $arguments = '"' + $entry + '" '
} elseif ([IO.Path]::GetFileName($exe) -ne 'RedInkUsageServer.exe') {
    throw 'Unexpected server executable.'
}
if ($data.Contains('"') -or $exe.Contains('"') -or $ScriptPath.Contains('"')) { throw 'Invalid quote in path.' }
$arguments += 'serve --no-browser --data-dir "' + $data + '"'
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $user
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$taskAction = New-ScheduledTaskAction -Execute $exe -Argument $arguments -WorkingDirectory ([IO.Path]::GetDirectoryName($exe))
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit ([TimeSpan]::Zero)
$task = New-ScheduledTask -Action $taskAction -Trigger $trigger -Principal $principal -Settings $settings -Description 'Red Ink Usage Server at Windows logon; headless'
Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null
'ENABLED'
