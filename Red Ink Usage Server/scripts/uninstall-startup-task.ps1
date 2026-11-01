#requires -Version 5.1
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param([string]$TaskName = 'Red Ink Usage Server')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Import-Module ScheduledTasks -ErrorAction Stop
$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -eq $task) { Write-Host "Task '$TaskName' is not installed."; return }
$executables = @($task.Actions | ForEach-Object { [IO.Path]::GetFileName($_.Execute) })
if ($executables.Count -ne 1 -or $executables[0] -ne 'RedInkUsageServer.exe') { throw 'The named task does not run RedInkUsageServer.exe. Review it manually; it was not changed.' }
if ($PSCmdlet.ShouldProcess($TaskName, 'Stop and unregister server startup task')) {
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host 'Startup task removed. Application files and data have been retained.'
}
