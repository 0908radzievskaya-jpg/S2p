[CmdletBinding()]
param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path,
    [string]$RemoteHost = "194.113.209.237",
    [string]$RemoteUser = "root",
    [string]$RemoteAppDir = "/opt/tender-dashboard",
    [string]$SshKeyPath = "",
    [string]$TaskName = "Tender Dashboard Data Sync",
    [int]$EveryMinutes = 15,
    [switch]$RunNow
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if ($EveryMinutes -lt 1) {
    throw "EveryMinutes must be greater than 0."
}

$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$syncScript = Join-Path $PSScriptRoot "sync_dashboard_data.ps1"
if (-not (Test-Path -LiteralPath $syncScript)) {
    throw "Sync script not found: $syncScript"
}

$pwsh = (Get-Command powershell.exe -ErrorAction Stop).Source
$arguments = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", "`"$syncScript`"",
    "-ProjectRoot", "`"$ProjectRoot`"",
    "-RemoteHost", "`"$RemoteHost`"",
    "-RemoteUser", "`"$RemoteUser`"",
    "-RemoteAppDir", "`"$RemoteAppDir`""
)

if ($SshKeyPath) {
    $arguments += @("-SshKeyPath", "`"$SshKeyPath`"")
}

$action = New-ScheduledTaskAction -Execute $pwsh -Argument ($arguments -join " ") -WorkingDirectory $ProjectRoot
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes $EveryMinutes) `
    -RepetitionDuration (New-TimeSpan -Days 3650)
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description "Uploads Tender Dashboard reports and relevant tender folders to the VPS." `
    -Force | Out-Null

Write-Host "Scheduled task installed: $TaskName"
Write-Host "Interval: every $EveryMinutes minutes"

if ($RunNow) {
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "Task started."
}
