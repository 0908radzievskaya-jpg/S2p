[CmdletBinding()]
param(
    [string]$ProjectRoot = "",
    [string]$RemoteHost = "194.113.209.237",
    [string]$RemoteUser = "root",
    [string]$RemoteAppDir = "/opt/tender-dashboard",
    [string]$SshKeyPath = "",
    [ValidateSet("Minimal", "Full")]
    [string]$Mode = "Minimal",
    [string]$TaskName = "Tender Dashboard Data Sync",
    [string]$DailyAt = "09:00",
    [int]$EveryMinutes = 0,
    [int]$MaxRunMinutes = 120,
    [switch]$RunNow
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if (-not $ProjectRoot) {
    $ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
}

if ($EveryMinutes -lt 0) {
    throw "EveryMinutes must be 0 or greater."
}
if ($MaxRunMinutes -lt 1) {
    throw "MaxRunMinutes must be greater than 0."
}

$culture = [System.Globalization.CultureInfo]::InvariantCulture
$parsedDailyAt = [datetime]::MinValue
$dailyAtFormats = [string[]]@("H:mm", "HH:mm")
if (-not [datetime]::TryParseExact($DailyAt, $dailyAtFormats, $culture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedDailyAt)) {
    throw "DailyAt must be in HH:mm format, for example 09:00."
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
    "-RemoteAppDir", "`"$RemoteAppDir`"",
    "-Mode", "`"$Mode`""
)

if ($SshKeyPath) {
    $arguments += @("-SshKeyPath", "`"$SshKeyPath`"")
}

$action = New-ScheduledTaskAction -Execute $pwsh -Argument ($arguments -join " ") -WorkingDirectory $ProjectRoot
if ($EveryMinutes -gt 0) {
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
        -RepetitionInterval (New-TimeSpan -Minutes $EveryMinutes) `
        -RepetitionDuration (New-TimeSpan -Days 3650)
    $scheduleDescription = "every $EveryMinutes minutes"
}
else {
    $trigger = New-ScheduledTaskTrigger -Daily -At ((Get-Date).Date.Add($parsedDailyAt.TimeOfDay))
    $scheduleDescription = "daily at $DailyAt"
}
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes $MaxRunMinutes)

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description "Uploads Tender Dashboard reports and relevant tender folders to the VPS." `
    -Force | Out-Null

Write-Host "Scheduled task installed: $TaskName"
Write-Host "Schedule: $scheduleDescription"
Write-Host "Max run time: $MaxRunMinutes minutes"

if ($RunNow) {
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "Task started."
}
