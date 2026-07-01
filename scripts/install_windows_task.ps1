param(
    [string]$TaskName = "Zensi Weekly Creator Analytics",
    [string]$At = "08:00",
    [string]$Day = "Friday"
)

$ErrorActionPreference = "Stop"
$Root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")
$RunScript = Join-Path $Root "scripts\run_weekly.ps1"
$Config = Join-Path $Root "config.json"

if (-not (Test-Path -LiteralPath $Config)) {
    throw "Missing config.json. Copy config.example.json to config.json and edit it before installing the task."
}

$Action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$RunScript`"" `
    -WorkingDirectory $Root

$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $Day -At $At
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Description "Runs the Zensi creator public social scraper and exports weekly CSV, Excel, and Word reports." `
    -Force | Out-Null

Write-Host "Installed scheduled task '$TaskName' for every $Day at $At."
Write-Host "Run manually with: Start-ScheduledTask -TaskName '$TaskName'"
