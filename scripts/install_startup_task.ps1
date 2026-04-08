param(
    [string]$TaskName = "SekigaeStreamlit",
    [int]$Port = 8501,
    [string]$BindAddress = "127.0.0.1"
)

$ErrorActionPreference = "Stop"

$scriptPath = Join-Path $PSScriptRoot "run_streamlit_background.ps1"
if (-not (Test-Path $scriptPath)) {
    throw "run_streamlit_background.ps1 not found: $scriptPath"
}

$startupMode = ""
try {
    $argument = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -Port $Port -BindAddress $BindAddress"
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argument
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -MultipleInstances IgnoreNew `
        -RestartCount 3 `
        -RestartInterval (New-TimeSpan -Minutes 1)

    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -Description "Run sekigae Streamlit app in background at logon." `
        -Force | Out-Null

    $startupMode = "ScheduledTask"
} catch {
    $startupFolder = [Environment]::GetFolderPath("Startup")
    $projectRoot = Split-Path -Parent $PSScriptRoot
    $launcherPath = Join-Path $startupFolder "$TaskName.cmd"
    $launcherContent = @"
@echo off
cd /d "$projectRoot"
powershell -NoProfile -ExecutionPolicy Bypass -File "$scriptPath" -Port $Port -BindAddress $BindAddress
"@
    Set-Content -Path $launcherPath -Value $launcherContent -Encoding Ascii
    $startupMode = "StartupFolder"
    Write-Warning "Scheduled task registration was denied. Installed startup launcher: $launcherPath"
}

& $scriptPath -Port $Port -BindAddress $BindAddress | Out-Host

Write-Output "Installed startup mode: $startupMode"
