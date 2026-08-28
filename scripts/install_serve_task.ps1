<#
.SYNOPSIS
    Register the hwp2pdf conversion server to start at logon.

.DESCRIPTION
    Deliberately a Scheduled Task and not a Windows Service: Hangul COM
    automation needs an interactive desktop session. A service runs in Session 0,
    which has no desktop, and leaves zombie Hwp.exe processes behind
    (see docs/known-issues.md).

    Run from an elevated PowerShell prompt.

    NOTE: the server is a console program, so it shows a console window on the
    desktop. Closing that window terminates the server (the task then reports
    LastTaskResult 0xC000013A, STATUS_CONTROL_C_EXIT). Minimize it, do not close
    it. Re-run the task with: Start-ScheduledTask -TaskName '<name>'
#>
param(
    [string]$Exe = "",
    [string]$Bind = "tailscale",
    [int]$Port = 8765,
    [string]$TaskName = "hwp2pdf serve"
)

$ErrorActionPreference = "Stop"

if (-not $Exe) {
    $Root = Split-Path -Parent $PSScriptRoot
    foreach ($Candidate in @(
        (Join-Path $Root "dist\hwp2pdf-cli.exe"),
        (Join-Path ${env:ProgramFiles} "hwp2pdf\hwp2pdf-cli.exe")
    )) {
        if (Test-Path $Candidate) { $Exe = $Candidate; break }
    }
}
if (-not $Exe -or -not (Test-Path $Exe)) {
    throw "Could not find hwp2pdf-cli.exe. Pass -Exe <path>."
}

$Action = New-ScheduledTaskAction -Execute $Exe -Argument "serve --bind $Bind --port $Port"
$Trigger = New-ScheduledTaskTrigger -AtLogOn
# Interactive == "run only when user is logged on": the whole point here, so the
# task lands in the desktop session instead of Session 0.
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero)

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger `
    -Principal $Principal -Settings $Settings -Force | Out-Null

Write-Host "Registered scheduled task '$TaskName':"
Write-Host "  $Exe serve --bind $Bind --port $Port"
Write-Host "  runs at logon, only while $env:USERNAME is logged on."
Write-Host ""
Write-Host "A console window will appear on the desktop. Minimize it -- closing it"
Write-Host "stops the server. Restart with: Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "Remove it with: Unregister-ScheduledTask -TaskName '$TaskName'"
