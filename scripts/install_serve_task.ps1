<#
.SYNOPSIS
    Register the hwp2pdf conversion server to start at logon.

.DESCRIPTION
    Deliberately a Scheduled Task and not a Windows Service: Hangul COM
    automation needs an interactive desktop session. A service runs in Session 0,
    which has no desktop, and leaves zombie Hwp.exe processes behind
    (see docs/known-issues.md).

    Run from an elevated PowerShell prompt.

    Prefers hwp2pdf-serve.exe, which is windowless: there is no console window
    to close by accident. It writes its output to
    %LOCALAPPDATA%\hwp2pdf\server.log instead.

    If only the console build (hwp2pdf-cli.exe) is present it is used as a
    fallback, and closing its window stops the server (the task then reports
    LastTaskResult 0xC000013A, STATUS_CONTROL_C_EXIT). Either way the task is
    configured to restart the server if it exits unexpectedly.
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
    # Windowless build first; the console build only as a fallback. build_windows.ps1
    # stamps the version into the file name, so match those too and take the newest.
    foreach ($Pattern in @(
        (Join-Path $Root "dist\hwp2pdf-serve*.exe"),
        (Join-Path ${env:ProgramFiles} "hwp2pdf\hwp2pdf-serve*.exe"),
        (Join-Path $Root "dist\hwp2pdf-cli*.exe"),
        (Join-Path ${env:ProgramFiles} "hwp2pdf\hwp2pdf-cli*.exe")
    )) {
        $Found = Get-ChildItem -Path $Pattern -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
        if ($Found) { $Exe = $Found.FullName; break }
    }
}
if (-not $Exe -or -not (Test-Path $Exe)) {
    throw "Could not find hwp2pdf-serve.exe or hwp2pdf-cli.exe. Pass -Exe <path>."
}

$Windowless = [IO.Path]::GetFileNameWithoutExtension($Exe) -like "hwp2pdf-serve*"
$Arguments = if ($Windowless) { "--bind $Bind --port $Port" } else { "serve --bind $Bind --port $Port" }
$Action = New-ScheduledTaskAction -Execute $Exe -Argument $Arguments

# Three triggers, because there are three ways to end up with no server:
#   - at logon: the normal case, including the sign-in after an update restart
#   - at startup: covers a boot where the session is already established
#   - every 10 minutes: self-heals anything else (a crash, a closed console
#     window, Tailscale not being up yet when the logon trigger fired).
# MultipleInstances Ignore means the repeat is a no-op while it is running.
$Triggers = @(
    (New-ScheduledTaskTrigger -AtLogOn),
    (New-ScheduledTaskTrigger -AtStartup)
)
$Repeat = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(2) `
    -RepetitionInterval ([TimeSpan]::FromMinutes(10))
$Triggers += $Repeat
# Interactive == "run only when user is logged on": the whole point here, so the
# task lands in the desktop session instead of Session 0.
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
# RestartCount/RestartInterval cover both a crash and someone closing the
# console window of the fallback build.
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -RestartCount 3 -RestartInterval ([TimeSpan]::FromMinutes(1)) `
    -MultipleInstances IgnoreNew -StartWhenAvailable

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Triggers `
    -Principal $Principal -Settings $Settings -Force | Out-Null

Write-Host "Registered scheduled task '$TaskName':"
Write-Host "  $Exe $Arguments"
Write-Host "  at logon, at startup, and re-checked every 10 minutes"
Write-Host "  (only while $env:USERNAME is logged on -- Hangul needs a desktop)."
Write-Host ""
if ($Windowless) {
    Write-Host "Windowless build: no console window. Output goes to"
    Write-Host "  $env:LOCALAPPDATA\hwp2pdf\server.log"
} else {
    Write-Host "Console build: a window will appear on the desktop. Minimize it --"
    Write-Host "closing it stops the server (the task will restart it within a minute)."
}
Write-Host "Start now with: Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "Remove it with: Unregister-ScheduledTask -TaskName '$TaskName'"
