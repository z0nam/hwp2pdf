$ErrorActionPreference = "Stop"

function Write-Check {
    param(
        [string]$Name,
        [bool]$Ok,
        [string]$Detail = ""
    )

    $Status = if ($Ok) { "OK" } else { "FAIL" }
    if ($Detail) {
        Write-Host "[$Status] $Name - $Detail"
    } else {
        Write-Host "[$Status] $Name"
    }
}

$Root = Split-Path -Parent $PSScriptRoot
$Failed = $false

Write-Host "hwp2pdf Windows preflight"
Write-Host "Root: $Root"
Write-Host ""

$IsWindows = $PSVersionTable.Platform -eq "Win32NT" -or $env:OS -eq "Windows_NT"
Write-Check "Windows OS" $IsWindows
if (-not $IsWindows) {
    $Failed = $true
}

$PythonCmd = Get-Command python -ErrorAction SilentlyContinue
Write-Check "python on PATH" ($null -ne $PythonCmd) $(if ($PythonCmd) { $PythonCmd.Source } else { "" })
if (-not $PythonCmd) {
    $Failed = $true
} else {
    & python --version
}

Push-Location $Root
try {
    $ImportResult = & python -c "import sys; sys.path.insert(0, 'src'); import hwp2pdf; print(hwp2pdf.__version__)"
    Write-Check "source package import" ($LASTEXITCODE -eq 0) $ImportResult
    if ($LASTEXITCODE -ne 0) {
        $Failed = $true
    }

    & python -c "import win32com.client, pythoncom; print('pywin32 import ok')" 2>$null
    $Pywin32Ok = $LASTEXITCODE -eq 0
    Write-Check "pywin32 import" $Pywin32Ok
    if (-not $Pywin32Ok) {
        $Failed = $true
        Write-Host "Install with: python -m pip install -r requirements.txt"
    }

    if ($Pywin32Ok) {
        & python -c "import pythoncom, win32com.client; pythoncom.CoInitialize(); hwp = win32com.client.Dispatch('HWPFrame.HwpObject'); print('HWPFrame.HwpObject dispatch ok'); hwp.Quit(); pythoncom.CoUninitialize()" 2>$null
        $ComOk = $LASTEXITCODE -eq 0
        Write-Check "Hancom HWP COM dispatch" $ComOk
        if (-not $ComOk) {
            $Failed = $true
            Write-Host "If Hancom Office is installed, try running Hwp.exe -regserver once."
        }
    }

    if (Test-Path "hwp2pdf.spec") {
        Write-Check "PyInstaller spec present" $true "hwp2pdf.spec"
    } else {
        Write-Check "PyInstaller spec present" $false
        $Failed = $true
    }
}
finally {
    Pop-Location
}

Write-Host ""
if ($Failed) {
    Write-Host "Preflight failed. Fix the failed checks above before building or converting."
    exit 1
}

Write-Host "Preflight passed. You can run from source or build with scripts\build_windows.ps1."
