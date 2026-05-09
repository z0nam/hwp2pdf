$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $PSScriptRoot
$VersionFile = Join-Path $Root "src\hwp2pdf\version.py"
$InstallerScript = Join-Path $Root "installer\hwp2pdf.iss"

if (-not (Test-Path $VersionFile)) {
    throw "Version file not found: $VersionFile"
}

if (-not (Test-Path $InstallerScript)) {
    throw "Installer script not found: $InstallerScript"
}

$Version = $null
foreach ($Line in Get-Content -LiteralPath $VersionFile) {
    if ($Line -match "__version__\s*=\s*`"([^`"]+)`"") {
        $Version = $Matches[1]
        break
    }
}
if (-not $Version) {
    throw "Could not read __version__ from $VersionFile"
}
$VersionedExe = Join-Path $Root "dist\hwp2pdf-$Version.exe"
if (-not (Test-Path $VersionedExe)) {
    throw "Versioned exe not found. Run scripts\build_windows.ps1 first: $VersionedExe"
}

$VersionedCliExe = Join-Path $Root "dist\hwp2pdf-cli-$Version.exe"
if (-not (Test-Path $VersionedCliExe)) {
    throw "Versioned CLI exe not found. Run scripts\build_windows.ps1 first: $VersionedCliExe"
}

$Iscc = (Get-Command iscc -ErrorAction SilentlyContinue).Source
if (-not $Iscc) {
    $Candidates = @(
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe",
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe")
    )
    foreach ($Candidate in $Candidates) {
        if (Test-Path $Candidate) {
            $Iscc = $Candidate
            break
        }
    }
}

if (-not $Iscc) {
    throw "Inno Setup 6 compiler (ISCC.exe) was not found. Install Inno Setup 6, then run this script again."
}

$env:HWP2PDF_VERSION = $Version
$env:HWP2PDF_ROOT = $Root

& $Iscc $InstallerScript
if ($LASTEXITCODE -ne 0) {
    throw "Inno Setup failed with exit code $LASTEXITCODE"
}

$SetupPath = Join-Path $Root "release\hwp2pdf-setup-$Version.exe"
if (-not (Test-Path $SetupPath)) {
    throw "Expected installer output not found: $SetupPath"
}

Write-Host "Built $SetupPath"
