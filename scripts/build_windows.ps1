param(
    [string]$Version = ""
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $Root ".venv\Scripts\python.exe"

function Invoke-Native {
    param(
        [string]$FilePath,
        [string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code ${LASTEXITCODE}: $FilePath $($Arguments -join ' ')"
    }
}

function Compress-WithRetry {
    param(
        [string[]]$Path,
        [string]$DestinationPath,
        [int]$Retries = 5
    )

    for ($Attempt = 1; $Attempt -le $Retries; $Attempt++) {
        try {
            Compress-Archive -Path $Path -DestinationPath $DestinationPath -Force
            return
        }
        catch {
            if ($Attempt -eq $Retries) {
                throw
            }
            Start-Sleep -Seconds 2
        }
    }
}

if (-not (Test-Path $Python)) {
    python -m venv (Join-Path $Root ".venv")
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to create virtual environment"
    }
}

$ReleaseDir = Join-Path $Root "release"
New-Item -ItemType Directory -Force -Path $ReleaseDir | Out-Null

$LegacyDistDir = Join-Path $Root "dist\hwp2pdf"
if (Test-Path $LegacyDistDir) {
    Remove-Item -LiteralPath $LegacyDistDir -Recurse -Force
}

$LegacyZipPath = Join-Path $ReleaseDir "hwp2pdf-windows.zip"
if (Test-Path $LegacyZipPath) {
    Remove-Item -LiteralPath $LegacyZipPath -Force
}

# Version numbering lives in scripts/set_version.py so the Windows and macOS
# builds can never disagree about what a given build is called. Pass -Version to
# pin an explicit yyyy.MM.dd.N (used when both platforms build the same release).
$SetVersionScript = Join-Path $Root "scripts\set_version.py"
if ($Version) {
    $Version = (& $Python $SetVersionScript $Version | Select-Object -Last 1).ToString().Trim()
}
else {
    $Version = (& $Python $SetVersionScript | Select-Object -Last 1).ToString().Trim()
}
if ($LASTEXITCODE -ne 0 -or -not $Version) {
    throw "Failed to compute build version"
}
Write-Host "Build version: $Version"

Invoke-Native $Python @("-m", "pip", "install", "--upgrade", "pip")
Invoke-Native $Python @("-m", "pip", "install", "-r", (Join-Path $Root "requirements-build.txt"))
Invoke-Native $Python @("-m", "PyInstaller", "--clean", "--noconfirm", (Join-Path $Root "hwp2pdf.spec"))

$DistExe = Join-Path $Root "dist\hwp2pdf.exe"
if (-not (Test-Path $DistExe)) {
    throw "Expected build output not found: $DistExe"
}

$DistCliExe = Join-Path $Root "dist\hwp2pdf-cli.exe"
if (-not (Test-Path $DistCliExe)) {
    throw "Expected CLI build output not found: $DistCliExe"
}

$DistServeExe = Join-Path $Root "dist\hwp2pdf-serve.exe"
if (-not (Test-Path $DistServeExe)) {
    throw "Expected server build output not found: $DistServeExe"
}

$VersionedExe = Join-Path $Root "dist\hwp2pdf-$Version.exe"
$VersionedCliExe = Join-Path $Root "dist\hwp2pdf-cli-$Version.exe"
$VersionedServeExe = Join-Path $Root "dist\hwp2pdf-serve-$Version.exe"
$ZipPath = Join-Path $ReleaseDir "hwp2pdf-windows-$Version.zip"

Move-Item -LiteralPath $DistExe -Destination $VersionedExe -Force
Move-Item -LiteralPath $DistCliExe -Destination $VersionedCliExe -Force
Move-Item -LiteralPath $DistServeExe -Destination $VersionedServeExe -Force
Compress-WithRetry -Path @(
    $VersionedExe,
    $VersionedCliExe,
    $VersionedServeExe,
    (Join-Path $Root "THIRD_PARTY_NOTICES.md")
) -DestinationPath $ZipPath

Write-Host "Version $Version"
Write-Host "Built $VersionedExe"
Write-Host "Built $VersionedCliExe"
Write-Host "Built $VersionedServeExe"
Write-Host "Built $ZipPath"
