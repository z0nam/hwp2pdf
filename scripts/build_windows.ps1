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

$DatePart = Get-Date -Format "yyyy.MM.dd"
$VersionPattern = "^hwp2pdf(?:-windows)?-$([regex]::Escape($DatePart))\.(\d+)(?:\.exe|\.zip)$"
$ExistingNumbers = @()

foreach ($Dir in @((Join-Path $Root "dist"), $ReleaseDir)) {
    if (Test-Path $Dir) {
        Get-ChildItem -LiteralPath $Dir -File | ForEach-Object {
            if ($_.Name -match $VersionPattern) {
                $ExistingNumbers += [int]$Matches[1]
            }
        }
    }
}

$BuildNumber = 1
if ($ExistingNumbers.Count -gt 0) {
    $BuildNumber = ($ExistingNumbers | Measure-Object -Maximum).Maximum + 1
}

$Version = "$DatePart.$BuildNumber"
$VersionFile = Join-Path $Root "src\hwp2pdf\version.py"
Set-Content -LiteralPath $VersionFile -Value "__version__ = `"$Version`"" -Encoding utf8

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

$VersionedExe = Join-Path $Root "dist\hwp2pdf-$Version.exe"
$VersionedCliExe = Join-Path $Root "dist\hwp2pdf-cli-$Version.exe"
$ZipPath = Join-Path $ReleaseDir "hwp2pdf-windows-$Version.zip"

Move-Item -LiteralPath $DistExe -Destination $VersionedExe -Force
Move-Item -LiteralPath $DistCliExe -Destination $VersionedCliExe -Force
Compress-WithRetry -Path @($VersionedExe, $VersionedCliExe) -DestinationPath $ZipPath

Write-Host "Version $Version"
Write-Host "Built $VersionedExe"
Write-Host "Built $VersionedCliExe"
Write-Host "Built $ZipPath"
