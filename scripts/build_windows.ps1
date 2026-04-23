$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $Root ".venv\Scripts\python.exe"

if (-not (Test-Path $Python)) {
    python -m venv (Join-Path $Root ".venv")
}

& $Python -m pip install --upgrade pip
& $Python -m pip install -r (Join-Path $Root "requirements-build.txt")
& $Python -m pip install -e $Root
& $Python -m PyInstaller --clean --noconfirm (Join-Path $Root "hwp2pdf.spec")

$ReleaseDir = Join-Path $Root "release"
New-Item -ItemType Directory -Force -Path $ReleaseDir | Out-Null

$ZipPath = Join-Path $ReleaseDir "hwp2pdf-windows.zip"
if (Test-Path $ZipPath) {
    Remove-Item $ZipPath
}

Compress-Archive -Path (Join-Path $Root "dist\hwp2pdf\*") -DestinationPath $ZipPath
Write-Host "Built $ZipPath"
