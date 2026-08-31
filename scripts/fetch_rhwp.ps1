<#
.SYNOPSIS
    Download the pinned rhwp release binary and install it to vendor\rhwp.

.DESCRIPTION
    rhwp is only needed for the optional local fallback that renders PDFs when
    the conversion server cannot be reached. hwp2pdf works without it.

    Uses only Invoke-WebRequest -- no GitHub CLI and no account, since the
    releases are public.
#>
param(
    [string]$Version = "v0.8.4",
    [string]$Repo = "edwardkim/rhwp"
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $PSScriptRoot
$Dest = Join-Path $Root "vendor\rhwp"
$Base = "https://github.com/$Repo/releases/download/$Version"
$Asset = "rhwp-$Version-windows-x86_64.zip"

$Tmp = Join-Path ([IO.Path]::GetTempPath()) ([Guid]::NewGuid().ToString())
New-Item -ItemType Directory -Force -Path $Tmp | Out-Null
try {
    Write-Host "Downloading $Asset ($Version)..."
    Invoke-WebRequest -Uri "$Base/$Asset" -OutFile (Join-Path $Tmp $Asset) -UseBasicParsing
    Invoke-WebRequest -Uri "$Base/SHA256SUMS.txt" -OutFile (Join-Path $Tmp "SHA256SUMS.txt") -UseBasicParsing

    Write-Host "Verifying checksum..."
    $want = (Get-Content (Join-Path $Tmp "SHA256SUMS.txt") |
        Where-Object { $_ -match [regex]::Escape($Asset) + '\s*$' } |
        ForEach-Object { ($_ -split '\s+')[0] } | Select-Object -First 1)
    $got = (Get-FileHash (Join-Path $Tmp $Asset) -Algorithm SHA256).Hash.ToLower()
    if (-not $want -or $want.ToLower() -ne $got) {
        throw "checksum mismatch for ${Asset}: want=$want got=$got"
    }
    Write-Host "  ok: $got"

    Write-Host "Extracting to $Dest..."
    Expand-Archive -Path (Join-Path $Tmp $Asset) -DestinationPath $Tmp -Force
    New-Item -ItemType Directory -Force -Path $Dest | Out-Null
    # The archive holds a top-level rhwp\ directory with the binary and LICENSE.
    Copy-Item (Join-Path $Tmp "rhwp\rhwp.exe") (Join-Path $Dest "rhwp.exe") -Force
    Copy-Item (Join-Path $Tmp "rhwp\LICENSE") (Join-Path $Dest "LICENSE") -Force -ErrorAction SilentlyContinue
    Copy-Item (Join-Path $Tmp "SHA256SUMS.txt") (Join-Path $Dest "SHA256SUMS.txt") -Force

    Write-Host ("Installed: " + (& (Join-Path $Dest "rhwp.exe") --help 2>&1 | Select-Object -First 1))
    Write-Host ""
    Write-Host "hwp2pdf finds this automatically. Enable the fallback with --rhwp-fallback,"
    Write-Host "or the matching checkbox in the GUI options."
}
finally {
    Remove-Item -Recurse -Force $Tmp -ErrorAction SilentlyContinue
}
