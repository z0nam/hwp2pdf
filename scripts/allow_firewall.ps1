<#
.SYNOPSIS
    Open the hwp2pdf conversion server port on the Windows Private firewall profile.

.DESCRIPTION
    Only needed for LAN access. If the server is started with `--bind tailscale`
    it listens on the Tailscale interface alone and no rule is required.

    Run from an elevated PowerShell prompt.
#>
param(
    [int]$Port = 8765,
    [string]$DisplayName = "hwp2pdf serve"
)

$ErrorActionPreference = "Stop"

$existing = Get-NetFirewallRule -DisplayName $DisplayName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Rule '$DisplayName' already exists; removing it first."
    $existing | Remove-NetFirewallRule
}

New-NetFirewallRule `
    -DisplayName $DisplayName `
    -Direction Inbound `
    -Protocol TCP `
    -LocalPort $Port `
    -Profile Private `
    -Action Allow | Out-Null

Write-Host "Allowed inbound TCP $Port on the Private profile."
Write-Host "Do NOT widen this to the Public profile or the open internet."
