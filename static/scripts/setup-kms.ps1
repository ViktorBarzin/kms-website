# setup-kms.ps1
#
# Minimal KMS-host wiring for an already-installed Volume License Windows.
# Pins the KMS host, then activates only if not already licensed.
# Idempotent - safe to re-run: if Windows is already licensed it reports the
# remaining days and exits without re-contacting the KMS server.
# Does NOT install Office. Does NOT change DNS suffix. Pin only.
#
# Usage:
#   iwr -UseBasicParsing https://kms.viktorbarzin.me/scripts/setup-kms.ps1 | iex
#
# Or with a custom KMS host (e.g. self-hosted):
#   $env:KMS_HOST = 'kms.example.com'; iwr ... | iex
#
# Source: https://kms.viktorbarzin.me/scripts/setup-kms.ps1
# Licence: MIT, no warranty, KMS activates Volume License SKUs only.

[CmdletBinding()]
param(
    [string]$KmsHost = $(if ($env:KMS_HOST) { $env:KMS_HOST } else { 'vlmcs.viktorbarzin.me' }),
    [int]   $KmsPort = $(if ($env:KMS_PORT) { [int]$env:KMS_PORT } else { 1688 })
)

$ErrorActionPreference = 'Stop'

function Step($m) { Write-Host "==> $m" -ForegroundColor Cyan }
function OK($m)   { Write-Host "    OK: $m" -ForegroundColor Green }
function Bad($m)  { Write-Host "    !!  $m" -ForegroundColor Red }

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Bad "Must run as Administrator. Right-click PowerShell -> 'Run as administrator', then retry."
    return
}

# Locale-independent license probe (slmgr /dlv text is localized; the WMI
# LicenseStatus integer is not). 1 = Licensed. Returns $null if no KMS-client
# Windows SKU is installed. GracePeriodRemaining is in minutes.
function Get-WindowsLicense {
    $q = "SELECT LicenseStatus, GracePeriodRemaining FROM SoftwareLicensingProduct WHERE Name LIKE 'Windows%' AND PartialProductKey IS NOT NULL"
    $p = $null
    try { $p = Get-CimInstance -Query $q -ErrorAction Stop | Select-Object -First 1 }
    catch { try { $p = Get-WmiObject -Query $q -ErrorAction Stop | Select-Object -First 1 } catch {} }
    if (-not $p) { return $null }
    [pscustomobject]@{ Licensed = ($p.LicenseStatus -eq 1); DaysLeft = [int]([math]::Round($p.GracePeriodRemaining / 1440)) }
}

$slmgr = "$env:WINDIR\System32\slmgr.vbs"

# Pin the KMS host. Idempotent: re-running just re-sets the same value.
Step "Pinning KMS host -> $KmsHost`:$KmsPort"
$out = & cscript //Nologo $slmgr /skms "$KmsHost`:$KmsPort" 2>&1
if ($LASTEXITCODE -ne 0) { Bad "slmgr /skms failed (exit $LASTEXITCODE): $out"; return }
OK "KMS host pinned"

$lic = Get-WindowsLicense
if ($null -eq $lic) {
    Bad "No KMS-client (Volume License) Windows SKU detected."
    Write-Host "    Install a GVLK first:  slmgr /ipk <GVLK>   (see https://kms.viktorbarzin.me/#windows)"
    return
}

# Idempotent: already activated -> don't re-hit the KMS server. The pinned host
# above means Windows keeps auto-renewing every 7 days on its own.
if ($lic.Licensed) {
    Write-Host ""
    Write-Host "==> Already licensed via KMS ($($lic.DaysLeft) days remaining). Nothing to do." -ForegroundColor Green
    Write-Host "    Host is pinned to $KmsHost; Windows auto-renews every 7 days."
    return
}

Step "Activating (slmgr /ato)"
$out = & cscript //Nologo $slmgr /ato 2>&1
Write-Host $out

$lic = Get-WindowsLicense
if ($lic.Licensed) {
    Write-Host ""
    Write-Host "==> SUCCESS: Windows is now licensed via KMS ($($lic.DaysLeft) days)." -ForegroundColor Green
    Write-Host "    Renews automatically every 7 days; lasts 180 days per renewal."
} else {
    Write-Host ""
    Write-Host "==> Activation did not stick." -ForegroundColor Yellow
    Write-Host "    Most common cause: this Windows is not a Volume License edition"
    Write-Host "    (Home / retail / OEM reject KMS). See https://kms.viktorbarzin.me/#faq"
}
