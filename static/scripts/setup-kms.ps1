# setup-kms.ps1
#
# Minimal KMS-host wiring for an already-installed Volume License Windows.
# Runs `slmgr /skms <host>:<port>` + `slmgr /ato` and prints the licence status.
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
    [string]$KmsHost = $(if ($env:KMS_HOST) { $env:KMS_HOST } else { 'kms.viktorbarzin.me' }),
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

Step "KMS host = $KmsHost`:$KmsPort"
$slmgr = "$env:WINDIR\System32\slmgr.vbs"

Step "slmgr /skms $KmsHost`:$KmsPort"
$out = & cscript //Nologo $slmgr /skms "$KmsHost`:$KmsPort" 2>&1
Write-Host $out
if ($LASTEXITCODE -ne 0) { Bad "slmgr /skms failed (exit $LASTEXITCODE)"; return }
OK "KMS host pinned"

Step "slmgr /ato (activate)"
$out = & cscript //Nologo $slmgr /ato 2>&1
Write-Host $out
if ($LASTEXITCODE -ne 0) {
    Bad "slmgr /ato failed (exit $LASTEXITCODE)"
    Write-Host ""
    Write-Host "Most common cause: this Windows is not a Volume License edition."
    Write-Host "KMS activates only VL SKUs (Pro, Enterprise, Education, LTSC, Server)."
    Write-Host "Home / retail / OEM keys reject KMS responses. See https://kms.viktorbarzin.me/#faq"
    return
}
OK "Activation request sent"

Step "slmgr /dlv (status)"
$out = & cscript //Nologo $slmgr /dlv 2>&1
Write-Host $out

if ($out -match 'License Status:\s*Licensed') {
    Write-Host ""
    Write-Host "==> SUCCESS: Windows is now licensed via KMS." -ForegroundColor Green
    Write-Host "    Licence renews automatically every 7 days; lasts 180 days per renewal."
} else {
    Write-Host ""
    Write-Host "==> Activation request sent but status is not 'Licensed' yet." -ForegroundColor Yellow
    Write-Host "    Re-run 'slmgr /dlv' in a minute, or check https://kms.viktorbarzin.me/#faq"
}
