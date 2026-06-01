# setup-kms.ps1
#
# Minimal KMS wiring for Volume License Windows. Pins the KMS host and activates
# only if not already licensed. If no VL key is installed it auto-detects the
# Windows edition and fetches the matching GVLK from the published key list
# (no manual key lookup). Idempotent - safe to re-run: an already-licensed
# machine reports the remaining days and exits without re-contacting KMS.
# Does NOT install Office. Does NOT change DNS suffix.
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
    [int]   $KmsPort = $(if ($env:KMS_PORT) { [int]$env:KMS_PORT } else { 1688 }),
    [string]$KeysUrl = $(if ($env:KMS_KEYS_URL) { $env:KMS_KEYS_URL } else { 'https://kms.viktorbarzin.me/keys.json' })
)

$ErrorActionPreference = 'Stop'

function Step($m) { Write-Host "==> $m" -ForegroundColor Cyan }
function OK($m)   { Write-Host "    OK: $m" -ForegroundColor Green }
function Bad($m)  { Write-Host "    !!  $m" -ForegroundColor Red }

# Anonymous, fire-and-forget diagnostics (see kms-bootstrap.ps1). No PII.
# $env:KMS_NO_TELEMETRY=1 opts out; $env:KMS_DIAG_URL overrides. Version baked at build.
$script:RunId = ([guid]::NewGuid().ToString('N')).Substring(0, 12)
function Send-Diag([string]$action, [string]$outcome, [string]$detail = '') {
    if ($env:KMS_NO_TELEMETRY) { return }
    try {
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
        if ($detail.Length -gt 600) { $detail = $detail.Substring(0, 600) }
        $body = @{ script = 'setup-kms.ps1'; ver = '__KMS_VERSION__'; runid = $script:RunId; ts = (Get-Date).ToUniversalTime().ToString('o'); action = $action; outcome = $outcome; detail = $detail; edition = "$($cv.EditionID)"; build = "$($cv.CurrentBuildNumber)"; locale = (Get-Culture).Name } | ConvertTo-Json -Compress
        $url = if ($env:KMS_DIAG_URL) { $env:KMS_DIAG_URL } else { 'https://kms.viktorbarzin.me/diag' }
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -UseBasicParsing -Method POST -Uri $url -Body $body -ContentType 'application/json' -TimeoutSec 3 | Out-Null
    } catch {}
}

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

# Auto-select the GVLK for THIS machine's edition from the published key list,
# so the user never has to look one up. Matches on the registry EditionID
# (locale-independent); for editions that share an EditionID across releases
# (LTSC, Server) it also narrows by the OS build number.
function Resolve-WindowsGvlk {
    $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    if (-not $cv) { return $null }
    $editionId = $cv.EditionID
    $build = "$($cv.CurrentBuildNumber)"
    $isServer = $false
    try { $isServer = ((Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).ProductType -ne 1) } catch {}
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try { $keys = (Invoke-WebRequest -UseBasicParsing -Uri $KeysUrl -TimeoutSec 20).Content | ConvertFrom-Json }
    catch { Bad "Could not fetch the key list from $KeysUrl"; return $null }
    $pool = if ($isServer) { $keys.windows_server } else { $keys.windows }
    $m = $pool | Where-Object { $_.editionid -eq $editionId -and ( -not $_.builds -or ($_.builds -contains $build) ) } | Select-Object -First 1
    if ($m) { Write-Host "    detected $editionId (build $build) -> $($m.edition)"; return $m.gvlk }
    Write-Host "    no published GVLK matches EditionID '$editionId' (build $build)"
    return $null
}

$slmgr = "$env:WINDIR\System32\slmgr.vbs"

# Pin the KMS host. Idempotent: re-running just re-sets the same value.
Step "Pinning KMS host -> $KmsHost`:$KmsPort"
$out = & cscript //Nologo $slmgr /skms "$KmsHost`:$KmsPort" 2>&1
if ($LASTEXITCODE -ne 0) { Bad "slmgr /skms failed (exit $LASTEXITCODE): $out"; return }
OK "KMS host pinned"

$lic = Get-WindowsLicense

# Idempotent: already activated (retail OR KMS) -> don't re-hit the KMS server and
# never clobber a working key. The pinned host keeps Windows auto-renewing.
if ($lic -and $lic.Licensed) {
    Write-Host ""
    Write-Host "==> Already licensed ($($lic.DaysLeft) days remaining). Nothing to do." -ForegroundColor Green
    Write-Host "    Host is pinned to $KmsHost; Windows auto-renews every 7 days."
    Send-Diag 'win' 'already-licensed'
    return
}

# No Volume License key installed -> fetch + install the GVLK for this edition.
if ($null -eq $lic) {
    Step "No Volume License key installed - detecting edition and fetching its GVLK"
    $gvlk = Resolve-WindowsGvlk
    if (-not $gvlk) {
        Bad "This edition has no KMS GVLK - Home/retail can't KMS-activate as-is."
        Write-Host "    To UPGRADE it in place to a Volume License edition (e.g. Home -> Pro) and"
        Write-Host "    activate, run the interactive bootstrap (it shows the consequences first):"
        Write-Host "      iwr -UseBasicParsing https://kms.viktorbarzin.me/scripts/kms-bootstrap.ps1 | iex"
        Send-Diag 'win' 'no-vl-edition' 'Home/retail - directed to bootstrap'
        return
    }
    Step "Installing GVLK $gvlk"
    $out = & cscript //Nologo $slmgr /ipk $gvlk 2>&1
    if ($LASTEXITCODE -ne 0) { Bad "slmgr /ipk failed (exit $LASTEXITCODE): $out"; Send-Diag 'win' 'fail' 'ipk failed'; return }
    OK "GVLK installed"
}

Step "Activating (slmgr /ato)"
$out = & cscript //Nologo $slmgr /ato 2>&1
Write-Host $out

$lic = Get-WindowsLicense
if ($lic.Licensed) {
    Write-Host ""
    Write-Host "==> SUCCESS: Windows is now licensed via KMS ($($lic.DaysLeft) days)." -ForegroundColor Green
    Write-Host "    Renews automatically every 7 days; lasts 180 days per renewal."
    Send-Diag 'win' 'ok' "$($lic.DaysLeft) days"
} else {
    Write-Host ""
    Write-Host "==> Activation did not stick." -ForegroundColor Yellow
    Write-Host "    Most common cause: this Windows is not a Volume License edition"
    Write-Host "    (Home / retail / OEM reject KMS). See https://kms.viktorbarzin.me/#faq"
    Send-Diag 'win' 'fail' 'ato did not stick (non-VL edition?)'
}
