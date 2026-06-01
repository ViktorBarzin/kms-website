# kms-bootstrap.ps1
#
# Interactive KMS activator. Asks what you want to activate (Windows /
# already-installed Office / Project / Visio) and runs only what you confirm.
# Points each product at the public KMS host (default: vlmcs.viktorbarzin.me:1688).
#
# Usage:
#   iwr -UseBasicParsing https://kms.viktorbarzin.me/scripts/kms-bootstrap.ps1 | iex
#
# Non-interactive (CI / automation):
#   $env:KMS_AUTO = 'win,office'; iwr ... | iex
#       (comma list of: win, office, project, visio)
#
# Custom KMS host:
#   $env:KMS_HOST = 'kms.example.com'; iwr ... | iex
#
# Source: https://kms.viktorbarzin.me/scripts/kms-bootstrap.ps1
# Licence: MIT, no warranty, KMS activates Volume License SKUs only.

[CmdletBinding()]
param(
    [string]$KmsHost = $(if ($env:KMS_HOST) { $env:KMS_HOST } else { 'vlmcs.viktorbarzin.me' }),
    [int]   $KmsPort = $(if ($env:KMS_PORT) { [int]$env:KMS_PORT } else { 1688 })
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

function Step($m) { Write-Host "==> $m" -ForegroundColor Cyan }
function OK($m)   { Write-Host "    OK: $m" -ForegroundColor Green }
function Warn($m) { Write-Host "    !!  $m" -ForegroundColor Yellow }
function Bad($m)  { Write-Host "    !!  $m" -ForegroundColor Red }

# --- Pre-flight ----------------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Bad "Must run as Administrator. Right-click PowerShell -> 'Run as administrator', then re-run the one-liner."
    return
}

Write-Host ""
Write-Host "  kms.viktorbarzin.me bootstrap" -ForegroundColor White
Write-Host "  KMS host: $KmsHost`:$KmsPort"
Write-Host "  Read the script: https://kms.viktorbarzin.me/scripts/kms-bootstrap.ps1"
Write-Host ""

# --- Decide what to do ---------------------------------------------------
function Ask([string]$question, [bool]$default) {
    $hint = if ($default) { '[Y/n]' } else { '[y/N]' }
    Write-Host -NoNewline "    $question $hint " -ForegroundColor Yellow
    $a = Read-Host
    if ([string]::IsNullOrWhiteSpace($a)) { return $default }
    return $a -match '^[yY]'
}

$auto = @{}
if ($env:KMS_AUTO) { $env:KMS_AUTO.Split(',') | ForEach-Object { $auto[$_.Trim().ToLower()] = $true } }

function Choice([string]$key, [string]$prompt, [bool]$default) {
    if ($auto.Count -gt 0) { return [bool]$auto[$key] }
    return Ask $prompt $default
}

Step "What would you like to do?"
$doWin       = Choice 'win'             "Activate this Windows installation against KMS?" $true
$doOfficeAct = Choice 'office'          "Activate an already-installed Office (Pro Plus 2024 / 2021 / 2019 / 2016)?" $false
$doProjAct   = Choice 'project'         "Activate an already-installed Project (Pro 2024 / 2021 / 2019 / 2016)?" $false
$doVisioAct  = Choice 'visio'           "Activate an already-installed Visio (Pro 2024 / 2021 / 2019 / 2016)?" $false

if (-not ($doWin -or $doOfficeAct -or $doProjAct -or $doVisioAct)) {
    Warn "Nothing selected. Exiting."
    return
}

# --- Windows -------------------------------------------------------------
function Activate-Windows {
    Step "Windows activation"
    $slmgr = "$env:WINDIR\System32\slmgr.vbs"
    & cscript //Nologo $slmgr /skms "$KmsHost`:$KmsPort" | Out-Host
    if ($LASTEXITCODE -ne 0) { Bad "slmgr /skms failed"; return }
    & cscript //Nologo $slmgr /ato | Out-Host
    if ($LASTEXITCODE -ne 0) {
        Bad "slmgr /ato failed. Most likely cause: not a Volume License edition (Home/retail/OEM cannot KMS-activate)."
        Write-Host "    See https://kms.viktorbarzin.me/#faq"
        return
    }
    $dlv = & cscript //Nologo $slmgr /dlv 2>&1 | Out-String
    if ($dlv -match 'License Status:\s*Licensed') { OK "Windows licensed" } else { Warn "Status not 'Licensed' yet — try 'slmgr /dlv' in a minute" }
}
if ($doWin) { Activate-Windows }

# --- Office / Project / Visio: activate already-installed ----------------
function Find-Ospp {
    $candidates = @(
        "${env:ProgramFiles}\Microsoft Office\Office16\ospp.vbs",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\ospp.vbs",
        "${env:ProgramFiles}\Microsoft Office\Office15\ospp.vbs",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15\ospp.vbs"
    )
    foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
    return $null
}

function Activate-Ospp([string]$label) {
    $ospp = Find-Ospp
    if (-not $ospp) {
        Warn "$label`: ospp.vbs not found (Office not installed?). Skipping."
        return
    }
    Step "$label activation via $ospp"
    & cscript //Nologo $ospp /sethst:$KmsHost | Out-Host
    & cscript //Nologo $ospp /setprt:$KmsPort | Out-Host
    & cscript //Nologo $ospp /act             | Out-Host
    $st = & cscript //Nologo $ospp /dstatus 2>&1 | Out-String
    if ($st -match '---LICENSED---') { OK "$label licensed" } else { Warn "$label status not LICENSED yet" }
}
if ($doOfficeAct) { Activate-Ospp 'Office'  }
if ($doProjAct)   { Activate-Ospp 'Project' }
if ($doVisioAct)  { Activate-Ospp 'Visio'   }

Write-Host ""
Step "Done."
Write-Host "    Re-run any time to re-check status. KMS licences renew automatically every 7 days."
Write-Host "    Privacy: see https://kms.viktorbarzin.me/#faq"
