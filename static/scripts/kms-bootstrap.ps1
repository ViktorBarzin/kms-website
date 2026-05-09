# kms-bootstrap.ps1
#
# Interactive activator for a public KMS host (default: kms.viktorbarzin.me:1688).
# Asks what you want to activate (Windows / already-installed Office / Project /
# Visio), and optionally what you want to *install* (Office LTSC 2024 ProPlus,
# Project Pro 2024, Visio Pro 2024 — all VL editions installed via the official
# Microsoft Office Deployment Tool). Runs only what you confirm.
#
# Usage:
#   iwr -UseBasicParsing https://kms.viktorbarzin.me/scripts/kms-bootstrap.ps1 | iex
#
# Non-interactive (CI / automation):
#   $env:KMS_AUTO = 'win,office'; iwr ... | iex
#       (comma list of: win, office, project, visio, install-office,
#                       install-project, install-visio)
#
# Custom KMS host:
#   $env:KMS_HOST = 'kms.example.com'; iwr ... | iex
#
# Source: https://forgejo.viktorbarzin.me/viktor/kms-website
# Licence: MIT, no warranty, KMS activates Volume License SKUs only.

[CmdletBinding()]
param(
    [string]$KmsHost = $(if ($env:KMS_HOST) { $env:KMS_HOST } else { 'kms.viktorbarzin.me' }),
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

$doInstOff   = Choice 'install-office'  "Install Office LTSC 2024 ProPlus (VL, ~3 GB) and activate?" $false
$doInstProj  = Choice 'install-project' "Install Project Pro 2024 (VL) and activate?" $false
$doInstVis   = Choice 'install-visio'   "Install Visio Pro 2024 (VL) and activate?" $false

if (-not ($doWin -or $doOfficeAct -or $doProjAct -or $doVisioAct -or $doInstOff -or $doInstProj -or $doInstVis)) {
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

# --- Install via ODT -----------------------------------------------------
$ODT_URL = 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_19127-20198.exe'

function Install-Odt-Bundle([string[]]$products) {
    $tmp = Join-Path $env:TEMP "kms-odt-$(Get-Random)"
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    $odtExe = Join-Path $tmp 'odt.exe'
    Step "Downloading Office Deployment Tool to $tmp"
    Invoke-WebRequest -UseBasicParsing -Uri $ODT_URL -OutFile $odtExe
    Step "Extracting ODT"
    Start-Process -FilePath $odtExe -ArgumentList "/extract:`"$tmp`"", '/quiet' -Wait
    $setup = Join-Path $tmp 'setup.exe'
    if (-not (Test-Path $setup)) { Bad "ODT extraction failed (no setup.exe in $tmp)"; return }

    # Build Configuration.xml — only the requested VL Products.
    $productXml = ($products | ForEach-Object { "<Product ID=`"$_`"><Language ID=`"en-us`" /></Product>" }) -join ''
    $cfgXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    $productXml
  </Add>
  <Updates Enabled="TRUE" Channel="PerpetualVL2024" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@
    $cfg = Join-Path $tmp 'Configuration.xml'
    Set-Content -Path $cfg -Value $cfgXml -Encoding UTF8

    Step "Running setup.exe /configure (this can take 5-15 min depending on bandwidth)"
    Start-Process -FilePath $setup -ArgumentList '/configure', "`"$cfg`"" -Wait
    if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne $null) { Warn "ODT exit code $LASTEXITCODE" }

    # Pin KMS host + activate
    $ospp = Find-Ospp
    if ($ospp) {
        Step "Pinning Office at $KmsHost`:$KmsPort and activating"
        & cscript //Nologo $ospp /sethst:$KmsHost | Out-Host
        & cscript //Nologo $ospp /setprt:$KmsPort | Out-Host
        & cscript //Nologo $ospp /act             | Out-Host
        $st = & cscript //Nologo $ospp /dstatus 2>&1 | Out-String
        if ($st -match '---LICENSED---') { OK "Office bundle licensed" } else { Warn "Status not LICENSED yet" }
    } else {
        Warn "ospp.vbs still not found post-install — manual /act needed."
    }

    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
}

$installList = @()
if ($doInstOff)  { $installList += 'ProPlus2024Volume'    }
if ($doInstProj) { $installList += 'ProjectPro2024Volume' }
if ($doInstVis)  { $installList += 'VisioPro2024Volume'   }
if ($installList.Count -gt 0) { Install-Odt-Bundle $installList }

Write-Host ""
Step "Done."
Write-Host "    Re-run any time to re-check status. KMS licences renew automatically every 7 days."
Write-Host "    Operator-side: this activation has been logged. Privacy notes: https://kms.viktorbarzin.me/#faq"
