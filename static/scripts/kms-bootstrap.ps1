# kms-bootstrap.ps1
#
# Interactive KMS activator. Asks what you want to activate (Windows /
# already-installed Office / Project / Visio) and runs only what you confirm.
# Points each product at the public KMS host (default: vlmcs.viktorbarzin.me:1688).
# When a Volume License key is missing it auto-detects the edition/product and
# fetches the matching GVLK from the published key list (no manual key lookup).
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
    [int]   $KmsPort = $(if ($env:KMS_PORT) { [int]$env:KMS_PORT } else { 1688 }),
    [string]$KeysUrl = $(if ($env:KMS_KEYS_URL) { $env:KMS_KEYS_URL } else { 'https://kms.viktorbarzin.me/keys.json' })
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

# Show the consequences of a destructive/heavy action and get consent (default No).
# Interactive -> prompt. Non-interactive (KMS_AUTO set or no console) -> proceed
# ONLY if an explicit env override gave consent ($envConsent), else skip.
function Approve([string]$consequences, [bool]$envConsent) {
    Write-Host ""
    Write-Host $consequences -ForegroundColor Yellow
    $interactive = ($auto.Count -eq 0) -and [Environment]::UserInteractive
    if ($interactive) { return (Ask "Proceed?" $false) }
    if ($envConsent)  { Write-Host "    (non-interactive; proceeding on explicit env override)"; return $true }
    Warn "Skipped (non-interactive and no explicit env override to consent)."
    return $false
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
# Locale-independent license probe (slmgr /dlv text is localized; the WMI
# LicenseStatus integer is not). 1 = Licensed. $null = no KMS-client SKU.
function Get-WindowsLicense {
    $q = "SELECT LicenseStatus, GracePeriodRemaining FROM SoftwareLicensingProduct WHERE Name LIKE 'Windows%' AND PartialProductKey IS NOT NULL"
    $p = $null
    try { $p = Get-CimInstance -Query $q -ErrorAction Stop | Select-Object -First 1 }
    catch { try { $p = Get-WmiObject -Query $q -ErrorAction Stop | Select-Object -First 1 } catch {} }
    if (-not $p) { return $null }
    [pscustomobject]@{ Licensed = ($p.LicenseStatus -eq 1); DaysLeft = [int]([math]::Round($p.GracePeriodRemaining / 1440)) }
}

# Fetch the published GVLK list once (single source of truth, no hardcoding).
$script:KeysCache = $null
function Get-Keys {
    if ($script:KeysCache) { return $script:KeysCache }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try { $script:KeysCache = (Invoke-WebRequest -UseBasicParsing -Uri $KeysUrl -TimeoutSec 20).Content | ConvertFrom-Json }
    catch { Warn "Could not fetch the key list from $KeysUrl"; $script:KeysCache = $null }
    return $script:KeysCache
}

# Auto-select the GVLK for THIS machine's edition (registry EditionID, locale-
# independent; narrowed by build for LTSC/Server which share an EditionID).
function Resolve-WindowsGvlk {
    $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    if (-not $cv) { return $null }
    $editionId = $cv.EditionID; $build = "$($cv.CurrentBuildNumber)"
    $isServer = $false
    try { $isServer = ((Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).ProductType -ne 1) } catch {}
    $keys = Get-Keys; if (-not $keys) { return $null }
    $pool = if ($isServer) { $keys.windows_server } else { $keys.windows }
    $m = $pool | Where-Object { $_.editionid -eq $editionId -and ( -not $_.builds -or ($_.builds -contains $build) ) } | Select-Object -First 1
    if ($m) { Write-Host "    detected $editionId (build $build) -> $($m.edition)"; return $m.gvlk }
    Write-Host "    no published GVLK matches EditionID '$editionId' (build $build)"
    return $null
}

# The current edition can't KMS-activate (Home/retail). Offer an in-place edition
# UPGRADE to a Volume License edition (default Pro; $env:KMS_EDITION overrides).
function Upgrade-WindowsEdition {
    $cv  = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    $cur = if ($cv) { $cv.EditionID } else { 'this edition' }
    $keys = Get-Keys; if (-not $keys) { Bad "No key list available - cannot pick a target edition."; return }
    $targetName = if ($env:KMS_EDITION) { $env:KMS_EDITION } else { 'Pro' }
    $t = $keys.windows | Where-Object { $_.edition -eq $targetName -or $_.editionid -eq $targetName } | Select-Object -First 1
    if (-not $t) { Bad "Target edition '$targetName' is not in the published list. Set `$env:KMS_EDITION to a listed edition (Pro, Enterprise, Education, ...)."; return }
    $text = @"
This Windows edition ($cur) cannot be activated by KMS.
It can be UPGRADED in place to a Volume License edition:

    $cur  ->  $($t.edition)   (GVLK $($t.gvlk))

Consequences:
  * Runs  changepk.exe /ProductKey <GVLK>  - an in-place edition UPGRADE
  * REQUIRES A REBOOT to complete; after reboot, re-run this one-liner to activate
  * One-way: you cannot downgrade back to $cur without a full reinstall
  * Only works along Microsoft's supported upgrade paths (e.g. Home -> Pro -> Enterprise)
"@
    if (-not (Approve $text ([bool]$env:KMS_EDITION))) { Warn "Edition upgrade skipped; $cur cannot KMS-activate as-is."; return }
    $changepk = "$env:WINDIR\System32\changepk.exe"
    if (-not (Test-Path $changepk)) { Bad "changepk.exe not found on this OS - cannot upgrade edition automatically."; return }
    Step "Upgrading $cur -> $($t.edition) (changepk.exe /ProductKey ...)"
    & $changepk /ProductKey $($t.gvlk)
    OK "Edition upgrade started. REBOOT, then re-run this one-liner to activate $($t.edition)."
}

function Activate-Windows {
    Step "Windows activation"
    $slmgr = "$env:WINDIR\System32\slmgr.vbs"
    & cscript //Nologo $slmgr /skms "$KmsHost`:$KmsPort" | Out-Host
    if ($LASTEXITCODE -ne 0) { Bad "slmgr /skms failed"; return }
    $lic = Get-WindowsLicense
    if ($lic -and $lic.Licensed) { OK "Windows already licensed ($($lic.DaysLeft) days) - host pinned, skipping /ato"; return }
    if ($null -eq $lic) {
        Step "No Volume License key - fetching the GVLK for this edition"
        $gvlk = Resolve-WindowsGvlk
        if (-not $gvlk) { Upgrade-WindowsEdition; return }
        Write-Host "    installing GVLK $gvlk"
        & cscript //Nologo $slmgr /ipk $gvlk | Out-Host
        if ($LASTEXITCODE -ne 0) { Bad "slmgr /ipk failed"; return }
    }
    & cscript //Nologo $slmgr /ato | Out-Host
    $lic = Get-WindowsLicense
    if ($lic -and $lic.Licensed) { OK "Windows licensed ($($lic.DaysLeft) days)" }
    else { Bad "Windows not licensed - likely not a VL edition (Home/retail/OEM reject KMS). See https://kms.viktorbarzin.me/#faq" }
}
if ($doWin) { Activate-Windows }

# --- Office / Project / Visio: activate already-installed ----------------
function Find-Ospp {
    # Covers MSI (Office16/15) and Click-to-Run (\root\Office16) layouts, 64- and 32-bit.
    $candidates = @(
        "${env:ProgramFiles}\Microsoft Office\Office16\ospp.vbs",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\ospp.vbs",
        "${env:ProgramFiles}\Microsoft Office\root\Office16\ospp.vbs",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\ospp.vbs",
        "${env:ProgramFiles}\Microsoft Office\Office15\ospp.vbs",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15\ospp.vbs"
    )
    foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
    return $null
}

# Installed Click-to-Run Volume products, e.g. ProPlus2024Volume, VisioPro2024Volume.
# These IDs match the `product` field in the published key list exactly.
function Get-OfficeReleaseIds {
    $c = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($c -and $c.ProductReleaseIds) {
        return $c.ProductReleaseIds.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -match 'Volume$' }
    }
    return @()
}

# Reinstall Office as a Volume License product via the Office Deployment Tool.
# Heavy (multi-GB download, closes Office). Only invoked after explicit consent.
$script:ODT_URL = 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_19127-20198.exe'
function Reinstall-OfficeVL([string]$product) {
    $tmp = Join-Path $env:TEMP "kms-odt-$(Get-Random)"
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    $odt = Join-Path $tmp 'odt.exe'
    Step "Downloading the Office Deployment Tool"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try { Invoke-WebRequest -UseBasicParsing -Uri $script:ODT_URL -OutFile $odt } catch { Bad "ODT download failed: $($_.Exception.Message)"; return $false }
    Start-Process -FilePath $odt -ArgumentList "/extract:`"$tmp`"", '/quiet' -Wait
    $setup = Join-Path $tmp 'setup.exe'
    if (-not (Test-Path $setup)) { Bad "ODT extraction failed (no setup.exe)."; return $false }
    $channel = if ($product -match '2021') { 'PerpetualVL2021' } elseif ($product -match '2019') { 'PerpetualVL2019' } else { 'PerpetualVL2024' }
    $cfg = Join-Path $tmp 'config.xml'
    @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$channel">
    <Product ID="$product"><Language ID="MatchOS" /></Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@ | Set-Content -Path $cfg -Encoding UTF8
    Step "Running setup.exe /configure (multi-GB download + reinstall; several minutes)"
    Start-Process -FilePath $setup -ArgumentList '/configure', "`"$cfg`"" -Wait
    Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    return $true
}

# Per-product license check. ospp /dstatus lists EVERY installed product, so a
# blanket '---LICENSED---' match treats Office-licensed as Project-licensed too.
# Walk the LICENSE NAME / STATUS blocks and check only THIS family.
function Test-OsppLicensed($ospp, $label) {
    $name = ''
    foreach ($l in ((& cscript //Nologo $ospp /dstatus 2>&1) -split "`r?`n")) {
        if ($l -match 'LICENSE NAME:\s*(.+)') { $name = $matches[1] }
        elseif ($l -match '---LICENSED---') {
            $isProj = $name -match 'Project'; $isVis = $name -match 'Visio'
            if ( ($label -eq 'Project' -and $isProj) -or ($label -eq 'Visio' -and $isVis) -or ($label -eq 'Office' -and -not $isProj -and -not $isVis) ) { return $true }
        }
    }
    return $false
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
    # Idempotent: skip /act when THIS family is already licensed.
    if (Test-OsppLicensed $ospp $label) { OK "$label already licensed - host set, skipping /act"; return }
    # Not licensed: install the matching GVLK for each installed VL product of
    # this family (Office = anything that isn't Project/Visio), fetched from the list.
    # Match this family's installed VL products. NB: avoid `switch ($label)` here:
    # inside a switch, $_ is the switch input (the label), not the pipeline item.
    $rels = Get-OfficeReleaseIds | Where-Object {
        ($label -eq 'Project' -and $_ -match 'Project') -or
        ($label -eq 'Visio'   -and $_ -match 'Visio')   -or
        ($label -eq 'Office'  -and $_ -notmatch 'Project|Visio')
    }
    # No Volume License product of this family installed (e.g. retail Office).
    # Offer a VL reinstall via the ODT (default product per family; override
    # with $env:KMS_OFFICE_PRODUCT). Heavy + closes apps, so consent-gated.
    if (-not $rels) {
        $target = if ($env:KMS_OFFICE_PRODUCT) { $env:KMS_OFFICE_PRODUCT }
        elseif ($label -eq 'Project') { 'ProjectPro2024Volume' }
        elseif ($label -eq 'Visio') { 'VisioPro2024Volume' }
        else { 'ProPlus2024Volume' }
        $text = @"
No Volume License $label is installed, so KMS cannot activate it.
The Office Deployment Tool can reinstall it as Volume License:

    ->  $target

Consequences:
  * Downloads ~3 GB and runs setup.exe /configure
  * CLOSES all running Office apps (Word/Excel/Outlook/...)
  * Several minutes; replaces the current $label install
"@
        if (-not (Approve $text ([bool]$env:KMS_OFFICE_PRODUCT))) { Warn "$label reinstall skipped; no VL $label to activate."; return }
        if (-not (Reinstall-OfficeVL $target)) { return }
        $ospp = Find-Ospp; if (-not $ospp) { Bad "ospp.vbs not found after reinstall."; return }
        $rels = @($target)
    }
    $keys = Get-Keys
    foreach ($rel in $rels) {
        $k = if ($keys) { ($keys.office | Where-Object { $_.product -eq $rel } | Select-Object -First 1).gvlk } else { $null }
        if ($k) { Write-Host "    $rel -> installing GVLK $k"; & cscript //Nologo $ospp /inpkey:$k | Out-Host }
    }
    & cscript //Nologo $ospp /act | Out-Host
    if (Test-OsppLicensed $ospp $label) { OK "$label licensed" } else { Warn "$label status not LICENSED yet (no VL $label SKU installed? See https://kms.viktorbarzin.me/#office)" }
}
if ($doOfficeAct) { Activate-Ospp 'Office'  }
if ($doProjAct)   { Activate-Ospp 'Project' }
if ($doVisioAct)  { Activate-Ospp 'Visio'   }

Write-Host ""
Step "Done."
Write-Host "    Re-run any time to re-check status. KMS licences renew automatically every 7 days."
Write-Host "    Privacy: see https://kms.viktorbarzin.me/#faq"
