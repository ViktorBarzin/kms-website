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

# Anonymous, fire-and-forget diagnostics so script failures can be debugged
# server-side (logged to Loki). No hostname / username / product keys. Opt out
# with $env:KMS_NO_TELEMETRY=1; $env:KMS_DIAG_URL overrides. Version baked at build.
$script:RunId = ([guid]::NewGuid().ToString('N')).Substring(0, 12)
function Send-Diag([string]$action, [string]$outcome, [string]$detail = '') {
    if ($env:KMS_NO_TELEMETRY) { return }
    try {
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
        $pt = $null; try { $pt = (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).ProductType } catch {}
        if ($detail.Length -gt 600) { $detail = $detail.Substring(0, 600) }
        $body = @{
            script = 'kms-bootstrap.ps1'; ver = '__KMS_VERSION__'; runid = $script:RunId
            ts = (Get-Date).ToUniversalTime().ToString('o')
            action = $action; outcome = $outcome; detail = $detail
            edition = "$($cv.EditionID)"; build = "$($cv.CurrentBuildNumber)"; producttype = $pt
            locale = (Get-Culture).Name
        } | ConvertTo-Json -Compress
        $url = if ($env:KMS_DIAG_URL) { $env:KMS_DIAG_URL } else { 'https://kms.viktorbarzin.me/diag' }
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -UseBasicParsing -Method POST -Uri $url -Body $body -ContentType 'application/json' -TimeoutSec 3 | Out-Null
    } catch {}
}

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
# Gate ONLY on whether there's a real console. NB: $env:KMS_AUTO just pre-selects
# WHICH products to activate (so the activate-windows/office one-liners can set it)
# -- it must NOT suppress this prompt. With a console -> prompt; headless -> proceed
# only if an explicit env override gave consent, else skip.
function Approve([string]$consequences, [bool]$envConsent) {
    Write-Host ""
    Write-Host $consequences -ForegroundColor Yellow
    $canPrompt = $false
    try { $canPrompt = [Environment]::UserInteractive -and -not [Console]::IsInputRedirected } catch { $canPrompt = $false }
    if ($canPrompt) { return (Ask "Proceed?" $false) }
    if ($envConsent) { Write-Host "    (non-interactive; proceeding on explicit env override)"; return $true }
    Warn "Skipped: this needs confirmation. Re-run in an interactive PowerShell window, or set the override env var (e.g. `$env:KMS_OFFICE_PRODUCT or `$env:KMS_EDITION) to consent non-interactively."
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
    if (-not (Approve $text ([bool]$env:KMS_EDITION))) { Warn "Edition upgrade skipped; $cur cannot KMS-activate as-is."; Send-Diag 'win-edition' 'skipped' $cur; return }
    $changepk = "$env:WINDIR\System32\changepk.exe"
    if (-not (Test-Path $changepk)) { Bad "changepk.exe not found on this OS - cannot upgrade edition automatically."; Send-Diag 'win-edition' 'fail' 'changepk missing'; return }
    Step "Upgrading $cur -> $($t.edition) (changepk.exe /ProductKey ...)"
    & $changepk /ProductKey $($t.gvlk)
    Send-Diag 'win-edition' 'upgrade-started' "$cur -> $($t.edition)"
    OK "Edition upgrade started. REBOOT, then re-run this one-liner to activate $($t.edition)."
}

function Activate-Windows {
    Step "Windows activation"
    $slmgr = "$env:WINDIR\System32\slmgr.vbs"
    & cscript //Nologo $slmgr /skms "$KmsHost`:$KmsPort" | Out-Host
    if ($LASTEXITCODE -ne 0) { Bad "slmgr /skms failed"; Send-Diag 'win' 'fail' 'skms failed'; return }
    $lic = Get-WindowsLicense
    if ($lic -and $lic.Licensed) { OK "Windows already licensed ($($lic.DaysLeft) days) - host pinned, skipping /ato"; Send-Diag 'win' 'already-licensed'; return }
    if ($null -eq $lic) {
        Step "No Volume License key - fetching the GVLK for this edition"
        $gvlk = Resolve-WindowsGvlk
        if (-not $gvlk) { Upgrade-WindowsEdition; return }
        Write-Host "    installing GVLK $gvlk"
        & cscript //Nologo $slmgr /ipk $gvlk | Out-Host
        if ($LASTEXITCODE -ne 0) { Bad "slmgr /ipk failed"; Send-Diag 'win' 'fail' 'ipk failed'; return }
    }
    & cscript //Nologo $slmgr /ato | Out-Host
    $lic = Get-WindowsLicense
    if ($lic -and $lic.Licensed) { OK "Windows licensed ($($lic.DaysLeft) days)"; Send-Diag 'win' 'ok' "$($lic.DaysLeft) days" }
    else { Bad "Windows not licensed - likely not a VL edition (Home/retail/OEM reject KMS). See https://kms.viktorbarzin.me/#faq"; Send-Diag 'win' 'fail' 'ato did not stick (non-VL edition?)' }
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

# ALL installed Click-to-Run products (VL + Retail/M365), unfiltered. Retail/M365
# IDs do NOT end in 'Volume' (e.g. O365ProPlusRetail, ProPlus2021Retail).
function Get-AllOfficeC2R {
    $c = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($c -and $c.ProductReleaseIds) {
        return $c.ProductReleaseIds.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
    return @()
}

# Newest VL product for a family from the published list (data-driven: "latest"
# tracks products.yaml, so adding a future LTSC there makes this follow with no
# script change). Picks the highest year among ProPlus/ProjectPro/VisioPro VL SKUs.
function Get-LatestOfficeProduct([string]$label) {
    $keys = Get-Keys; if (-not $keys) { return $null }
    $pat = if ($label -eq 'Project') { '^ProjectPro\d+Volume$' } elseif ($label -eq 'Visio') { '^VisioPro\d+Volume$' } else { '^ProPlus\d+Volume$' }
    $keys.office | Where-Object { $_.product -match $pat } |
        Sort-Object { [int]([regex]::Match($_.product, '\d+').Value) } -Descending | Select-Object -First 1
}

# Office Deployment Tool plumbing. Self-hosted ODT bootstrapper (Microsoft's
# download.microsoft.com URL is build-numbered, rotates every release and 404s;
# we serve a known-good copy from our own /scripts). $env:KMS_ODT_URL overrides.
# Invoke-Odt runs ONE setup.exe /configure with the given <Configuration> XML and
# is reused for both install (<Add>) and uninstall (<Remove>).
$script:ODT_URL   = $(if ($env:KMS_ODT_URL) { $env:KMS_ODT_URL } else { 'https://kms.viktorbarzin.me/scripts/odt-setup.exe' })
# ODT writes its log here (the config XMLs point <Logging> at it). Read back on
# failure so a non-zero / incomplete install reports the real error code instead
# of a bare "ospp.vbs not found". Cleared at the start of every Invoke-Odt run.
$script:OdtLogDir = Join-Path $env:TEMP 'kms-odt-logs'

# Tail of the newest ODT log - the setup.exe error code lives here on failure.
function Get-OdtLogTail([int]$lines = 20) {
    $log = Get-ChildItem -Path $script:OdtLogDir -Filter *.log -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $log) { return '' }
    return ((Get-Content -LiteralPath $log.FullName -Tail $lines -ErrorAction SilentlyContinue) -join ' | ')
}

# "A reboot is pending" probe. C2R half-completes an install when a prior removal
# left the servicing stack pending - the usual cause of an install run right after
# uninstalling the bundled consumer Office.
function Test-PendingReboot {
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { return $true }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { return $true }
    $sm = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
    return [bool]$sm.PendingFileRenameOperations
}

# Poll until the C2R install actually lands. setup.exe /configure can return before
# the Click-to-Run service finishes the multi-GB install, so trust on-disk state
# (ospp.vbs + the product in ProductReleaseIds), not setup.exe's return. Heartbeat
# so a long download does not look hung.
function Wait-OfficeInstalled([string]$product, [int]$timeoutSec = 1800) {
    $sw = [Diagnostics.Stopwatch]::StartNew(); $next = 0
    while ($sw.Elapsed.TotalSeconds -lt $timeoutSec) {
        if ((Find-Ospp) -and ((Get-AllOfficeC2R) -contains $product)) { return $true }
        if ($sw.Elapsed.TotalSeconds -ge $next) { Write-Host "    still installing $product... ($([int]$sw.Elapsed.TotalSeconds)s)"; $next += 20 }
        Start-Sleep -Seconds 5
    }
    return $false
}

function Invoke-Odt([string]$configXml, [string]$stepMsg) {
    Remove-Item -Path (Join-Path $script:OdtLogDir '*') -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force -Path $script:OdtLogDir | Out-Null
    $tmp = Join-Path $env:TEMP "kms-odt-$(Get-Random)"
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    try {
        $odt = Join-Path $tmp 'odt.exe'
        Step "Downloading the Office Deployment Tool"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        try { Invoke-WebRequest -UseBasicParsing -Uri $script:ODT_URL -OutFile $odt } catch { Bad "ODT download failed: $($_.Exception.Message)"; Send-Diag 'odt' 'fail' "download: $($_.Exception.Message)"; return $false }
        Start-Process -FilePath $odt -ArgumentList "/extract:`"$tmp`"", '/quiet' -Wait
        $setup = Join-Path $tmp 'setup.exe'
        if (-not (Test-Path $setup)) { Bad "ODT extraction failed (no setup.exe)."; Send-Diag 'odt' 'fail' 'extract: no setup.exe'; return $false }
        $cfg = Join-Path $tmp 'config.xml'
        $configXml | Set-Content -Path $cfg -Encoding UTF8
        Step $stepMsg
        $p = Start-Process -FilePath $setup -ArgumentList '/configure', "`"$cfg`"" -Wait -PassThru
        $code = $p.ExitCode
        # 0 = success, 3010 = success/reboot-required. Anything else is a real ODT
        # failure - surface the log tail (carries the error code) so we can diagnose.
        if ($code -ne 0 -and $code -ne 3010) {
            $tail = Get-OdtLogTail
            Bad "Office Deployment Tool failed (setup.exe exit $code)."
            if ($tail) { Write-Host "    ODT log: $tail" }
            Send-Diag 'odt' 'fail' "exit=$code; $tail"
            return $false
        }
        if ($code -eq 3010) { Warn "ODT reports a reboot is required to finish." }
        return $true
    }
    finally { Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue }
}

function Reinstall-OfficeVL([string]$product, [string]$channel) {
    if (-not $channel) { $channel = if ($product -match '2021') { 'PerpetualVL2021' } elseif ($product -match '2019') { 'PerpetualVL2019' } else { 'PerpetualVL2024' } }
    $xml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$channel">
    <Product ID="$product"><Language ID="MatchOS" /></Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Logging Level="Standard" Path="$($script:OdtLogDir)" />
</Configuration>
"@
    if (-not (Invoke-Odt $xml "Installing $product (multi-GB download + reinstall; several minutes)")) { return $false }
    # setup.exe returning is NOT proof Office is on disk - verify it actually landed.
    if (Wait-OfficeInstalled $product) { OK "$product installed"; return $true }
    $reboot = Test-PendingReboot
    $hint = if ($reboot) { 'A reboot is pending - reboot, then re-run the one-liner (it skips straight to the install).' } else { 'Re-run the one-liner; if it still fails, reboot first to clear the servicing stack.' }
    Bad "$product install did not complete (ospp.vbs / $product not registered). $hint"
    $tail = Get-OdtLogTail; if ($tail) { Write-Host "    ODT log: $tail" }
    Send-Diag 'odt' 'install-incomplete' "product=$product; reboot=$reboot; $tail"
    return $false
}

# Uninstall specific Click-to-Run products (the incompatible retail/M365 Office
# that blocks a VL install). Only the listed products are removed; VL products of
# other families are preserved.
function Uninstall-OfficeC2R([string[]]$products) {
    $items = ($products | ForEach-Object { "    <Product ID=`"$_`" />" }) -join "`n"
    $xml = @"
<Configuration>
  <Remove>
$items
  </Remove>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Logging Level="Standard" Path="$($script:OdtLogDir)" />
</Configuration>
"@
    return (Invoke-Odt $xml "Uninstalling incompatible Office: $($products -join ', ')")
}

# Offer to install the LATEST Volume License product for a family (or the
# explicit $env:KMS_OFFICE_PRODUCT) via the ODT, after showing the consequences.
# Returns the installed product id, or $null if skipped/failed.
function Install-LatestOfficeVL([string]$label) {
    $best = $null; $target = $env:KMS_OFFICE_PRODUCT; $channel = $null
    if (-not $target) { $best = Get-LatestOfficeProduct $label; if ($best) { $target = $best.product; $channel = $best.channel } }
    if (-not $target) { Warn "Couldn't determine the latest VL $label product from the key list."; return $null }
    $disp = if ($best) { "$($best.edition) ($target)" } else { $target }
    # Retail/M365 Click-to-Run Office (IDs not ending in 'Volume') can't coexist
    # with a VL install of the same suite and must be removed first.
    $incompat = @(Get-AllOfficeC2R | Where-Object { $_ -notmatch 'Volume$' })
    $rm = if ($incompat) { "`n  * FIRST UNINSTALLS the incompatible (non-VL) Office found: $($incompat -join ', ')" } else { '' }
    $text = @"
KMS can only activate Volume License $label, and none is installed here.
The Office Deployment Tool can install the LATEST Volume License edition:

    ->  $disp

Consequences:$rm
  * Downloads ~3 GB and runs setup.exe /configure
  * CLOSES all running Office apps (Word/Excel/Outlook/...)
  * Several minutes
"@
    if (-not (Approve $text ([bool]$env:KMS_OFFICE_PRODUCT))) { Warn "$label install skipped."; Send-Diag $label.ToLower() 'install-skipped' $target; return $null }
    if ($incompat) {
        if (-not (Uninstall-OfficeC2R $incompat)) { Bad "Uninstall of incompatible Office failed; aborting before VL install."; Send-Diag $label.ToLower() 'uninstall-fail' ($incompat -join ','); return $null }
        Send-Diag $label.ToLower() 'uninstalled-incompatible' ($incompat -join ',')
        # Removing the bundled consumer Office usually leaves the servicing stack
        # pending a reboot; a VL install in the SAME session then half-completes with
        # no ospp.vbs. Stop here and have the user reboot + re-run - the script is
        # idempotent and (no incompatible Office left) goes straight to the install.
        if (Test-PendingReboot) {
            Warn "Incompatible Office removed, but a reboot is required before the Volume License install."
            Write-Host "    Please REBOOT, then re-run the one-liner - it will skip the uninstall and install $target directly."
            Send-Diag $label.ToLower() 'reboot-required-after-uninstall' $target
            return $null
        }
    }
    if (-not (Reinstall-OfficeVL $target $channel)) { return $null }
    return $target
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
    # No Office found at all -> offer to install the latest VL edition.
    if (-not $ospp) {
        Step "$label`: no Office found - offering the latest Volume License install"
        $t = Install-LatestOfficeVL $label
        if (-not $t) { return }
        $ospp = Find-Ospp; if (-not $ospp) { Bad "$label`: ospp.vbs not found after install."; return }
    }
    Step "$label activation via $ospp"
    & cscript //Nologo $ospp /sethst:$KmsHost | Out-Host
    & cscript //Nologo $ospp /setprt:$KmsPort | Out-Host
    # Idempotent: skip /act when THIS family is already licensed.
    if (Test-OsppLicensed $ospp $label) { OK "$label already licensed - host set, skipping /act"; Send-Diag $label.ToLower() 'already-licensed'; return }
    # Installed VL products of this family. NB: avoid `switch ($label)` here -
    # inside a switch, $_ is the switch input (the label), not the pipeline item.
    $rels = Get-OfficeReleaseIds | Where-Object {
        ($label -eq 'Project' -and $_ -match 'Project') -or
        ($label -eq 'Visio'   -and $_ -match 'Visio')   -or
        ($label -eq 'Office'  -and $_ -notmatch 'Project|Visio')
    }
    # No VL product of this family (e.g. retail/M365 Office) -> offer latest VL.
    if (-not $rels) {
        $t = Install-LatestOfficeVL $label
        if (-not $t) { return }
        $ospp = Find-Ospp; if (-not $ospp) { Bad "ospp.vbs not found after install."; return }
        $rels = @($t)
    }
    $keys = Get-Keys
    foreach ($rel in $rels) {
        $k = if ($keys) { ($keys.office | Where-Object { $_.product -eq $rel } | Select-Object -First 1).gvlk } else { $null }
        if ($k) { Write-Host "    $rel -> installing GVLK $k"; & cscript //Nologo $ospp /inpkey:$k | Out-Host }
    }
    & cscript //Nologo $ospp /act | Out-Host
    if (Test-OsppLicensed $ospp $label) { OK "$label licensed"; Send-Diag $label.ToLower() 'ok' ($rels -join ',') }
    else { Warn "$label status not LICENSED yet (no VL $label SKU? See https://kms.viktorbarzin.me/#office)"; Send-Diag $label.ToLower() 'fail' ("products=" + ($rels -join ',')) }
}
if ($doOfficeAct) { Activate-Ospp 'Office'  }
if ($doProjAct)   { Activate-Ospp 'Project' }
if ($doVisioAct)  { Activate-Ospp 'Visio'   }

Write-Host ""
Step "Done."
Write-Host "    Re-run any time to re-check status. KMS licences renew automatically every 7 days."
Write-Host "    Privacy: see https://kms.viktorbarzin.me/#faq"
