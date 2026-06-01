# activate-office.ps1 - one-shot Office KMS activation.
#
# The common case: activate an installed Volume License Office against the public
# KMS host. Thin wrapper that runs kms-bootstrap.ps1 in Office-only mode (finds
# ospp.vbs, installs the right GVLK if needed, then activates; if no VL Office is
# present it offers to install the latest VL edition). Does NOT touch Project/Visio.
#
# Usage (PowerShell as Administrator):
#   iwr -UseBasicParsing https://kms.viktorbarzin.me/scripts/activate-office.ps1 | iex
#
# Source: https://kms.viktorbarzin.me/scripts/activate-office.ps1
# Licence: MIT, no warranty, KMS activates Volume License SKUs only.

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$env:KMS_AUTO = 'office'
iwr -UseBasicParsing https://kms.viktorbarzin.me/scripts/kms-bootstrap.ps1 | iex
