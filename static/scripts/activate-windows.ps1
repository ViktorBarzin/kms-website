# activate-windows.ps1 - one-shot Windows KMS activation.
#
# The common case: activate this Windows install against the public KMS host.
# Thin wrapper that runs kms-bootstrap.ps1 in Windows-only mode (auto-detects
# the edition, fetches + installs the right GVLK if needed, then activates).
#
# Usage (PowerShell as Administrator):
#   iwr -UseBasicParsing https://kms.viktorbarzin.me/scripts/activate-windows.ps1 | iex
#
# Source: https://kms.viktorbarzin.me/scripts/activate-windows.ps1
# Licence: MIT, no warranty, KMS activates Volume License SKUs only.

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$env:KMS_AUTO = 'win'
iwr -UseBasicParsing https://kms.viktorbarzin.me/scripts/kms-bootstrap.ps1 | iex
