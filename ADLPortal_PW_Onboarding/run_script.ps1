Install-Module pwps_dab -Scope CurrentUser -Force
Import-Module pwps_dab -Scope CurrentUser
Set-Location -Path $PSScriptRoot

$env:PW_WWHD_USER = "_ADL_Automation"
$env:PW_WWHD_PASSWORD = "_ADL_Automation"

# Dry-run (safe — no PW changes):
powershell -NoProfile -ExecutionPolicy Bypass -File .\Invoke-PWUserOnboarding.ps1 `
  -Project WWHD -UserName first.last -Email first.last@arcadis.com -DryRun -UpdateGroupsLists