# Requires PnP.PowerShell
Install-Module PnP.PowerShell -Scope CurrentUser

$pwd = Read-Host "PFX password" -AsSecureString

Register-PnPEntraIDApp `
  -ApplicationName "SPO-Inventory" `
  -Tenant "<yourtenant.onmicrosoft.com or tenantid>" `
  -OutPath ".\cert" `
  -Store CurrentUser `
  -ValidYears 2 `
  -CertificatePassword $pwd `
  -SharePointApplicationPermissions Sites.FullControl.All