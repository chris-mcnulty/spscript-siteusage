# Requires PnP.PowerShell
Install-Module PnP.PowerShell -Scope CurrentUser

$pwd = Read-Host "PFX password" -AsSecureString

# Permission set covers: inventory enumeration, library-default labels, site container labels,
# file-level label assignment, label-name resolution under app-only auth, and (optional)
# M365 Group sync for container labels on Teams/Group-connected sites.
Register-PnPEntraIDApp `
  -ApplicationName "SPO-Inventory" `
  -Tenant "<yourtenant.onmicrosoft.com or tenantid>" `
  -OutPath ".\cert" `
  -Store CurrentUser `
  -ValidYears 2 `
  -CertificatePassword $pwd `
  -SharePointApplicationPermissions "Sites.FullControl.All" `
  -GraphApplicationPermissions "Sites.FullControl.All","Files.ReadWrite.All","InformationProtectionPolicy.Read.All","User.Read.All","Group.ReadWrite.All"