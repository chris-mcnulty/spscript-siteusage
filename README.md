# spscripts
SharePoint and M365 Scripts

## Overview
This repository contains PowerShell scripts for managing and reporting on SharePoint Online and Microsoft 365 environments.

## Scripts

### Get-SPOSiteUsageReports.ps1
Enumerate usage reports from all SharePoint sites in a tenant and export to CSV.

**Quick Start:**
```powershell
.\Get-SPOSiteUsageReports.ps1 -TenantName "contoso"
```

**Features:**
- Collects comprehensive usage statistics for all SharePoint sites
- Supports both SharePoint Online Management Shell and Microsoft Graph API
- Automatic module installation
- Exports to CSV with customizable output path
- Progress tracking and error handling

**[Full Documentation](README-Get-SPOSiteUsageReports.md)**

---

### Sensitivity label inventory & apply (`Inventory/`)

A PnP.PowerShell-based toolkit for inventorying and applying Microsoft Purview
sensitivity labels across an M365 tenant. All four scripts share the same
Entra app + certificate created by the setup script.

| Script | Purpose |
|---|---|
| `Inventory/Setup-SPOInventoryApp.ps1` | One-time: create Entra app, generate + upload certificate, grant `Sites.FullControl.All`. |
| `Inventory/Export-M365SensitivityLabelInventory-PnPOnly.ps1` | Read-only inventory of site + library sensitivity labels (CSV). |
| `Inventory/Set-M365LibraryDefaultSensitivityLabel-PnPOnly.ps1` | Set `DefaultSensitivityLabelForLibrary` on every document library. |
| `Inventory/Apply-M365FileSensitivityLabel-PnPOnly.ps1` | Apply a sensitivity label to existing files in every library. |

**Quick start — set a default sensitivity label on every library:**
```powershell
$pwd = Read-Host "PFX password" -AsSecureString

.\Inventory\Set-M365LibraryDefaultSensitivityLabel-PnPOnly.ps1 `
  -TenantName "contoso" `
  -ClientId "<app-client-id>" `
  -Tenant "contoso.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -DefaultLabelId "<label-guid>" `
  -WhatIf
```

**Quick start — apply each library's default label to existing files:**
```powershell
.\Inventory\Apply-M365FileSensitivityLabel-PnPOnly.ps1 `
  -TenantName "contoso" `
  -ClientId "<app-client-id>" `
  -Tenant "contoso.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -WhatIf
```

Both apply scripts support `-WhatIf` for dry runs, `-Resume` for interruption
recovery, `-OverwriteExisting` to replace different labels, and
`-SiteUrlLike` to scope to a subset of sites. The file apply script also
takes `-LabelId` to force one label tenant-wide and `-MaxFilesPerLibrary` to
cap a rollout.

**[Full Documentation](Inventory/README.md)**

## Requirements
- PowerShell 5.1 or later for `Get-SPOSiteUsageReports.ps1`
- **PowerShell 7.4+ required** for the `Inventory/` PnP-based scripts
- Appropriate SharePoint/Microsoft 365 admin permissions
- Internet connection for module installation

## Getting Started

1. Clone this repository:
```powershell
git clone https://github.com/chris-mcnulty/spscripts.git
cd spscripts
```

2. Run the desired script with appropriate parameters

3. Required PowerShell modules will be automatically installed if missing

## Contributing
Contributions are welcome! Please feel free to submit pull requests or open issues.

## License
This project is provided as-is without warranty.
