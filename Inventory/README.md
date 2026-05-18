# 🔐 SharePoint Sensitivity Label Inventory (PnP.PowerShell)

This repo contains two scripts:

1. **[Setup-SPOInventoryApp.ps1](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=187&web=1&EntityRepresentationId=b0ad99cd-e6dd-477e-9b4a-2685f46239df)** — one-time setup to create the Entra app + cert and grant baseline permissions. 【1-227cf3】  
2. **[Export-M365SensitivityLabelInventory-PnPOnly.ps1](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1&EntityRepresentationId=58464a48-71e7-483b-9c2d-1a3bca122d82)** — the actual tenant inventory run (PnP-only, with progress output and resume). 【2-547b7b】  

The inventory script enumerates SharePoint sites, then enumerates document libraries per site, and exports a CSV with site/library sensitivity label signals. 【2-547b7b】  

---

## Table of contents

- #what-the-setup-script-does
- #prerequisites
- #one-time-setup-run-the-setup-script
- #required-entra-permissions
- #enable-sensitivity-labels-for-sharepointonedrive-required
- #enable-pdf-support-recommended
- #run-the-inventory-script-pnp-version
- #parameters-inventory-script
- #outputs
- #runtime-signs-of-life
- [Troubleshooting](#troubleshooting)

---

## What the setup script does

**[Setup-SPOInventoryApp.ps1](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=187&web=1&EntityRepresentationId=b0ad99cd-e6dd-477e-9b4a-2685f46239df)** automates the full app + certificate bootstrap so you do **not** have to manually create an app or upload a `.cer` in the portal. 【1-227cf3】  

Specifically, it: 【1-227cf3】  

- Creates an **Entra app registration** and **service principal**. 【1-227cf3】  
- Generates a self-signed certificate (**PFX + CER**) and exports both. 【1-227cf3】  
- Uploads the certificate to the app registration using **keyCredentials** (so you don’t upload the `.cer` manually). 【1-227cf3】  
- Grants SharePoint application permission **Sites.FullControl.All** to the new service principal. 【1-227cf3】  
- Outputs the exact values you’ll paste into the inventory run: **ClientId**, **TenantId**, **CertificatePath**, **CertificateCer**. 【1-227cf3】  

> Optional: the script includes logic to assign the Entra role **SharePoint Administrator** (tenant scope) if you enable that option in the script. 【1-227cf3】  

---

## Prerequisites

### PowerShell + modules

- **PowerShell 7.4+ is required** for PnP.PowerShell. 【3-d1f9ad】  
- Install/upgrade PnP.PowerShell: 【3-d1f9ad】  

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser -Force -SkipPublisherCheck
Import-Module PnP.PowerShell
````

### Permissions to run setup

The setup script connects to Microsoft Graph and creates/updates app registrations and app-role assignments.   
You’ll need an admin identity that can consent and create apps in your tenant. [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=187&web=1)

***

## One-time setup (run the setup script)

Run **[Setup-SPOInventoryApp.ps1](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=187\&web=1\&EntityRepresentationId=b0ad99cd-e6dd-477e-9b4a-2685f46239df)** once to create the app + cert and grant baseline SharePoint permissions. [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=187&web=1)

Example pattern (your parameters may vary based on your local copy of the script): [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=187&web=1)

```powershell
# Example: run the setup script (edit parameters in the script if needed)
.\Setup-SPOInventoryApp.ps1
```

When it completes, it prints a “copy these values” block including: [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=187&web=1)

* `ClientId`
* `TenantId`
* `CertificatePath` (PFX path)
* `CertificateCer` (CER path)

> ✅ You do **not** need to upload the `.cer` manually if you used the setup script — it already uploads the cert to Entra. [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=187&web=1)

***

## Required Entra permissions

### 1) SharePoint permission (set by setup script)

The setup script grants the app **Sites.FullControl.All (Application)** for SharePoint Online. [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=187&web=1)

### 2) Microsoft Graph permission (needed to resolve label names)

If you want **label names** (not just label GUIDs), you must grant Graph permissions required by `Get-PnPAvailableSensitivityLabel`. [\[pnp.github.io\]](https://pnp.github.io/powershell/cmdlets/Get-PnPAvailableSensitivityLabel.html), [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)

PnP documents the required Graph permissions as: [\[pnp.github.io\]](https://pnp.github.io/powershell/cmdlets/Get-PnPAvailableSensitivityLabel.html)

* **Delegated:** `InformationProtectionPolicy.Read`
* **Application:** `InformationProtectionPolicy.Read.All`

Because this inventory runs **app-only** (certificate auth), you want **Application** permission: `InformationProtectionPolicy.Read.All`. [\[pnp.github.io\]](https://pnp.github.io/powershell/cmdlets/Get-PnPAvailableSensitivityLabel.html), [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1), [\[graphpermi...merill.net\]](https://graphpermissions.merill.net/permission/InformationProtectionPolicy.Read.All)

> If you don’t grant this, the script will still export label IDs but will warn: “Label name mapping unavailable; exporting label IDs only.” [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)

***

## Enable sensitivity labels for SharePoint/OneDrive (required)

Microsoft Purview requires an explicit enablement step so SharePoint/OneDrive can apply and process sensitivity labels on files (including encrypted files). [\[learn.microsoft.com\]](https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-onedrive-files)

From Microsoft guidance, enabling sensitivity labels for SharePoint and OneDrive: [\[learn.microsoft.com\]](https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-onedrive-files)

* enables built-in labeling for supported Office files and **PDF files**
* lets users apply labels in Office for the web and in SharePoint/OneDrive UI
* allows SharePoint/OneDrive to **process** encrypted labeled content (eDiscovery, search, coauthoring, etc.) once enabled [\[learn.microsoft.com\]](https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-onedrive-files)

**Do this before relying on inventory results for governance decisions.** [\[learn.microsoft.com\]](https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-onedrive-files)

***

## Enable PDF support (recommended)

If you want correct coverage for labeled PDFs (and especially for labeling/encryption scenarios), enable PDF label support.

### Option A — Enable in Purview (UI)

Microsoft describes enabling built-in labeling for supported Office files and **PDF files** in SharePoint/OneDrive as part of the Purview feature enablement. [\[learn.microsoft.com\]](https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-onedrive-files)

### Option B — Enable via PowerShell (tenant setting)

Microsoft’s message center guidance notes PDF labeling can be enabled via: [\[m365admin....sontek.net\]](https://m365admin.handsontek.net/microsoft-purview-information-protection-auto-labeling-for-files-at-rest-in-sharepoint-online-can-now-label-pdf-files/)

```powershell
Set-SPOTenant -EnableSensitivityLabelforPDF $true
```

> Note: the message center guidance indicates this feature can be off by default and calls out a short propagation window after enabling. [\[m365admin....sontek.net\]](https://m365admin.handsontek.net/microsoft-purview-information-protection-auto-labeling-for-files-at-rest-in-sharepoint-online-can-now-label-pdf-files/)

***

## Run the inventory script (PnP version)

Use **[Export-M365SensitivityLabelInventory-PnPOnly.ps1](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216\&web=1\&EntityRepresentationId=58464a48-71e7-483b-9c2d-1a3bca122d82)** for the PnP-only inventory run. [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)

**Cut/paste example:**

```powershell
$pwd = Read-Host "PFX password" -AsSecureString

.\Export-M365SensitivityLabelInventory-PnPOnly.ps1 `
  -TenantName "synozur" `
  -ClientId "df7b6a64-8d68-4e75-b15d-2cffc07cb554" `
  -Tenant "synozur.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -Resume
```

The script connects to your admin site using `Connect-PnPOnline` with `-ClientId`, `-Tenant`, `-CertificatePath`, and `-CertificatePassword`. [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)

***

## Parameters (inventory script)

The inventory script uses these parameters internally: [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)

### Required (app-only auth)

* `TenantName` — used to build `https://<TenantName>-admin.sharepoint.com` [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)
* `ClientId` — Entra app (Application/Client ID) used for app-only auth [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)
* `Tenant` — tenant identifier used by `Connect-PnPOnline` [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)
* `CertificatePath` — PFX file path used by `Connect-PnPOnline` [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)
* `CertificatePassword` — SecureString password for the PFX [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)

### Optional behavior switches

* `IncludeOneDriveSites` — passed to `Get-PnPTenantSite -IncludeOneDriveSites:$IncludeOneDriveSites` [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)
* `Resume` — if set, reads existing CSV and skips already-written rows [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)

> Repo default behavior: the script is designed so you can exclude OneDrive sites by default by setting `IncludeOneDriveSites` false in your script configuration. [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)

***

## Outputs

The script writes a CSV containing these columns: [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)

* `Timestamp`
* `Scope` (site vs library rows)
* `SiteUrl`
* `SiteSensitivityLabelId`
* `SiteSensitivityLabelName`
* `LibraryTitle`
* `LibraryId`
* `LibraryServerRelativeUrl`
* `LibraryDefaultLabelId`
* `LibraryDefaultLabelName` [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)

> Note: If label name mapping cannot be retrieved, label “Name” fields will be empty and the script will warn accordingly. [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1), [\[pnp.github.io\]](https://pnp.github.io/powershell/cmdlets/Get-PnPAvailableSensitivityLabel.html)

***

## Runtime “signs of life”

The script prints:

* “Connecting…” to the admin URL [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)
* “Retrieving sites…” and “Sites found…” [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1)

If you are using a version with per-site/per-library progress output (recommended), you’ll see rolling progress as each site and library is processed.

***

## Troubleshooting

### “Label name mapping unavailable; exporting label IDs only.”

This means `Get-PnPAvailableSensitivityLabel` couldn’t run successfully.   
Grant **Graph Application** permission `InformationProtectionPolicy.Read.All` as documented for that cmdlet. [\[synozur.sh...epoint.com\]](https://synozur.sharepoint.com/sites/SynozurIT/Shared%20Documents/Forms/DispForm.aspx?ID=216&web=1) [\[pnp.github.io\]](https://pnp.github.io/powershell/cmdlets/Get-PnPAvailableSensitivityLabel.html), [\[graphpermi...merill.net\]](https://graphpermissions.merill.net/permission/InformationProtectionPolicy.Read.All)

### PDF label coverage looks wrong

Ensure sensitivity labels for SharePoint/OneDrive are enabled, and optionally enable PDF support. [\[learn.microsoft.com\]](https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-onedrive-files), [\[m365admin....sontek.net\]](https://m365admin.handsontek.net/microsoft-purview-information-protection-auto-labeling-for-files-at-rest-in-sharepoint-online-can-now-label-pdf-files/)

### Archived / restricted sites

Some archived/restricted sites may still appear in tenant enumeration but block list enumeration (403 Forbidden). Your inventory run should treat these as “skipped” sites, not hard failures (recommended behavior for governance inventories).

***

## References

* [Enable sensitivity labels for files in SharePoint and OneDrive](https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-onedrive-files)
* [Auto-labeling for PDFs + Set-SPOTenant -EnableSensitivityLabelforPDF](https://m365admin.handsontek.net/microsoft-purview-information-protection-auto-labeling-for-files-at-rest-in-sharepoint-online-can-now-label-pdf-files/)
* [PnP.PowerShell installation](https://deepwiki.com/pnp/powershell/4.1-site-and-tenant-administration)
* [Get-PnPAvailableSensitivityLabel (required Graph permissions)](https://www.linkedin.com/pulse/securing-ms-copilot-access-over-sharepoint-site-content-sher-azam-fcy4f)

***

```
