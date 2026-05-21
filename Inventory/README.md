# SharePoint Sensitivity Label Inventory & Apply (PnP.PowerShell)

This folder contains four main scripts that together let you stand up an app
registration, inventory sensitivity labels across the tenant, and apply
default labels at the library level and explicit labels at the file level.

1. **Setup-SPOInventoryApp.ps1** — one-time setup. Creates the Entra app +
   certificate and grants baseline SharePoint permissions.
2. **Export-M365SensitivityLabelInventory-PnPOnly.ps1** — read-only inventory.
   Enumerates sites and document libraries and exports a CSV with the
   sensitivity-label signals on each.
3. **Set-M365LibraryDefaultSensitivityLabel-PnPOnly.ps1** — write. Sets a
   `DefaultSensitivityLabelForLibrary` on every document library in the
   tenant (with `-WhatIf`, `-Resume`, and `-OverwriteExisting` support).
4. **Apply-M365FileSensitivityLabel-PnPOnly.ps1** — write. Applies a
   sensitivity label to existing files in every document library, either
   the library's own default label or a single forced `-LabelId`.

A legacy `Export-M365SensitivityLabelInventory.ps1` (SPO + PnP hybrid) and an
`exampleAIP.ps1` snippet (tenant `EnableAIPIntegration` toggle) are also kept
here. Prefer the `-PnPOnly` variants for new work.

---

## Table of contents

- [What the setup script does](#what-the-setup-script-does)
- [Prerequisites](#prerequisites)
- [One-time setup (run the setup script)](#one-time-setup-run-the-setup-script)
- [Required Entra permissions](#required-entra-permissions)
- [Enable sensitivity labels for SharePoint/OneDrive (required)](#enable-sensitivity-labels-for-sharepointonedrive-required)
- [Enable PDF support (recommended)](#enable-pdf-support-recommended)
- [Run the inventory script (PnP version)](#run-the-inventory-script-pnp-version)
- [Run the library default-label apply script](#run-the-library-default-label-apply-script)
- [Run the file label apply script](#run-the-file-label-apply-script)
- [Outputs](#outputs)
- [Runtime "signs of life"](#runtime-signs-of-life)
- [Troubleshooting](#troubleshooting)

---

## What the setup script does

**Setup-SPOInventoryApp.ps1** automates the full app + certificate bootstrap
so you do **not** have to manually create an app or upload a `.cer` in the
portal.

Specifically, it:

- Creates an **Entra app registration** and **service principal**.
- Generates a self-signed certificate (**PFX + CER**) and exports both.
- Uploads the certificate to the app registration via **keyCredentials** (so
  you don't upload the `.cer` manually).
- Grants the full application-permission set on the service principal:
  SharePoint `Sites.FullControl.All`, Graph `Sites.FullControl.All`,
  `Files.ReadWrite.All`, `InformationProtectionPolicy.Read.All`,
  `User.Read.All`, and `Group.ReadWrite.All` by default. In
  `Setup-SPOInventoryApp.ps1`, `Group.ReadWrite.All` is omitted only when you
  pass `-SkipGroupReadWrite`.
- Outputs the exact values you'll paste into the inventory and apply runs:
  **ClientId**, **TenantId**, **CertificatePath**, **CertificateCer**.

> Optional: the script includes logic to assign the Entra role **SharePoint
> Administrator** (tenant scope) if you enable that option in the script.

---

## Prerequisites

### PowerShell + modules

- **PowerShell 7.4+ is required** for PnP.PowerShell.
- Install/upgrade PnP.PowerShell:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser -Force -SkipPublisherCheck
Import-Module PnP.PowerShell
```

### Permissions to run setup

The setup script connects to Microsoft Graph and creates/updates app
registrations and app-role assignments. You'll need an admin identity that
can consent and create apps in your tenant.

---

## One-time setup (run the setup script)

Run **Setup-SPOInventoryApp.ps1** once to create the app + cert and grant
baseline SharePoint permissions.

```powershell
.\Setup-SPOInventoryApp.ps1 `
  -AppName "SPO-Inventory" `
  -TenantId "<your-tenant-guid>" `
  -OutDir ".\cert"
```

When it completes, it prints a "copy these values" block including:

- `ClientId`
- `TenantId`
- `CertificatePath` (PFX path)
- `CertificateCer` (CER path)

> You do **not** need to upload the `.cer` manually if you used the setup
> script — it already uploads the cert to Entra.

---

## Required Entra permissions

Both `Setup-SPOInventoryApp.ps1` and `Setup-SPOInventoryQuick.ps1` now grant
the full permission set below in a single run. Re-running the setup script
against an existing app is idempotent (already-granted roles are skipped).

### 1) SharePoint Online (Application)

| Permission | Used for |
|---|---|
| `Sites.FullControl.All` | tenant site enumeration, per-site connect, `Set-PnPList -DefaultSensitivityLabelForLibrary` |

### 2) Microsoft Graph (Application)

| Permission | Used for |
|---|---|
| `Sites.FullControl.All` | `Set-PnPSite -SensitivityLabel` (site container label) |
| `Files.ReadWrite.All` | file-level `POST /drives/{id}/items/{itemId}/assignSensitivityLabel` |
| `InformationProtectionPolicy.Read.All` | `Get-PnPAvailableSensitivityLabel` — resolve label GUIDs to display names |
| `User.Read.All` | scope label-policy lookup to a UPN under app-only auth (`-LabelOwnerUpn`) |
| `Group.ReadWrite.All` | sync container label to backing M365 Group for Teams/Group-connected sites (omit with `-SkipGroupReadWrite` on the setup script) |

Without `InformationProtectionPolicy.Read.All` + `User.Read.All`, the apply
scripts will warn "Label name mapping unavailable; logging GUIDs only" and
all label-name columns will contain GUIDs only. Without `Files.ReadWrite.All`
the file-level apply script will 403. Without Graph `Sites.FullControl.All`
the site container label cannot be set (libraries can still be set via the
SharePoint permission).

---

## Enable sensitivity labels for SharePoint/OneDrive (required)

Microsoft Purview requires an explicit enablement step so SharePoint/OneDrive
can apply and process sensitivity labels on files (including encrypted
files). Enabling sensitivity labels for SharePoint and OneDrive:

- enables built-in labeling for supported Office files and **PDF files**
- lets users apply labels in Office for the web and in SharePoint/OneDrive UI
- allows SharePoint/OneDrive to **process** encrypted labeled content
  (eDiscovery, search, coauthoring, etc.) once enabled

**Do this before relying on inventory results — or running the apply
scripts — for governance decisions.**

---

## Enable PDF support (recommended)

If you want correct coverage for labeled PDFs (and especially for
labeling/encryption scenarios), enable PDF label support.

### Option A — Enable in Purview (UI)

Microsoft describes enabling built-in labeling for supported Office files and
**PDF files** in SharePoint/OneDrive as part of the Purview feature
enablement.

### Option B — Enable via PowerShell (tenant setting)

```powershell
Set-SPOTenant -EnableSensitivityLabelforPDF $true
```

> Note: this feature can be off by default and there is a short propagation
> window after enabling.

---

## Run the inventory script (PnP version)

Use **Export-M365SensitivityLabelInventory-PnPOnly.ps1** for the PnP-only
inventory run.

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

The script connects to your admin site using `Connect-PnPOnline` with
`-ClientId`, `-Tenant`, `-CertificatePath`, and `-CertificatePassword`.

### Parameters

Required (app-only auth):

- `TenantName` — used to build `https://<TenantName>-admin.sharepoint.com`
- `ClientId` — Entra app (Application/Client ID) used for app-only auth
- `Tenant` — tenant identifier used by `Connect-PnPOnline`
- `CertificatePath` — PFX file path used by `Connect-PnPOnline`
- `CertificatePassword` — SecureString password for the PFX

Optional:

- `IncludeOneDriveSites` — passed through to `Get-PnPTenantSite`
- `Resume` — read existing CSV and skip already-written rows

---

## Run the library default-label apply script

Use **Set-M365LibraryDefaultSensitivityLabel-PnPOnly.ps1** to set
`DefaultSensitivityLabelForLibrary` on every document library across the
tenant. This is the library-level default that controls what label new files
inherit; it does **not** label existing files (use the file-label script for
that).

### Dry run first (recommended)

```powershell
$pwd = Read-Host "PFX password" -AsSecureString

.\Set-M365LibraryDefaultSensitivityLabel-PnPOnly.ps1 `
  -TenantName "contoso" `
  -ClientId "<app-client-id>" `
  -Tenant "contoso.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -DefaultLabelId "00000000-0000-0000-0000-000000000000" `
  -WhatIf
```

`-WhatIf` writes audit rows with `Action=WouldSet` so you can review which
libraries would be touched.

### Apply — only libraries that currently have no default label

```powershell
.\Set-M365LibraryDefaultSensitivityLabel-PnPOnly.ps1 `
  -TenantName "synozur" `
  -ClientId "df7b6a64-8d68-4e75-b15d-2cffc07cb554" `
  -Tenant "synozur.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -DefaultLabelId "00000000-0000-0000-0000-000000000000" `
  -Resume
```

Libraries that already have a *different* label are logged with
`Action=SkippedExistingLabel` and left alone.

### Apply — replace existing different labels too

```powershell
.\Set-M365LibraryDefaultSensitivityLabel-PnPOnly.ps1 `
  -TenantName "synozur" `
  -ClientId "df7b6a64-8d68-4e75-b15d-2cffc07cb554" `
  -Tenant "synozur.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -DefaultLabelId "00000000-0000-0000-0000-000000000000" `
  -OverwriteExisting `
  -Resume
```

### Scope to a single set of sites

```powershell
.\Set-M365LibraryDefaultSensitivityLabel-PnPOnly.ps1 `
  -TenantName "synozur" `
  -ClientId "df7b6a64-8d68-4e75-b15d-2cffc07cb554" `
  -Tenant "synozur.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -DefaultLabelId "00000000-0000-0000-0000-000000000000" `
  -SiteUrlLike "*/sites/Finance*"
```

### Parameters

Required:

- `TenantName`, `ClientId`, `Tenant`, `CertificatePath`, `CertificatePassword` — same as inventory
- `DefaultLabelId` — sensitivity label GUID to set as each library's default

Optional:

- `OverwriteExisting` — replace a library's existing different default label
- `Resume` — skip libraries already written as `Set`/`AlreadySet`
- `SiteUrlLike` — wildcard filter on site URL (e.g. `*/sites/Finance*`)
- `IncludeOneDriveSites` — include personal OneDrive sites (default `$false`)
- `IncludeHiddenLibraries` — include hidden lists (default `$false`)
- `OutputCsvPath` / `ErrorCsvPath` — override CSV locations
- `-WhatIf` — dry run (writes `Action=WouldSet` rows)

### Actions written to the audit CSV

| Action | Meaning |
|---|---|
| `Set` | Default label was applied. |
| `AlreadySet` | Library already had the target label. |
| `SkippedExistingLabel` | Library has a different label; re-run with `-OverwriteExisting` to replace. |
| `WouldSet` | `-WhatIf` dry-run row. |
| `Error` | Set call failed; details in the errors CSV. |

Only `Set` and `AlreadySet` are treated as terminal for `-Resume` —
`WouldSet` and `SkippedExistingLabel` rows are re-evaluated on the next run.

---

## Run the file label apply script

Use **Apply-M365FileSensitivityLabel-PnPOnly.ps1** to label existing files.
By default each file inherits its library's current
`DefaultSensitivityLabelForLibrary`; pass `-LabelId` to force one label
across every library instead.

> Microsoft's recommended approach for backfilling labels at scale is a
> Purview auto-labeling policy. Use this script when that is not viable
> (small tenants, targeted libraries, incremental rollout).

Labels are applied via the Graph endpoint
`POST /drives/{drive-id}/items/{item-id}/assignSensitivityLabel`, so the
same app registration must have `Files.ReadWrite.All`, `Sites.Read.All`, and
ideally `InformationProtectionPolicy.Read.All` (see the permissions table
above).

### Dry run first (recommended)

```powershell
$pwd = Read-Host "PFX password" -AsSecureString

.\Apply-M365FileSensitivityLabel-PnPOnly.ps1 `
  -TenantName "synozur" `
  -ClientId "df7b6a64-8d68-4e75-b15d-2cffc07cb554" `
  -Tenant "synozur.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -WhatIf
```

`-WhatIf` writes audit rows with `Action=WouldLabel` for every file that
would be touched.

### Apply each library's default label to its files

```powershell
.\Apply-M365FileSensitivityLabel-PnPOnly.ps1 `
  -TenantName "synozur" `
  -ClientId "df7b6a64-8d68-4e75-b15d-2cffc07cb554" `
  -Tenant "synozur.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -Resume
```

Libraries with no `DefaultSensitivityLabelForLibrary` are reported
`[skip-no-label]` and skipped — run
`Set-M365LibraryDefaultSensitivityLabel-PnPOnly.ps1` first if you want
those covered.

### Apply one label to every file across the tenant

```powershell
.\Apply-M365FileSensitivityLabel-PnPOnly.ps1 `
  -TenantName "synozur" `
  -ClientId "df7b6a64-8d68-4e75-b15d-2cffc07cb554" `
  -Tenant "synozur.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -LabelId "00000000-0000-0000-0000-000000000000" `
  -AssignmentMethod standard `
  -Resume
```

### Overwrite files that already have a different label

```powershell
.\Apply-M365FileSensitivityLabel-PnPOnly.ps1 `
  -TenantName "synozur" `
  -ClientId "df7b6a64-8d68-4e75-b15d-2cffc07cb554" `
  -Tenant "synozur.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -LabelId "00000000-0000-0000-0000-000000000000" `
  -AssignmentMethod privileged `
  -OverwriteExisting `
  -Resume
```

Use `-AssignmentMethod privileged` when an authorized user is explicitly
overriding an existing label; `auto` is reserved for automated policy
contexts; `standard` (the default) reflects normal user-initiated
application.

### Cap a rollout to a subset of sites and a per-library limit

```powershell
.\Apply-M365FileSensitivityLabel-PnPOnly.ps1 `
  -TenantName "synozur" `
  -ClientId "df7b6a64-8d68-4e75-b15d-2cffc07cb554" `
  -Tenant "synozur.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -SiteUrlLike "*/sites/Finance*" `
  -MaxFilesPerLibrary 500
```

### Parameters

Required:

- `TenantName`, `ClientId`, `Tenant`, `CertificatePath`, `CertificatePassword` — same as inventory

Optional:

- `LabelId` — force one label across every library; if omitted, each
  library's `DefaultSensitivityLabelForLibrary` is used and libraries with
  no default are skipped
- `AssignmentMethod` — `standard` (default), `privileged`, or `auto`
- `OverwriteExisting` — relabel files that already carry a different label
- `Resume` — skip files already written as `Labeled`/`AlreadyLabeled`
- `SiteUrlLike` — wildcard filter on site URL
- `MaxFilesPerLibrary` — cap per library (0 = no cap)
- `IncludeOneDriveSites` / `IncludeHiddenLibraries`
- `OutputCsvPath` / `ErrorCsvPath`
- `-WhatIf` — dry run

### Actions written to the audit CSV

| Action | Meaning |
|---|---|
| `Labeled` | Label was assigned to the file. |
| `AlreadyLabeled` | File already had the target label. |
| `SkippedExistingLabel` | File has a different label; re-run with `-OverwriteExisting` to replace. |
| `WouldLabel` | `-WhatIf` dry-run row. |
| `Error` | Assign call failed; details in the errors CSV. |

The script honors `Retry-After` on 429 / 503 responses and falls back to
exponential backoff (2s, 4s, ... capped at 60s, up to 5 attempts) so large
rollouts cooperate with Graph throttling.

---

## Outputs

### Inventory CSV (`Export-M365SensitivityLabelInventory-PnPOnly.ps1`)

- `Timestamp`
- `Scope` (site vs library rows)
- `SiteUrl`
- `SiteSensitivityLabelId`
- `SiteSensitivityLabelName`
- `LibraryTitle`
- `LibraryId`
- `LibraryServerRelativeUrl`
- `LibraryDefaultLabelId`
- `LibraryDefaultLabelName`

### Library default-label apply CSV (`M365_LibraryDefaultLabel_Apply.csv` by default)

- `Timestamp`, `SiteUrl`, `LibraryTitle`, `LibraryId`, `LibraryServerRelativeUrl`
- `PreviousLabelId`, `PreviousLabelName`
- `TargetLabelId`, `TargetLabelName`
- `Action`, `Detail`

### File apply CSV (`M365_FileLabel_Apply.csv` by default)

- `Timestamp`, `SiteUrl`, `LibraryTitle`
- `DriveId`, `ItemId`, `ItemPath`
- `PreviousLabelId`, `TargetLabelId`, `TargetLabelName`
- `Action`, `Detail`

> If label name mapping is unavailable (`InformationProtectionPolicy.Read.All`
> not granted), `*LabelName` columns will be empty and label IDs will still
> be written.

---

## Runtime "signs of life"

The scripts print:

- "Connecting…" to the admin URL
- "Retrieving sites…" and "Sites to process: N"
- Per-site banner `[i/N] https://.../sites/X`
- Per-library / per-file action lines (`[+] set`, `[=] already`, `[!] skipped`, `[?] would`, `[x] error`)
- Rolling progress every 10 sites (and every 100 files for the file apply script)
- A final summary block with totals and the CSV paths

---

## Troubleshooting

### "Label name mapping unavailable; logging GUIDs only."

`Get-PnPAvailableSensitivityLabel` couldn't run. Grant
**Graph Application** permission `InformationProtectionPolicy.Read.All` on
the same app registration and admin-consent it.

### File apply script returns 403 on `assignSensitivityLabel`

Confirm `Files.ReadWrite.All` and `Sites.Read.All` are granted as
**Application** permissions and admin-consented, and that sensitivity
labels for SharePoint/OneDrive are enabled (see
[Enable sensitivity labels for SharePoint/OneDrive](#enable-sensitivity-labels-for-sharepointonedrive-required)).

### Library default-label set call succeeds but new files aren't labeled

Setting `DefaultSensitivityLabelForLibrary` only affects newly uploaded /
created files. Use `Apply-M365FileSensitivityLabel-PnPOnly.ps1` to label
existing content.

### PDF label coverage looks wrong

Ensure sensitivity labels for SharePoint/OneDrive are enabled, and
optionally enable PDF support
(`Set-SPOTenant -EnableSensitivityLabelforPDF $true`).

### Archived / restricted sites

Some archived/restricted sites still appear in tenant enumeration but block
list enumeration (403 Forbidden). The scripts treat these as skipped sites
(written to the errors CSV with `Operation=EnumerateLibraries`) rather than
hard failures.

### Throttling on large rollouts

The file apply script honors `Retry-After` and falls back to exponential
backoff. For very large tenants, prefer `-SiteUrlLike` and
`-MaxFilesPerLibrary` to stage the rollout, and run with `-Resume` so an
interrupted run picks up where it left off.

---

## References

- [Enable sensitivity labels for files in SharePoint and OneDrive](https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-onedrive-files)
- [Auto-labeling for PDFs + `Set-SPOTenant -EnableSensitivityLabelforPDF`](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/set-spotenant?view=sharepoint-ps#-enablesensitivitylabelforpdf)
- [PnP.PowerShell `Get-PnPAvailableSensitivityLabel`](https://pnp.github.io/powershell/cmdlets/Get-PnPAvailableSensitivityLabel.html)
- [PnP.PowerShell `Set-PnPList -DefaultSensitivityLabelForLibrary`](https://pnp.github.io/powershell/cmdlets/Set-PnPList.html)
- [Graph `driveItem: assignSensitivityLabel`](https://learn.microsoft.com/en-us/graph/api/driveitem-assignsensitivitylabel)
