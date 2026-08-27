# AGENTS.md

## Cursor Cloud specific instructions

This repository is a **PowerShell script toolkit** for Microsoft 365 / SharePoint Online
administration. There is **no web app, no local server, no database, and no build step**.
Each `.ps1` is a one-shot CLI script that runs against a live M365 tenant over HTTPS.

### What is already installed (baked into the VM snapshot)

- **PowerShell 7.6.x** (`pwsh`) — installed from the Microsoft apt repo. `Inventory/`
  PnP scripts require PowerShell 7.4+.
- **PnP.PowerShell** (primary runtime dependency for the `Inventory/` scripts) and
  **PSScriptAnalyzer** (linting) are installed for the current user.
- Individual scripts also **auto-install their own runtime modules on first use**
  (`Microsoft.Online.SharePoint.PowerShell`, `Microsoft.Graph` / `Microsoft.Graph.Reports`,
  `ExchangeOnlineManagement`), so you usually don't need to pre-install them.

### Lint

```bash
pwsh -NoProfile -Command "Import-Module PSScriptAnalyzer; Invoke-ScriptAnalyzer -Path . -Recurse"
```

There are **0 error-level** findings; the remaining findings are all non-blocking style
warnings (mostly `PSAvoidUsingWriteHost`), which are expected for interactive admin scripts.

### Syntax check (fast, no cloud needed)

```bash
pwsh -NoProfile -Command '$e=$null;$t=$null;Get-ChildItem -Recurse -Filter *.ps1 | ForEach-Object { [System.Management.Automation.Language.Parser]::ParseFile($_.FullName,[ref]$t,[ref]$e)|Out-Null; if($e){\"$($_.Name): $($e.Count) errors\"} }'
```

### Tests

There is **no automated test suite** (no Pester tests, no CI). Verification is limited to
syntax parsing, PSScriptAnalyzer lint, and `Get-Help` / parameter-validation checks.

### Running the scripts (IMPORTANT: needs a real tenant)

- **Every script's core functionality requires a live Microsoft 365 tenant plus admin
  credentials** (interactive sign-in, or an Entra app + certificate for app-only auth).
  You **cannot** produce real report/inventory output in this environment without those.
- Interactive auth (`Connect-MgGraph` / `Connect-PnPOnline` without cert) will try to open a
  browser and **hang in this headless VM** — avoid it. Use `-WhatIf` on the apply scripts and
  app-only certificate auth when a tenant is available.
- The safest offline "signs of life" are: `Get-Help ./<script>.ps1`, ValidateSet/mandatory
  parameter checks, and generating an app-only certificate with
  `New-PnPAzureCertificate` (the real one-time-setup action the `Inventory/` scripts consume).
  Note: `Inventory/Setup-SPOInventoryApp.ps1` uses Windows-only PKI cmdlets
  (`New-SelfSignedCertificate`, `Export-PfxCertificate`) and will **not** run on Linux; use
  `New-PnPAzureCertificate` / `Register-PnPEntraIDApp` (see `Setup-SPOInventoryQuick.ps1`) instead.

See `README.md`, `README-Get-SPOSiteUsageReports.md`, and `Inventory/README.md` for the full
per-script parameter and usage documentation.
