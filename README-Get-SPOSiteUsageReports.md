# Get-SPOSiteUsageReports.ps1

## Overview
This PowerShell script enumerates usage reports from all SharePoint sites in a tenant and exports the data to a CSV file. It provides comprehensive site usage statistics including storage usage, activity dates, file counts, and more.

## Features
- **Three Methods**: Choose between Microsoft Graph API, SharePoint Online Management Shell, or a Combined mode that merges the best of both
- **Comprehensive Data**: Collects detailed usage statistics for all sites in the tenant
- **Automatic Module Installation**: Detects and installs required PowerShell modules
- **Progress Tracking**: Shows real-time progress when processing multiple sites
- **Error Handling**: Robust error handling with detailed logging
- **Flexible Output**: Customizable CSV output path with automatic timestamping
- **Summary Statistics**: Displays aggregate statistics after export

## Prerequisites

### For SharePoint Online Management Shell Method (Default)
- PowerShell 5.1 or later
- Microsoft.Online.SharePoint.PowerShell module (auto-installed if missing)
- SharePoint Administrator or Global Administrator role

### For Microsoft Graph API Method
- PowerShell 5.1 or later
- Microsoft.Graph.Reports module (auto-installed if missing)
- Microsoft.Graph.Authentication module (auto-installed if missing)
- Reports.Read.All, Sites.Read.All, and ReportSettings.ReadWrite.All permissions

### For Combined Mode
- PowerShell 5.1 or later
- All modules from both methods above (auto-installed if missing)
- Both SharePoint Administrator role and Graph API permissions

## Installation

1. Download the script:
```powershell
# Clone the repository
git clone https://github.com/chris-mcnulty/spscripts.git
cd spscripts
```

2. Ensure you have appropriate permissions in your SharePoint tenant

3. Run the script (modules will be auto-installed if needed)

## Usage

### Basic Usage (SharePoint Online Management Shell)
```powershell
.\Get-SPOSiteUsageReports.ps1 -TenantName "contoso"
```
This will:
- Connect to contoso.sharepoint.com
- Retrieve usage data for all sites
- Export to a timestamped CSV file in the current directory

### Using Microsoft Graph API
```powershell
.\Get-SPOSiteUsageReports.ps1 -TenantName "contoso" -UseGraphAPI
```
This method provides additional metrics like page views and active file counts.

### Combined Mode (Best of Both)
```powershell
.\Get-SPOSiteUsageReports.ps1 -TenantName "contoso" -UseCombined
```
This mode uses SPO as the source of truth for site inventory (Title, URL, Owner — never obfuscated), then attaches Graph usage metrics (page views, file counts, activity dates). SPO URLs are resolved to Graph SiteIds using path-based addressing (`Get-MgSite -SiteId "contoso.sharepoint.com:/sites/SiteName:"`), which avoids `getAllSites` and works with delegated permissions.

### Specify Output Path
```powershell
.\Get-SPOSiteUsageReports.ps1 -TenantName "contoso" -OutputPath "C:\Reports\MyUsageReport.csv"
```

### Complete Example with All Parameters
```powershell
.\Get-SPOSiteUsageReports.ps1 `
    -TenantName "contoso" `
    -OutputPath "C:\Reports\usage.csv" `
    -AuthMethod "Interactive" `
    -UseGraphAPI
```

## Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| TenantName | String | Yes | - | Your SharePoint tenant name (e.g., 'contoso' for contoso.sharepoint.com) |
| OutputPath | String | No | Auto-generated | Full path for the output CSV file. If not specified, creates a timestamped file in script directory |
| AuthMethod | String | No | Interactive | Authentication method: 'Interactive', 'Certificate', or 'ClientSecret' |
| UseGraphAPI | Switch | No | False | Use Microsoft Graph API instead of SharePoint Online Management Shell |
| UseCombined | Switch | No | False | Combine SPO Management Shell (friendly names) with Graph API (page views/activity) into a single report |

## Output Data

### SharePoint Online Management Shell Method
The CSV file includes:
- **SiteUrl**: Full URL of the site
- **Title**: Site title
- **Owner**: Site owner email/username
- **Template**: Site template type
- **Status**: Site status (Active, ReadOnly, etc.)
- **StorageQuotaMB**: Storage quota in megabytes
- **StorageUsedMB**: Current storage usage in megabytes
- **StorageUsedPercentage**: Percentage of storage quota used
- **LastContentModifiedDate**: Date of last content modification
- **SharingCapability**: External sharing settings
- **LockState**: Site lock status
- **WebsCount**: Number of subsites
- **IsHubSite**: Whether the site is a hub site
- **HubSiteId**: Hub site ID if connected to a hub
- **SensitivityLabel**: Applied sensitivity label
- **CreatedDate**: Site creation date

### Microsoft Graph API Method
The Graph method retrieves usage data — including page views — as follows:

1. The script calls `Invoke-MgGraphRequest` against the Graph REST endpoint [`getSharePointSiteUsageDetail`](https://learn.microsoft.com/en-us/graph/api/reportroot-getsharepointsiteusagedetail) (bypasses the `Get-MgReportSharePointSiteUsageDetail` cmdlet which has a known PercentComplete overflow bug).
2. That endpoint returns a CSV report where each row is a site and columns include `Page View Count` and `Visited Page Count`.
3. The script parses the CSV with `Import-Csv` and maps those columns to the output fields listed below.

Page views are tracked server-side by SharePoint and surfaced through this report; there is no separate API call to fetch them.

The CSV file includes:
- **SiteUrl**: Full URL of the site
- **SiteId**: Unique site identifier
- **OwnerDisplayName**: Display name of the owner
- **OwnerPrincipalName**: Owner's principal name
- **IsDeleted**: Whether the site is deleted
- **LastActivityDate**: Date of last activity
- **FileCount**: Total number of files
- **ActiveFileCount**: Number of active files
- **PageViewCount**: Number of page views (from the report's `Page View Count` column)
- **VisitedPageCount**: Number of distinct pages visited (from the report's `Visited Page Count` column)
- **StorageUsedInBytes**: Storage used in bytes
- **StorageAllocatedInBytes**: Allocated storage in bytes
- **RootWebTemplate**: Site template
- **ReportRefreshDate**: When the report was generated
- **ReportPeriod**: Report period (e.g., last 7 days)

### Combined Mode
The combined mode uses SPO as the source of truth for site inventory and enriches each site with Graph usage metrics. When the Graph report data is obfuscated (zeroed SiteIds, blank URLs), the script automatically falls back to **per-site Graph analytics** using the compound site IDs resolved from SPO URLs. The CSV file includes:
- **SiteUrl**: Full URL of the site (from SPO — always reliable)
- **Title**: Site title (from SPO)
- **Owner**: Site owner email/username (from SPO)
- **Template**: Site template type (from SPO)
- **StorageUsedMB**: Current storage usage in megabytes (from SPO)
- **StorageQuotaMB**: Storage quota in megabytes (from SPO)
- **StorageUsedPercentage**: Percentage of storage quota used (from SPO)
- **LastContentModifiedDate**: Date of last content modification (from SPO)
- **LastActivityDate**: Date of last activity (from Graph)
- **FileCount**: Total number of files (from Graph)
- **ActiveFileCount**: Number of active files (from Graph)
- **PageViewCount**: Number of page views in last 7 days (from Graph)
- **VisitedPageCount**: Number of distinct pages visited in last 7 days (from Graph)
- **SharingCapability**: External sharing settings (from SPO)
- **LockState**: Site lock status (from SPO)
- **IsHubSite**: Whether the site is a hub site (from SPO)
- **HubSiteId**: Hub site ID if connected to a hub (from SPO)
- **SensitivityLabel**: Applied sensitivity label (from SPO)
- **RootWebTemplate**: Site template (from Graph)
- **IsDeleted**: Whether the site has been deleted (from Graph)
- **CreatedDate**: Site creation date (from SPO)
- **ReportRefreshDate**: When the Graph report was generated

## Examples

### Example 1: Quick Report for Contoso Tenant
```powershell
.\Get-SPOSiteUsageReports.ps1 -TenantName "contoso"
```
Output: `SPO_SiteUsage_contoso_20260212_143022.csv`

### Example 2: Detailed Report Using Graph API
```powershell
.\Get-SPOSiteUsageReports.ps1 -TenantName "fabrikam" -UseGraphAPI -OutputPath "D:\Reports\Fabrikam_Usage.csv"
```
Provides additional metrics like page views and active file counts.

### Example 3: Combined Report with Friendly Names and Page Views
```powershell
.\Get-SPOSiteUsageReports.ps1 -TenantName "contoso" -UseCombined -OutputPath "C:\Reports\Combined_Usage.csv"
```
Merges SPO friendly site info (titles, owners, storage in MB) with Graph activity metrics (page views, active files).

### Example 4: Automated Monthly Report
```powershell
# Schedule this script to run monthly via Task Scheduler
$monthYear = Get-Date -Format "yyyy-MM"
$outputPath = "\\fileserver\Reports\SharePoint\Usage_$monthYear.csv"

.\Get-SPOSiteUsageReports.ps1 -TenantName "contoso" -OutputPath $outputPath -UseGraphAPI
```

## Troubleshooting

### Issue: Module Installation Fails ("Administrator rights are required")
The script installs modules with `-Scope CurrentUser` and bootstraps the NuGet provider the same way. You do **not** need an elevated (Run as Administrator) session.

If install still fails, run these once as your **normal** user (not Administrator), then re-run the script:
```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force
Install-Module -Name Microsoft.Graph.Reports, Microsoft.Graph.Authentication, Microsoft.Online.SharePoint.PowerShell `
    -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -SkipPublisherCheck
```

### Which PowerShell should I use?
- **Preferred for `-UseCombined` / default SPO mode**: Windows PowerShell 5.1 (`powershell.exe`), or the SharePoint Online Management Shell (same engine with SPO preinstalled).
- **PowerShell 7+ (`pwsh`)**: Works. Graph mode is fine natively; SPO mode imports `Microsoft.Online.SharePoint.PowerShell` via `-UseWindowsPowerShell` (Windows only).
- **Avoid**: An elevated "Run as Administrator" session just to install modules — use CurrentUser scope instead.

### Issue: Authentication Fails
**Solution**: Ensure you have the required permissions:
- For SPO method: SharePoint Administrator or Global Administrator
- For Graph method: Reports.Read.All and Sites.Read.All permissions

### Issue: Script Takes Too Long
**Solution**: 
- Use the `-UseGraphAPI` switch for faster processing
- The Graph API method is generally faster for large tenants
- For SPO method, the script processes sites one by one with progress indicators

### Issue: Some Sites Missing from Report
**Solution**:
- Check that you have access to all sites
- Personal sites are excluded by default; modify the script if needed
- Some sites may be hidden or require special permissions

### Issue: `-Verbose` Produces Garbled CSV or Excessive Output
**Solution**: The script now automatically suppresses `-Verbose` propagation to internal Graph API calls. If you still see excessive VERBOSE output:
- This is expected behavior from the Microsoft Graph PowerShell SDK — each HTTP request generates two VERBOSE lines
- The script isolates Graph calls so they don't inherit `-Verbose`, but some cmdlets may still produce output
- If the CSV file appears corrupted (contains module manifest properties instead of report data), the script will automatically detect this and fall back to the typed cmdlet

### Issue: Graph API Returns Obfuscated Data (Zeroed SiteIds, Hashed Names)
**Solution**: The Microsoft 365 admin center has a privacy setting called **"Conceal user, group, and site names in all reports"** that causes the Graph Reports API to return hashed owner names, zeroed-out Site IDs, and empty Site URLs.

The script now handles this automatically:
1. **Uses SPO as source of truth** — `Get-SPOSite` reliably returns Title, URL, and Owner even when the Graph Reports API is obfuscated
2. **Robust report download** — `Get-GraphReportCsv` tries `Invoke-MgGraphRequest -OutputFilePath` first, validates the result is real CSV, and falls back to `Get-MgReportSharePointSiteUsageDetail -OutFile` with progress suppression
3. **Suppresses `-Verbose` noise** — internal Graph API calls (`Get-MgSite`, `Invoke-MgGraphRequest`) are isolated from the `-Verbose` preference to prevent hundreds of VERBOSE lines and potential file corruption
4. **Resolves SiteIds with visible error logging** — uses `Get-MgSite` path-based addressing and `Invoke-MgGraphRequest` to resolve SPO URLs to Graph compound site IDs; logs first failures as warnings so you can diagnose 403 vs 404 errors
5. **Falls back to per-site analytics** — when the report join fails (concealed data), automatically falls back to `/sites/{id}/analytics/lastSevenDays` and `/sites/{id}/drive` using resolved compound site IDs (not subject to report concealment)
6. **Checks the admin setting** via `(Get-MgAdminReportSetting).DisplayConcealedNames` and **attempts to disable it** via `Update-MgAdminReportSetting -DisplayConcealedNames:$false`
7. **Validates output** — checks that SiteUrl and Title are populated in the final output before export, warns if they are blank

**Important notes:**
- In combined mode (`-UseCombined`), SPO metadata (Title, URL, Owner) is always available regardless of Graph obfuscation
- When the report join fails, per-site analytics provides PageViewCount, FileCount, StorageUsedBytes, and LastActivityDate
- `ActiveFileCount` and `VisitedPageCount` are only available from the Reports API (not from per-site analytics)
- If SiteId resolution fails with 403 errors, you may need to temporarily grant Site Collection Admin rights
- After disabling the concealment setting, report data can take **up to 48 hours** to reflect the change
- Grant `ReportSettings.ReadWrite.All` permission for automatic setting correction

## Permissions Required

### SharePoint Online Management Shell Method
- SharePoint Administrator role, or
- Global Administrator role

### Microsoft Graph API Method
Application permissions needed:
- Reports.Read.All
- Sites.Read.All
- ReportSettings.ReadWrite.All (for automatic concealment setting detection and correction)

Or delegated permissions:
- Reports.Read.All
- Sites.Read.All
- ReportSettings.ReadWrite.All (optional, but recommended)

## Best Practices

1. **Run during off-peak hours**: For large tenants, run the script during off-peak hours to minimize impact
2. **Regular exports**: Schedule regular exports to track usage trends over time
3. **Store securely**: Store output CSV files in a secure location as they contain sensitive information
4. **Review permissions**: Regularly review and ensure you have the minimum required permissions
5. **Graph API for scale**: For tenants with 1000+ sites, use the Graph API method for better performance

## Limitations

- Personal sites (OneDrive) are excluded by default in the SPO method
- Graph API method provides data for the last 7 days by default
- Rate limiting may occur with very large tenants (10,000+ sites)
- Some fields may be empty depending on site configuration and permissions

## Support

For issues, questions, or contributions, please visit:
https://github.com/chris-mcnulty/spscripts

## License

This script is provided as-is without warranty. Use at your own risk.

## Version History

- **2.2.0** (2026-02-13): Fix `-Verbose` garbled CSV and improve report download robustness
  - Added `Get-GraphReportCsv` helper that downloads the usage report with validation: tries `Invoke-MgGraphRequest -OutputFilePath` first, validates CSV headers, falls back to `Get-MgReportSharePointSiteUsageDetail -OutFile` with progress suppression
  - Suppressed `-Verbose` propagation to all internal Graph API calls (`Get-MgSite`, `Invoke-MgGraphRequest`) — prevents hundreds of noisy VERBOSE lines and file corruption during resolution loops
  - Added diagnostic validation: SPO data sample at load time, URL/Title presence check before export
  - CSV header validation catches garbled report downloads (Graph manifest properties) before parsing
- **2.1.0** (2026-02-13): Fix reports cmdlet crash, add per-site analytics fallback, improve error visibility
  - Replaced `Get-MgReportSharePointSiteUsageDetail` with `Invoke-MgGraphRequest` to fix PercentComplete overflow crash
  - Added `Resolve-GraphSiteId` helper with dual approach (Get-MgSite + Invoke-MgGraphRequest) and visible error logging
  - First N resolution failures are shown as warnings (diagnose 403 vs 404 vs format issues)
  - When report join fails (<20% match rate), automatically falls back to per-site Graph analytics using resolved compound site IDs
  - Per-site analytics fallback provides PageViewCount, FileCount, StorageUsedBytes, and LastActivityDate even when the report is fully concealed
- **2.0.0** (2026-02-13): SPO-first combined mode with `Get-MgSite` path-based addressing
  - Combined mode now uses SPO as source of truth for site inventory (Title, URL, Owner — never obfuscated)
  - Resolves SPO URLs to Graph SiteIds via `Get-MgSite -SiteId "hostname:/sites/Name:"` — avoids `getAllSites` entirely
  - Matches SPO sites to Graph usage report by URL first, then by resolved SiteId
  - Graph connection managed independently from SPO (no session lifecycle issues)
  - Graceful degradation: returns SPO-only data if Graph connection fails
  - Validates Graph permissions at start of combined mode
- **1.4.2** (2026-02-12): Fix three bugs preventing Graph-to-SPO matching
  - Fixed `Get-MgSite` to use compound SiteId format (`hostname,siteGuid,siteGuid`) instead of simple GUIDs
  - Fixed Graph session being disconnected before combined mode could use it (`-KeepConnection` switch)
  - Fixed sites search fallback using invalid `?search=*` syntax (now uses tenant name as search term)
- **1.4.1** (2026-02-12): Get-MgSite lookup for blank-URL sites
  - For Graph sites with valid SiteIds but blank URLs, calls `Get-MgSite -SiteId` to resolve displayName and webUrl
  - Applied in both Graph-only mode (resolves blank URLs before obfuscation check) and combined mode (third matching step after URL and SiteId)
  - Resolved webUrl is then used to match back to SPO for full metadata in combined mode
- **1.4.0** (2026-02-12): Reversed combined mode — Graph-first approach
  - Combined mode now starts from Graph report data (authoritative list with SiteIds and all usage metrics)
  - Each Graph site row is enriched with SPO metadata (Title, Owner, Template, etc.) when a match is found
  - Eliminates duplicate rows that occurred when starting from SPO and appending unmatched Graph sites
  - Matching: URL first, then SiteId (SPO URLs resolved to Graph SiteIds via `/sites/{hostname}:/{path}`)
- **1.3.0** (2026-02-12): SiteId-based matching for combined mode
  - Combined mode now falls back to SiteId-based matching when Graph report URLs are blank/obfuscated
  - Each SPO site URL is resolved to its Graph Site ID via `/sites/{hostname}:/{path}`, then matched to Graph report data by SiteId
  - Unmatched SPO sites are enriched with per-site analytics (page views, file counts) via direct Graph API calls
  - URL-based matching remains the preferred first pass when Graph URLs are available
- **1.2.1** (2026-02-12): Use proper Graph SDK cmdlets for report settings
  - Replaced raw `Invoke-MgGraphRequest` calls with `Get-MgAdminReportSetting` and `Update-MgAdminReportSetting` for reliable concealment setting detection and correction
  - Added explicit logging of the actual `DisplayConcealedNames` value so users can confirm the setting
- **1.2.0** (2026-02-12): Graph API obfuscation handling
  - Automatic detection of obfuscated Graph report data (zeroed SiteIds, hashed names)
  - Proactive check and correction of the admin report-privacy concealment setting
  - Sites API fallback (`/sites/getAllSites`) for real site metadata when reports are obfuscated
  - Per-site analytics enrichment via `/sites/{id}/analytics/lastSevenDays` and `/sites/{id}/drive` to retrieve page views, file counts, storage, and last activity dates even when the report is obfuscated
  - Improved combined mode diagnostics when partial metrics are unavailable
- **1.1.0** (2026-02-12): Combined mode
  - New `-UseCombined` switch merges SPO + Graph data into one report
  - Friendly names/owners from SPO with page views/activity from Graph
- **1.0.0** (2026-02-12): Initial release
  - Support for SharePoint Online Management Shell
  - Support for Microsoft Graph API
  - Automatic module installation
  - Comprehensive usage data export
  - Progress tracking and error handling
