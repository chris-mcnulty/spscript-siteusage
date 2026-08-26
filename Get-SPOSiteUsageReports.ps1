<#
.SYNOPSIS
    Enumerate usage reports from all SharePoint sites in a tenant to a CSV file.

.DESCRIPTION
    This script connects to SharePoint Online and retrieves usage information for all sites 
    in the tenant, then exports the data to a CSV file. The script uses Microsoft Graph API
    or SharePoint Online Management Shell to collect site usage statistics.

.PARAMETER TenantName
    The name of your SharePoint Online tenant (e.g., 'contoso' for contoso.sharepoint.com).

.PARAMETER OutputPath
    The path where the CSV file will be saved. Defaults to current directory with timestamp.

.PARAMETER AuthMethod
    The authentication method to use: 'Interactive' (default), 'Certificate', or 'ClientSecret'.

.PARAMETER UseGraphAPI
    If specified, uses Microsoft Graph API instead of SharePoint Online Management Shell.
    Requires Microsoft.Graph PowerShell modules.

.PARAMETER UseCombined
    If specified, combines data from both SPO Management Shell and Graph API.
    Produces a single report with friendly site names/owners from SPO and
    page view/activity metrics from Graph. Requires both sets of modules.

.EXAMPLE
    .\Get-SPOSiteUsageReports.ps1 -TenantName "contoso"
    Connects to contoso.sharepoint.com and exports usage reports to a timestamped CSV file.

.EXAMPLE
    .\Get-SPOSiteUsageReports.ps1 -TenantName "contoso" -OutputPath "C:\Reports\usage.csv" -UseGraphAPI
    Uses Microsoft Graph API to get usage data and exports to specified path.

.EXAMPLE
    .\Get-SPOSiteUsageReports.ps1 -TenantName "contoso" -UseCombined
    Combines SPO Management Shell (friendly names) and Graph API (page views) into one report.

.NOTES
    Author: chris-mcnulty/spscripts
    Requirements:
    - For SharePoint method: SharePointPnPPowerShellOnline or Microsoft.Online.SharePoint.PowerShell module
    - For Graph method: Microsoft.Graph.Reports module
    - Appropriate permissions: SharePoint Admin or Global Admin
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter your SharePoint tenant name (e.g., 'contoso' for contoso.sharepoint.com)")]
    [string]$TenantName,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Certificate', 'ClientSecret')]
    [string]$AuthMethod = 'Interactive',

    [Parameter(Mandatory = $false)]
    [switch]$UseGraphAPI,

    [Parameter(Mandatory = $false, HelpMessage = "Combines SPO Management Shell (friendly names) with Graph API (page views/activity) into a single report")]
    [switch]$UseCombined
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Function to check, install, and import required modules into the current session.
# Microsoft.Online.SharePoint.PowerShell is Windows PowerShell 5.1-native; on PowerShell 7+
# it must be imported with -UseWindowsPowerShell or Connect/Disconnect-SPOService will not
# resolve (and a bare Disconnect in a finally block can mask the real connect error).
function Install-RequiredModules {
    param(
        [string[]]$ModuleNames
    )

    foreach ($moduleName in $ModuleNames) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            Write-Host "Module '$moduleName' not found. Attempting to install..." -ForegroundColor Yellow
            try {
                # Prefer CurrentUser so elevation is not required. Suppress the common
                # PackageManagement/NuGet bootstrap warning that still allows a CurrentUser install.
                Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Host "Module '$moduleName' installed successfully." -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to install module '$moduleName': $_"
                Write-Error "If you see 'Administrator rights are required', close this session and re-run as a normal user with: Install-Module $moduleName -Scope CurrentUser -Force"
                return $false
            }

            if (-not (Get-Module -ListAvailable -Name $moduleName)) {
                Write-Error "Module '$moduleName' still not found after Install-Module. Install may have failed silently (often a NuGet provider / permissions issue)."
                return $false
            }
        }

        # Ensure the module is loaded in *this* session (ListAvailable alone is not enough).
        if (-not (Get-Module -Name $moduleName)) {
            try {
                $importParams = @{
                    Name        = $moduleName
                    Force       = $true
                    ErrorAction = 'Stop'
                }
                # SPO Management Shell needs Windows PowerShell compatibility on PS 7+.
                if ($moduleName -eq 'Microsoft.Online.SharePoint.PowerShell' -and
                    $PSVersionTable.PSEdition -eq 'Core') {
                    $importParams['UseWindowsPowerShell'] = $true
                    Write-Host "Importing '$moduleName' with -UseWindowsPowerShell (required on PowerShell 7+)..." -ForegroundColor Cyan
                }
                Import-Module @importParams
            }
            catch {
                Write-Error "Failed to import module '$moduleName': $_"
                if ($moduleName -eq 'Microsoft.Online.SharePoint.PowerShell' -and $PSVersionTable.PSEdition -eq 'Core') {
                    Write-Error "On PowerShell 7+, Microsoft.Online.SharePoint.PowerShell requires Windows + Import-Module -UseWindowsPowerShell. Prefer Windows PowerShell 5.1, or run Graph-only mode (-UseGraphAPI)."
                }
                return $false
            }
        }
    }
    return $true
}

# Safely disconnect SPO so a missing Disconnect-SPOService cmdlet never masks the real error.
function Disconnect-SPOServiceSafe {
    $cmd = Get-Command Disconnect-SPOService -ErrorAction SilentlyContinue
    if ($cmd) {
        try {
            Disconnect-SPOService -ErrorAction SilentlyContinue
        }
        catch {
            # Best-effort cleanup only
        }
    }
}

# Function to get usage reports using Microsoft Graph API
function Get-UsageReportsViaGraph {
    param(
        [string]$TenantName,
        [switch]$KeepConnection
    )
    
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    
    try {
        # Connect to Microsoft Graph — include ReportSettings.ReadWrite.All so the
        # script can detect and disable the report-privacy concealment setting.
        $null = Connect-MgGraph -Scopes "Reports.Read.All", "Sites.Read.All", "ReportSettings.ReadWrite.All" -NoWelcome
        
        # Proactively check the admin concealment setting before pulling the report.
        $privacyResult = Resolve-GraphReportPrivacy
        
        Write-Host "Retrieving SharePoint site usage data from Microsoft Graph..." -ForegroundColor Cyan
        
        # Get SharePoint site usage details for the last 7 days
        $usageData = @()
        
        # Get site usage detail report via the robust download helper, which
        # tries Invoke-MgGraphRequest first (avoids PercentComplete overflow)
        # and falls back to the typed cmdlet with progress suppressed.
        $tempFile = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName() + ".csv")
        $downloadOk = Get-GraphReportCsv -TempFile $tempFile
        if (-not $downloadOk) {
            throw "Could not download SharePoint site usage report from Graph. Both download methods failed."
        }
        
        # Parse the CSV data returned by Graph API
        $sites = Import-Csv -Path $tempFile
        
        foreach ($site in $sites) {
            $siteInfo = [PSCustomObject]@{
                SiteUrl              = $site.'Site URL'
                SiteId               = $site.'Site Id'
                OwnerDisplayName     = $site.'Owner Display Name'
                OwnerPrincipalName   = $site.'Owner Principal Name'
                IsDeleted            = $site.'Is Deleted'
                LastActivityDate     = $site.'Last Activity Date'
                FileCount            = $site.'File Count'
                ActiveFileCount      = $site.'Active File Count'
                PageViewCount        = $site.'Page View Count'
                VisitedPageCount     = $site.'Visited Page Count'
                StorageUsedInBytes   = $site.'Storage Used (Byte)'
                StorageAllocatedInBytes = $site.'Storage Allocated (Byte)'
                RootWebTemplate      = $site.'Root Web Template'
                ReportRefreshDate    = $site.'Report Refresh Date'
                ReportPeriod         = $site.'Report Period'
            }
            $usageData += $siteInfo
        }
        
        Write-Host "Retrieved usage data for $($usageData.Count) sites." -ForegroundColor Green

        # --- Resolve blank URLs via Get-MgSite for sites with valid SiteIds ---
        # The Graph report may return real SiteIds but blank URLs when the
        # concealment setting is active.  Get-MgSite is NOT affected by report
        # concealment and returns real displayName and webUrl from a SiteId.
        $emptyGuid = '00000000-0000-0000-0000-000000000000'
        $blankUrlSites = $usageData | Where-Object {
            [string]::IsNullOrWhiteSpace($_.SiteUrl) -and
            $_.SiteId -and $_.SiteId -ne $emptyGuid
        }

        if ($blankUrlSites.Count -gt 0) {
            Write-Host "Resolving $($blankUrlSites.Count) sites with blank URLs via Get-MgSite..." -ForegroundColor Yellow
            $resolvedCount = 0
            $hostname = "$TenantName.sharepoint.com"

            # Suppress verbose output from Graph API calls
            $savedVerbose = $VerbosePreference
            $VerbosePreference = 'SilentlyContinue'

            foreach ($blankSite in $blankUrlSites) {
                try {
                    # The Graph report returns simple GUIDs as SiteIds, but Get-MgSite
                    # requires a compound ID: "hostname,siteGuid,webGuid".  For site
                    # collections the root web GUID typically equals the site GUID.
                    $siteGuid = $blankSite.SiteId
                    $compoundId = "$hostname,$siteGuid,$siteGuid"
                    $mgSite = Get-MgSite -SiteId $compoundId -Property "id,displayName,webUrl" -ErrorAction Stop
                    if ($mgSite) {
                        if ($mgSite.WebUrl)      { $blankSite.SiteUrl = $mgSite.WebUrl }
                        # Store the site display name so it can be used as a fallback
                        # title/label when the OwnerDisplayName is obfuscated.
                        if ($mgSite.DisplayName -and
                            ($blankSite.OwnerDisplayName -match '^[A-Fa-f0-9]{32}$' -or
                             [string]::IsNullOrWhiteSpace($blankSite.OwnerDisplayName))) {
                            $blankSite.OwnerDisplayName = $mgSite.DisplayName
                        }
                        $resolvedCount++
                    }
                }
                catch {
                    # Could not resolve this SiteId — leave as-is
                }
            }

            $VerbosePreference = $savedVerbose
            Write-Host "Resolved $resolvedCount of $($blankUrlSites.Count) blank-URL sites via Get-MgSite." -ForegroundColor Green
        }

        # --- Obfuscation detection and Sites API fallback ---
        if (Test-GraphDataObfuscated -ReportData $usageData) {
            Write-Warning "Graph report data is obfuscated — site URLs, IDs, and owner names are concealed."

            if ($privacyResult.WasEnabled -eq $true) {
                Write-Warning "The report-privacy setting may have been recently changed. Cached report data can take up to 48 hours to reflect the new setting."
            }

            # Fall back to the Sites API which is unaffected by report concealment.
            $realSites = Get-SiteMetadataViaGraph -TenantName $TenantName

            if ($realSites.Count -gt 0) {
                Write-Host "Enriching site metadata with per-site analytics (not affected by report concealment)..." -ForegroundColor Yellow

                # Get per-site analytics — these endpoints are NOT subject to report privacy.
                $siteAnalytics = Get-PerSiteAnalyticsViaGraph -Sites $realSites

                $usageData = @()
                foreach ($realSite in $realSites) {
                    $enrichment = if ($realSite.webUrl) { $siteAnalytics[$realSite.webUrl] } else { $null }
                    $usageData += [PSCustomObject]@{
                        SiteUrl                 = $realSite.webUrl
                        SiteId                  = $realSite.id
                        OwnerDisplayName        = $realSite.displayName
                        OwnerPrincipalName      = ''
                        IsDeleted               = $false
                        LastActivityDate        = if ($enrichment -and $null -ne $enrichment.LastActivityDate) { $enrichment.LastActivityDate } else { '' }
                        FileCount               = if ($enrichment -and $null -ne $enrichment.FileCount) { $enrichment.FileCount } else { '' }
                        ActiveFileCount         = ''
                        PageViewCount           = if ($enrichment -and $null -ne $enrichment.PageViewCount) { $enrichment.PageViewCount } else { '' }
                        VisitedPageCount        = ''
                        StorageUsedInBytes      = if ($enrichment -and $null -ne $enrichment.StorageUsedBytes) { $enrichment.StorageUsedBytes } else { '' }
                        StorageAllocatedInBytes = ''
                        RootWebTemplate         = ''
                        ReportRefreshDate       = ''
                        ReportPeriod            = ''
                    }
                }
                Write-Host "Rebuilt report with $($usageData.Count) sites using real site metadata and per-site analytics." -ForegroundColor Green
            }
            else {
                Write-Warning "Could not retrieve site metadata from Sites API. Report will contain obfuscated data."
                Write-Warning "Ensure the 'Conceal user, group, and site names in all reports' setting is disabled and re-run after 48 hours."
            }
        }

        return $usageData
    }
    catch {
        Write-Error "Error retrieving data from Microsoft Graph: $_"
        throw
    }
    finally {
        if ($tempFile) { Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue }
        if (-not $KeepConnection) {
            $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        }
    }
}

# Function to get usage reports using SharePoint Online Management Shell
function Get-UsageReportsViaSPO {
    param(
        [string]$TenantName
    )
    
    $adminUrl = "https://$TenantName-admin.sharepoint.com"
    
    Write-Host "Connecting to SharePoint Online Admin Center: $adminUrl" -ForegroundColor Cyan
    
    try {
        # Connect to SharePoint Online
        $null = Connect-SPOService -Url $adminUrl
        
        Write-Host "Retrieving all SharePoint sites..." -ForegroundColor Cyan
        
        # Get all site collections
        $sites = Get-SPOSite -Limit All -IncludePersonalSite $false
        
        Write-Host "Found $($sites.Count) sites. Collecting usage data..." -ForegroundColor Cyan
        
        $usageData = @()
        $counter = 0
        
        foreach ($site in $sites) {
            $counter++
            Write-Progress -Activity "Processing Sites" -Status "Processing $counter of $($sites.Count): $($site.Url)" -PercentComplete (($counter / $sites.Count) * 100)
            
            try {
                # Get site details
                $siteDetails = Get-SPOSite -Identity $site.Url -Detailed
                
                $siteInfo = [PSCustomObject]@{
                    SiteUrl                     = $site.Url
                    Title                       = $site.Title
                    Owner                       = $site.Owner
                    Template                    = $site.Template
                    Status                      = $site.Status
                    StorageQuotaMB              = [math]::Round($site.StorageQuota / 1024, 2)
                    StorageUsedMB               = [math]::Round($site.StorageUsageCurrent / 1024, 2)
                    StorageUsedPercentage       = if ($site.StorageQuota -gt 0) { [math]::Round(($site.StorageUsageCurrent / $site.StorageQuota) * 100, 2) } else { 0 }
                    LastContentModifiedDate     = $site.LastContentModifiedDate
                    SharingCapability           = $site.SharingCapability
                    LockState                   = $site.LockState
                    PWAEnabled                  = $site.PWAEnabled
                    ConditionalAccessPolicy     = $site.ConditionalAccessPolicy
                    AllowSelfServiceUpgrade     = $site.AllowSelfServiceUpgrade
                    LocaleId                    = $siteDetails.LocaleId
                    CompatibilityLevel          = $siteDetails.CompatibilityLevel
                    WebsCount                   = $siteDetails.WebsCount
                    IsHubSite                   = $siteDetails.IsHubSite
                    HubSiteId                   = $siteDetails.HubSiteId
                    SensitivityLabel            = $siteDetails.SensitivityLabel
                    CreatedDate                 = $site.CreatedDate
                }
                
                $usageData += $siteInfo
            }
            catch {
                Write-Warning "Error processing site $($site.Url): $_"
            }
        }
        
        Write-Progress -Activity "Processing Sites" -Completed
        Write-Host "Retrieved usage data for $($usageData.Count) sites." -ForegroundColor Green
        
        return $usageData
    }
    catch {
        Write-Error "Error retrieving data from SharePoint Online: $_"
        throw
    }
    finally {
        # Must not call Disconnect-SPOService when the module failed to load —
        # CommandNotFoundException is terminating and would mask the real error.
        Disconnect-SPOServiceSafe
    }
}

# Robustly download the SharePoint Site Usage Detail CSV report from Graph.
# Tries Invoke-MgGraphRequest -OutputFilePath first (avoids the cmdlet's
# PercentComplete overflow bug), validates the result is real CSV, and falls
# back to Get-MgReportSharePointSiteUsageDetail -OutFile (with progress
# suppressed) if the first approach produces garbage.
#
# The function suppresses the VerbosePreference for internal Graph calls
# because -Verbose propagation causes Invoke-MgGraphRequest to log every
# HTTP request/response as verbose output, and in some SDK versions this
# interferes with -OutputFilePath, producing a file containing module
# manifest properties instead of the report CSV.
function Get-GraphReportCsv {
    param(
        [Parameter(Mandatory)][string]$TempFile
    )

    $savedVerbose = $VerbosePreference
    $VerbosePreference = 'SilentlyContinue'

    # Approach 1: Invoke-MgGraphRequest -OutputFilePath (avoids Write-Progress bug)
    try {
        Invoke-MgGraphRequest -Method GET `
            -Uri "/v1.0/reports/getSharePointSiteUsageDetail(period='D7')" `
            -OutputFilePath $TempFile -ErrorAction Stop

        # Validate the file is real CSV with expected report columns
        $header = Get-Content -Path $TempFile -TotalCount 1 -ErrorAction Stop
        if ($header -and $header -match 'Site Id') {
            $VerbosePreference = $savedVerbose
            return $true
        }
        Write-Host "Invoke-MgGraphRequest produced unexpected file content (not a valid report CSV). Trying fallback..." -ForegroundColor Yellow
    }
    catch {
        Write-Host "Invoke-MgGraphRequest for reports failed: $($_.Exception.Message). Trying fallback..." -ForegroundColor Yellow
    }

    # Approach 2: Use the typed cmdlet with progress suppressed
    $savedProgress = $ProgressPreference
    try {
        $ProgressPreference = 'SilentlyContinue'
        Get-MgReportSharePointSiteUsageDetail -Period D7 -OutFile $TempFile -ErrorAction Stop
    }
    catch {
        Write-Host "Get-MgReportSharePointSiteUsageDetail also failed: $($_.Exception.Message)" -ForegroundColor Yellow
        $ProgressPreference = $savedProgress
        $VerbosePreference = $savedVerbose
        return $false
    }
    finally {
        $ProgressPreference = $savedProgress
    }

    $header = Get-Content -Path $TempFile -TotalCount 1 -ErrorAction SilentlyContinue
    if ($header -and $header -match 'Site Id') {
        $VerbosePreference = $savedVerbose
        return $true
    }
    Write-Host "Get-MgReportSharePointSiteUsageDetail also produced invalid content." -ForegroundColor Yellow

    $VerbosePreference = $savedVerbose
    return $false
}

# Helper to normalize a URL so SPO and Graph values match reliably.
# Decodes percent-encoded characters, strips trailing slashes, query strings,
# and fragments, and lowercases the result.
function Normalize-SiteUrl {
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return $null }
    $n = [System.Uri]::UnescapeDataString($Url.Trim())
    $n = $n.TrimEnd('/')
    $qi = $n.IndexOf('?'); if ($qi -ge 0) { $n = $n.Substring(0, $qi) }
    $fi = $n.IndexOf('#'); if ($fi -ge 0) { $n = $n.Substring(0, $fi) }
    return $n.ToLowerInvariant()
}

# Resolve a SharePoint site URL to a Graph compound site ID.
#
# Tries two approaches in order:
#   1. Graph path-based addressing:
#      Get-MgSite -SiteId "hostname:/sites/Name:" → returns compound Id
#   2. Direct Graph REST:
#      Invoke-MgGraphRequest /v1.0/sites/{hostname}:/{path}: → returns id
#
# Returns $null if both fail.  The caller should log the failure.
function Resolve-GraphSiteId {
    param(
        [Parameter(Mandatory)][string]$SiteUrl
    )

    $u = [System.Uri]$SiteUrl
    $hostname = $u.Host
    $path = $u.AbsolutePath.TrimEnd('/')

    # Suppress verbose output from Graph API calls — when the user runs with
    # -Verbose these calls produce a line per HTTP request/response, which
    # generates hundreds of noisy VERBOSE lines during the resolution loop.
    $savedVerbose = $VerbosePreference
    $VerbosePreference = 'SilentlyContinue'

    # Approach 1: Get-MgSite with path-based addressing
    try {
        $pathBasedId = if ($path -and $path -ne '/') {
            "${hostname}:${path}:"
        } else {
            $hostname
        }
        $mgSite = Get-MgSite -SiteId $pathBasedId -Property "id" -ErrorAction Stop
        if ($mgSite -and $mgSite.Id) {
            $parts = $mgSite.Id -split ','
            if ($parts.Count -ge 2) {
                $VerbosePreference = $savedVerbose
                return [PSCustomObject]@{
                    CompoundId = $mgSite.Id
                    SiteGuid   = $parts[1]
                }
            }
        }
    }
    catch {
        # Will try fallback approach
    }

    # Approach 2: Invoke-MgGraphRequest directly (sometimes more reliable)
    try {
        $graphUri = if ($path -and $path -ne '/') {
            "/v1.0/sites/${hostname}:${path}:?`$select=id"
        } else {
            "/v1.0/sites/${hostname}?`$select=id"
        }
        $resp = Invoke-MgGraphRequest -Method GET -Uri $graphUri -ErrorAction Stop
        if ($resp.id) {
            $parts = $resp.id -split ','
            if ($parts.Count -ge 2) {
                $VerbosePreference = $savedVerbose
                return [PSCustomObject]@{
                    CompoundId = $resp.id
                    SiteGuid   = $parts[1]
                }
            }
        }
    }
    catch {
        # Both approaches failed
    }

    $VerbosePreference = $savedVerbose
    return $null
}

# Detect whether Graph report data has been obfuscated by the M365 privacy
# setting "Conceal user, group, and site names in all reports".  Obfuscated
# rows have zeroed-out SiteIds, empty/null Site URLs, and hashed owner names.
function Test-GraphDataObfuscated {
    param([array]$ReportData)
    if (-not $ReportData -or $ReportData.Count -eq 0) { return $false }
    $sampleSize = [Math]::Min(5, $ReportData.Count)
    for ($i = 0; $i -lt $sampleSize; $i++) {
        $entry = $ReportData[$i]
        if ([string]::IsNullOrWhiteSpace($entry.SiteUrl) -or
            $entry.SiteId -eq '00000000-0000-0000-0000-000000000000' -or
            ($entry.OwnerDisplayName -match '^[A-Fa-f0-9]{32}$')) {
            return $true
        }
    }
    return $false
}

# Check the tenant admin report-privacy setting using the Graph PowerShell SDK
# cmdlets (Get-MgAdminReportSetting / Update-MgAdminReportSetting) and, if
# concealment is enabled, attempt to disable it so that future reports contain
# real identifiers.  Returns a hashtable:
#   WasEnabled - $true/$false/$null (null = could not read the setting)
#   Fixed      - $true if the setting is now disabled
function Resolve-GraphReportPrivacy {
    try {
        # Use the typed Graph SDK cmdlet — more reliable than raw Invoke-MgGraphRequest.
        $reportSetting = Get-MgAdminReportSetting -ErrorAction Stop
        $concealed = $reportSetting.DisplayConcealedNames

        Write-Host "Current DisplayConcealedNames value: $concealed" -ForegroundColor Cyan

        if ($concealed -eq $true) {
            Write-Warning "Report privacy setting 'Conceal user, group, and site names in all reports' is ENABLED in your tenant (DisplayConcealedNames = True)."
            Write-Host "Attempting to disable the concealment setting via Update-MgAdminReportSetting..." -ForegroundColor Yellow
            try {
                $null = Update-MgAdminReportSetting -DisplayConcealedNames:$false -ErrorAction Stop
                Write-Host "Concealment setting has been disabled (DisplayConcealedNames set to False)." -ForegroundColor Green
                Write-Host "Note: report data may take up to 48 hours to reflect this change." -ForegroundColor Yellow
                Write-Host "Re-run this script after the propagation period for fully de-obfuscated data." -ForegroundColor Yellow
                return @{ WasEnabled = $true; Fixed = $true }
            }
            catch {
                Write-Warning "Could not disable the concealment setting (requires ReportSettings.ReadWrite.All permission): $_"
                return @{ WasEnabled = $true; Fixed = $false }
            }
        }
        else {
            Write-Host "Report concealment setting is already disabled (DisplayConcealedNames = False)." -ForegroundColor Green
            return @{ WasEnabled = $false; Fixed = $true }
        }
    }
    catch {
        Write-Warning "Could not read report privacy settings via Get-MgAdminReportSetting (requires ReportSettings.Read.All permission): $_"
        return @{ WasEnabled = $null; Fixed = $false }
    }
}

# Retrieve per-site analytics and drive information via the Microsoft Graph
# Sites API.  These endpoints are NOT subject to the report-privacy concealment
# setting, so they always return real data.  For each site the function queries:
#   /sites/{id}/analytics/lastSevenDays  → access.actionCount (page views)
#   /sites/{id}/drive                    → quota/driveItem info
# Sites for which the call fails (permissions, missing drive, etc.) are silently
# skipped.  Returns a hashtable mapping webUrl → enrichment object.
function Get-PerSiteAnalyticsViaGraph {
    param([array]$Sites)

    $analytics = @{}
    $total = $Sites.Count
    $counter = 0

    Write-Host "Retrieving per-site analytics for $total sites (this may take a moment)..." -ForegroundColor Cyan

    # Suppress verbose output from Graph API calls to avoid hundreds of
    # VERBOSE lines when the user runs with -Verbose.
    $savedVerbose = $VerbosePreference
    $VerbosePreference = 'SilentlyContinue'

    foreach ($site in $Sites) {
        $counter++
        $siteId  = $site.id
        $webUrl  = $site.webUrl
        if (-not $siteId) { continue }

        Write-Progress -Activity "Retrieving per-site analytics" `
            -Status "Processing $counter of $total" `
            -PercentComplete (($counter / $total) * 100)

        $info = [PSCustomObject]@{
            PageViewCount    = $null
            FileCount        = $null
            StorageUsedBytes = $null
            LastActivityDate = $null
        }

        # --- Site analytics (page views) ---
        try {
            $resp = Invoke-MgGraphRequest -Method GET `
                -Uri "/v1.0/sites/$siteId/analytics/lastSevenDays" `
                -ErrorAction Stop
            if ($resp.access) {
                $info.PageViewCount = $resp.access.actionCount
            }
        }
        catch {
            # Analytics not available for this site — continue silently
        }

        # --- Drive info (file count + storage) ---
        try {
            $driveResp = Invoke-MgGraphRequest -Method GET `
                -Uri "/v1.0/sites/$siteId/drive?`$select=quota,lastModifiedDateTime" `
                -ErrorAction Stop
            if ($driveResp.quota) {
                $info.StorageUsedBytes = $driveResp.quota.used
            }
            if ($driveResp.lastModifiedDateTime) {
                $info.LastActivityDate = $driveResp.lastModifiedDateTime
            }
        }
        catch {
            # Drive not available for this site — continue silently
        }

        # --- File count via root children count (closest approximation to the
        #     report's FileCount; includes both files and sub-folders at root) ---
        try {
            $rootResp = Invoke-MgGraphRequest -Method GET `
                -Uri "/v1.0/sites/$siteId/drive/root?`$select=folder" `
                -ErrorAction Stop
            if ($rootResp.folder) {
                $info.FileCount = $rootResp.folder.childCount
            }
        }
        catch {
            # Root folder not available — continue silently
        }

        if ($webUrl) {
            $analytics[$webUrl] = $info
        }
    }

    Write-Progress -Activity "Retrieving per-site analytics" -Completed
    $VerbosePreference = $savedVerbose
    $populated = ($analytics.Values | Where-Object { $null -ne $_.PageViewCount }).Count
    Write-Host "Retrieved analytics for $populated of $total sites." -ForegroundColor Green

    return $analytics
}

# Retrieve real site metadata via the Microsoft Graph Sites API.  This endpoint
# is not subject to the report-privacy concealment setting, so it always returns
# real site IDs, URLs, and display names.
function Get-SiteMetadataViaGraph {
    param([string]$TenantName)

    Write-Host "Retrieving real site metadata from Microsoft Graph Sites API..." -ForegroundColor Cyan

    $allSites = @()

    # Suppress verbose output from Graph API calls
    $savedVerbose = $VerbosePreference
    $VerbosePreference = 'SilentlyContinue'

    # Try getAllSites first (works with Sites.Read.All application permission)
    $uri = '/v1.0/sites/getAllSites?$select=id,displayName,webUrl,createdDateTime&$top=999'
    try {
        while ($uri) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($response.value) { $allSites += $response.value }
            $uri = $response.'@odata.nextLink'
        }
        Write-Host "Retrieved metadata for $($allSites.Count) sites from Sites API." -ForegroundColor Green
        $VerbosePreference = $savedVerbose
        return $allSites
    }
    catch {
        Write-Warning "getAllSites endpoint failed, trying search-based approach: $_"
    }

    # Fallback: search-based approach (works with delegated Sites.Read.All).
    # Use the tenant name as the search keyword — matches most SharePoint sites
    # since they live under the tenant's *.sharepoint.com domain.
    $searchTerm = if ($TenantName) { $TenantName } else { 'sharepoint' }
    $uri = "/v1.0/sites?search=$searchTerm&`$select=id,displayName,webUrl,createdDateTime&`$top=999"
    try {
        while ($uri) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            if ($response.value) { $allSites += $response.value }
            $uri = $response.'@odata.nextLink'
        }
        Write-Host "Retrieved metadata for $($allSites.Count) sites from Sites API (search)." -ForegroundColor Green
        $VerbosePreference = $savedVerbose
        return $allSites
    }
    catch {
        Write-Warning "Sites search API also failed: $_"
        $VerbosePreference = $savedVerbose
        return @()
    }
}

# Function to combine SPO and Graph data for best-of-both-worlds reporting.
#
# Strategy (SPO-first, per ChatGPT analysis):
#   SPO is the source of truth for site inventory — Get-SPOSite reliably returns
#   Title, URL, and Owner even when the Graph Reports API is obfuscated.  We then
#   attach Graph usage metrics (page views, file counts, activity dates) to each
#   SPO site by matching on URL or by resolving the SPO URL to a Graph site ID
#   via path-based addressing (/sites/{hostname}:/{path}:).
#
#   This avoids getAllSites (requires app-only permission) and doesn't depend on
#   the Graph report having usable URLs or SiteIds.
#
# Matching order for each SPO site → Graph usage row:
#   1. URL match — normalised SiteUrl from both sides.
#   2. SiteId match — resolve SPO URL → Graph compound site ID via
#      Get-MgSite -SiteId "{hostname}:/{path}:", extract the site GUID,
#      and match it to the Graph report's "Site Id" field.
function Get-UsageReportsCombined {
    param(
        [string]$TenantName
    )

    Write-Host "Running combined mode (SPO-first): SPO for site inventory, Graph for usage metrics..." -ForegroundColor Cyan

    # ── Step 1: Get SPO data — source of truth for site inventory ──
    $spoData = Get-UsageReportsViaSPO -TenantName $TenantName
    Write-Host "SPO returned $($spoData.Count) site records (Title, URL, Owner)." -ForegroundColor Cyan

    # Validate that SPO data has expected properties
    if ($spoData.Count -gt 0) {
        $sampleSpo = $spoData[0]
        Write-Verbose "SPO sample — SiteUrl: '$($sampleSpo.SiteUrl)', Title: '$($sampleSpo.Title)', Owner: '$($sampleSpo.Owner)'"
        $spoWithUrls = ($spoData | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SiteUrl) }).Count
        if ($spoWithUrls -lt $spoData.Count) {
            Write-Warning "$($spoData.Count - $spoWithUrls) SPO sites have blank URLs."
        }
    }

    # ── Step 2: Connect to Graph and pull the usage report ──
    Write-Host "Connecting to Microsoft Graph for usage metrics..." -ForegroundColor Cyan
    try {
        $null = Connect-MgGraph -Scopes "Reports.Read.All", "Sites.Read.All", "ReportSettings.ReadWrite.All" -NoWelcome
    }
    catch {
        Write-Warning "Could not connect to Microsoft Graph: $_"
        Write-Warning "Combined report will contain SPO metadata only (no page views or activity data)."
        # Return SPO-only data with empty Graph columns
        return Build-SPOOnlyOutput -SpoData $spoData
    }

    # Validate that we have the expected permissions
    try {
        $context = Get-MgContext
        if ($context) {
            Write-Host "Graph context: Account=$($context.Account), Scopes=$($context.Scopes -join ', ')" -ForegroundColor Cyan
        }
    }
    catch {
        # Non-fatal — just skip the permission dump
    }

    # Proactively check/fix the report-privacy concealment setting
    $privacyResult = Resolve-GraphReportPrivacy

    # Pull the usage report via the robust download helper, which tries
    # Invoke-MgGraphRequest first (avoids the PercentComplete overflow bug)
    # and falls back to Get-MgReportSharePointSiteUsageDetail with progress
    # suppression.  Both approaches validate that the output is real CSV.
    Write-Host "Retrieving SharePoint site usage report from Microsoft Graph..." -ForegroundColor Cyan
    $graphData = @()
    $tempFile = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName() + ".csv")
    try {
        $downloadOk = Get-GraphReportCsv -TempFile $tempFile
        if (-not $downloadOk) {
            throw "Both report download methods failed or produced invalid content."
        }

        $sites = Import-Csv -Path $tempFile
        foreach ($site in $sites) {
            $graphData += [PSCustomObject]@{
                SiteUrl              = $site.'Site URL'
                SiteId               = $site.'Site Id'
                OwnerDisplayName     = $site.'Owner Display Name'
                OwnerPrincipalName   = $site.'Owner Principal Name'
                IsDeleted            = $site.'Is Deleted'
                LastActivityDate     = $site.'Last Activity Date'
                FileCount            = $site.'File Count'
                ActiveFileCount      = $site.'Active File Count'
                PageViewCount        = $site.'Page View Count'
                VisitedPageCount     = $site.'Visited Page Count'
                StorageUsedInBytes   = $site.'Storage Used (Byte)'
                StorageAllocatedInBytes = $site.'Storage Allocated (Byte)'
                RootWebTemplate      = $site.'Root Web Template'
                ReportRefreshDate    = $site.'Report Refresh Date'
                ReportPeriod         = $site.'Report Period'
            }
        }
        Write-Host "Graph returned $($graphData.Count) site usage records." -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not retrieve Graph usage report: $_"
        Write-Warning "Combined report will contain SPO metadata only."
        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        return Build-SPOOnlyOutput -SpoData $spoData
    }
    finally {
        if ($tempFile) { Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue }
    }

    # ── Step 3: Build Graph lookups ──
    $emptyGuid = '00000000-0000-0000-0000-000000000000'

    # 3a: Lookup by normalised URL (works when Graph report URLs are populated)
    $graphUrlLookup = @{}
    foreach ($g in $graphData) {
        $key = Normalize-SiteUrl -Url $g.SiteUrl
        if ($key) { $graphUrlLookup[$key] = $g }
    }

    # 3b: Lookup by simple SiteId GUID (works even when Graph URLs are blank)
    $graphIdLookup = @{}
    foreach ($g in $graphData) {
        if ($g.SiteId -and $g.SiteId -ne $emptyGuid) {
            $simpleId = ($g.SiteId -split ',')[-1]  # handles both simple and compound IDs
            $graphIdLookup[$simpleId] = $g
        }
    }

    $graphUrlCount = $graphUrlLookup.Count
    Write-Host "Graph lookup: $graphUrlCount by URL, $($graphIdLookup.Count) by SiteId." -ForegroundColor Cyan

    # ── Step 4: Resolve SPO URLs to Graph SiteIds ──
    # Try two approaches: Get-MgSite path-based addressing and Invoke-MgGraphRequest.
    # Log the first few failures as warnings so the user can diagnose (403 vs 404).
    Write-Host "Resolving SPO URLs to Graph SiteIds..." -ForegroundColor Cyan
    $spoSiteIdMap = @{}    # SPO URL → simple site GUID from Graph
    $spoCompoundMap = @{}  # SPO URL → full compound Id (for per-site analytics fallback)
    $resolveCounter = 0
    $resolveTotal = $spoData.Count
    $failureCount = 0
    $maxResolutionWarnings = 5  # Show first N failures as warnings for diagnostics

    foreach ($spoSite in $spoData) {
        $resolveCounter++
        if ($resolveCounter % 10 -eq 0 -or $resolveCounter -eq $resolveTotal) {
            Write-Progress -Activity "Resolving SPO URLs to Graph SiteIds" `
                -Status "$resolveCounter of $resolveTotal" `
                -PercentComplete (($resolveCounter / $resolveTotal) * 100)
        }

        $resolved = Resolve-GraphSiteId -SiteUrl $spoSite.SiteUrl
        if ($resolved) {
            $spoSiteIdMap[$spoSite.SiteUrl] = $resolved.SiteGuid
            $spoCompoundMap[$spoSite.SiteUrl] = $resolved.CompoundId
        }
        else {
            $failureCount++
            if ($failureCount -le $maxResolutionWarnings) {
                Write-Warning "Could not resolve Graph SiteId for: $($spoSite.SiteUrl)"
            }
        }
    }
    Write-Progress -Activity "Resolving SPO URLs to Graph SiteIds" -Completed

    if ($failureCount -gt $maxResolutionWarnings) {
        Write-Warning "... and $($failureCount - $maxResolutionWarnings) more sites could not be resolved. Run with -Verbose for details."
    }
    Write-Host "Resolved $($spoSiteIdMap.Count) of $resolveTotal SPO sites to Graph SiteIds ($failureCount failed)." -ForegroundColor Green

    # ── Step 5: Try to match SPO → Graph report data ──
    $matchedByUrl = 0
    $matchedById = 0

    # Build per-SPO-site match results
    $spoGraphMatch = @{}  # SPO URL → Graph report row (or $null)

    foreach ($spoSite in $spoData) {
        $graphSite = $null

        # Try 1: match by normalised URL
        $spoKey = Normalize-SiteUrl -Url $spoSite.SiteUrl
        if ($spoKey -and $graphUrlLookup.ContainsKey($spoKey)) {
            $graphSite = $graphUrlLookup[$spoKey]
            $matchedByUrl++
        }

        # Try 2: match by resolved SiteId
        if (-not $graphSite) {
            $resolvedId = $spoSiteIdMap[$spoSite.SiteUrl]
            if ($resolvedId -and $graphIdLookup.ContainsKey($resolvedId)) {
                $graphSite = $graphIdLookup[$resolvedId]
                $matchedById++
            }
        }

        $spoGraphMatch[$spoSite.SiteUrl] = $graphSite
    }

    $totalMatched = $matchedByUrl + $matchedById
    Write-Host "Report matching: $matchedByUrl by URL, $matchedById by SiteId, $($spoData.Count - $totalMatched) unmatched." -ForegroundColor Cyan

    # ── Step 6: If report join mostly failed, fall back to per-site analytics ──
    # When the report is still effectively concealed (zeroed SiteIds, blank URLs),
    # the join will fail.  In that case, use Graph per-site analytics (which are
    # NOT subject to report concealment) with the compound site IDs we resolved.
    $usePerSiteAnalytics = $false
    $perSiteAnalytics = @{}
    $matchRateThreshold = 0.2  # Fall back to per-site analytics if less than 20% of sites matched

    if ($totalMatched -lt ($spoData.Count * $matchRateThreshold) -and $spoCompoundMap.Count -gt 0) {
        Write-Warning "Report join mostly failed ($totalMatched of $($spoData.Count) matched). Graph report data appears to be obfuscated."
        Write-Host "Falling back to per-site analytics via Graph (not subject to report concealment)..." -ForegroundColor Yellow

        # Build a sites array for Get-PerSiteAnalyticsViaGraph from the resolved compound IDs
        $analyticsInput = @()
        foreach ($url in $spoCompoundMap.Keys) {
            $analyticsInput += [PSCustomObject]@{
                id     = $spoCompoundMap[$url]
                webUrl = $url
            }
        }

        if ($analyticsInput.Count -gt 0) {
            $perSiteAnalytics = Get-PerSiteAnalyticsViaGraph -Sites $analyticsInput
            $usePerSiteAnalytics = $true
            Write-Host "Retrieved per-site analytics for $($perSiteAnalytics.Count) sites." -ForegroundColor Green
        }
    }

    # ── Step 7: Build final output ──
    $combinedData = @()

    Write-Host "Building combined report for $($spoData.Count) SPO sites..." -ForegroundColor Cyan

    foreach ($spoSite in $spoData) {
        $graphSite = $spoGraphMatch[$spoSite.SiteUrl]

        # If we're using per-site analytics as fallback, get the enrichment data
        $enrichment = $null
        if ($usePerSiteAnalytics -and -not $graphSite) {
            $enrichment = if ($spoSite.SiteUrl) { $perSiteAnalytics[$spoSite.SiteUrl] } else { $null }
        }

        # Build combined row: SPO metadata + Graph usage metrics (from report or per-site analytics)
        $combined = [PSCustomObject]@{
            SiteUrl                 = $spoSite.SiteUrl
            Title                   = $spoSite.Title
            Owner                   = $spoSite.Owner
            Template                = $spoSite.Template
            StorageUsedMB           = $spoSite.StorageUsedMB
            StorageQuotaMB          = $spoSite.StorageQuotaMB
            StorageUsedPercentage   = $spoSite.StorageUsedPercentage
            LastContentModifiedDate = $spoSite.LastContentModifiedDate
            LastActivityDate        = if ($graphSite) { $graphSite.LastActivityDate } `
                                      elseif ($enrichment -and $null -ne $enrichment.LastActivityDate) { $enrichment.LastActivityDate } `
                                      else { '' }
            FileCount               = if ($graphSite) { $graphSite.FileCount } `
                                      elseif ($enrichment -and $null -ne $enrichment.FileCount) { $enrichment.FileCount } `
                                      else { '' }
            ActiveFileCount         = if ($graphSite) { $graphSite.ActiveFileCount } else { '' }
            PageViewCount           = if ($graphSite) { $graphSite.PageViewCount } `
                                      elseif ($enrichment -and $null -ne $enrichment.PageViewCount) { $enrichment.PageViewCount } `
                                      else { '' }
            VisitedPageCount        = if ($graphSite) { $graphSite.VisitedPageCount } else { '' }
            SharingCapability       = $spoSite.SharingCapability
            LockState               = $spoSite.LockState
            IsHubSite               = $spoSite.IsHubSite
            HubSiteId               = $spoSite.HubSiteId
            SensitivityLabel        = $spoSite.SensitivityLabel
            RootWebTemplate         = if ($graphSite) { $graphSite.RootWebTemplate } else { '' }
            IsDeleted               = if ($graphSite) { $graphSite.IsDeleted } else { '' }
            CreatedDate             = $spoSite.CreatedDate
            ReportRefreshDate       = if ($graphSite) { $graphSite.ReportRefreshDate } else { '' }
        }
        $combinedData += $combined
    }

    $analyticsCount = if ($usePerSiteAnalytics) { ($perSiteAnalytics.Values | Where-Object { $null -ne $_.PageViewCount }).Count } else { 0 }
    Write-Host "Combined report: $($combinedData.Count) sites — $matchedByUrl matched by URL, $matchedById matched by SiteId" -ForegroundColor Green -NoNewline
    if ($usePerSiteAnalytics) {
        Write-Host ", $analyticsCount enriched via per-site analytics" -ForegroundColor Green -NoNewline
    }
    Write-Host "." -ForegroundColor Green

    if ($totalMatched -eq 0 -and -not $usePerSiteAnalytics -and $graphData.Count -gt 0) {
        Write-Warning "No SPO sites matched Graph usage data and per-site analytics was not available."
        Write-Warning "Activity columns (PageViewCount, FileCount, etc.) will be empty."
    }

    # Validate final output has expected data
    if ($combinedData.Count -gt 0) {
        $withUrls = ($combinedData | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SiteUrl) }).Count
        $withTitles = ($combinedData | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Title) }).Count
        Write-Host "Output validation: $withUrls/$($combinedData.Count) have URLs, $withTitles/$($combinedData.Count) have Titles." -ForegroundColor Cyan
        if ($withUrls -eq 0) {
            Write-Warning "All SiteUrl values are blank. This indicates SPO data was not properly loaded or is missing URLs."
        }
    }

    # Disconnect Graph
    $null = Disconnect-MgGraph -ErrorAction SilentlyContinue

    return $combinedData
}

# Helper to return SPO-only output with empty Graph columns when Graph is unavailable.
function Build-SPOOnlyOutput {
    param([array]$SpoData)
    $output = @()
    foreach ($spoSite in $SpoData) {
        $output += [PSCustomObject]@{
            SiteUrl                 = $spoSite.SiteUrl
            Title                   = $spoSite.Title
            Owner                   = $spoSite.Owner
            Template                = $spoSite.Template
            StorageUsedMB           = $spoSite.StorageUsedMB
            StorageQuotaMB          = $spoSite.StorageQuotaMB
            StorageUsedPercentage   = $spoSite.StorageUsedPercentage
            LastContentModifiedDate = $spoSite.LastContentModifiedDate
            LastActivityDate        = ''
            FileCount               = ''
            ActiveFileCount         = ''
            PageViewCount           = ''
            VisitedPageCount        = ''
            SharingCapability       = $spoSite.SharingCapability
            LockState               = $spoSite.LockState
            IsHubSite               = $spoSite.IsHubSite
            HubSiteId               = $spoSite.HubSiteId
            SensitivityLabel        = $spoSite.SensitivityLabel
            RootWebTemplate         = ''
            IsDeleted               = ''
            CreatedDate             = $spoSite.CreatedDate
            ReportRefreshDate       = ''
        }
    }
    return $output
}

# Main script execution
try {
    Write-Host "`n=== SharePoint Site Usage Report Generator ===" -ForegroundColor Cyan
    Write-Host "Tenant: $TenantName" -ForegroundColor Cyan
    Write-Host "Authentication Method: $AuthMethod" -ForegroundColor Cyan
    Write-Host "Mode: $(if ($UseCombined) { 'Combined (SPO + Graph)' } elseif ($UseGraphAPI) { 'Graph API' } else { 'SPO Management Shell' })`n" -ForegroundColor Cyan
    
    # Set default output path if not provided
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $OutputPath = Join-Path -Path $PSScriptRoot -ChildPath "SPO_SiteUsage_$($TenantName)_$timestamp.csv"
    }
    
    # Check and install required modules
    if ($UseCombined) {
        Write-Host "Checking for required modules (SPO + Graph)..." -ForegroundColor Cyan
        $modulesInstalled = Install-RequiredModules -ModuleNames @('Microsoft.Online.SharePoint.PowerShell', 'Microsoft.Graph.Reports', 'Microsoft.Graph.Authentication')
        if (-not $modulesInstalled) {
            throw "Failed to install required modules for combined mode."
        }

        # Get usage data via combined SPO + Graph approach
        $usageData = Get-UsageReportsCombined -TenantName $TenantName
    }
    elseif ($UseGraphAPI) {
        Write-Host "Checking for required Microsoft Graph modules..." -ForegroundColor Cyan
        $modulesInstalled = Install-RequiredModules -ModuleNames @('Microsoft.Graph.Reports', 'Microsoft.Graph.Authentication')
        if (-not $modulesInstalled) {
            throw "Failed to install required Microsoft Graph modules."
        }
        
        # Get usage data via Graph API
        $usageData = Get-UsageReportsViaGraph -TenantName $TenantName
    }
    else {
        Write-Host "Checking for required SharePoint Online Management Shell module..." -ForegroundColor Cyan
        $modulesInstalled = Install-RequiredModules -ModuleNames @('Microsoft.Online.SharePoint.PowerShell')
        if (-not $modulesInstalled) {
            throw "Failed to install required SharePoint Online Management Shell module."
        }
        
        # Get usage data via SPO Management Shell
        $usageData = Get-UsageReportsViaSPO -TenantName $TenantName
    }
    
    # Export to CSV
    if ($usageData -and $usageData.Count -gt 0) {
        Write-Host "`nExporting data to CSV: $OutputPath" -ForegroundColor Cyan

        # Ensure the output directory exists
        $outputDir = Split-Path -Path $OutputPath -Parent
        if ($outputDir -and -not (Test-Path -Path $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        }

        # Attempt to export with retry logic in case the file is locked
        $maxRetries = 3
        $retryDelay = 2
        $exported = $false
        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                $usageData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
                $exported = $true
                break
            }
            catch [System.IO.IOException] {
                if ($attempt -lt $maxRetries) {
                    Write-Warning "File '$OutputPath' is in use. Retrying in $retryDelay seconds... (attempt $attempt of $maxRetries)"
                    Start-Sleep -Seconds $retryDelay
                    $retryDelay *= 2
                }
            }
        }

        if (-not $exported) {
            # All retries failed — write to a fallback file with a timestamp suffix
            $fallbackName = [System.IO.Path]::GetFileNameWithoutExtension($OutputPath) +
                "_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv"
            $fallbackDir = Split-Path -Path $OutputPath -Parent
            $fallbackPath = if ($fallbackDir) { Join-Path -Path $fallbackDir -ChildPath $fallbackName } else { $fallbackName }
            Write-Warning "Could not write to '$OutputPath' because it is locked by another process (e.g., Excel)."
            Write-Warning "Saving report to fallback location: $fallbackPath"
            $usageData | Export-Csv -Path $fallbackPath -NoTypeInformation -Encoding UTF8
            $OutputPath = $fallbackPath
        }

        Write-Host "Export completed successfully!" -ForegroundColor Green
        Write-Host "Total sites exported: $($usageData.Count)" -ForegroundColor Green
        Write-Host "File location: $OutputPath" -ForegroundColor Green
        
        # Display summary statistics
        if ($UseCombined) {
            $totalStorageMB = ($usageData | Measure-Object -Property StorageUsedMB -Sum).Sum
            $totalFiles = ($usageData | Where-Object { $_.FileCount -ne '' } | Measure-Object -Property FileCount -Sum).Sum
            $totalPageViews = ($usageData | Where-Object { $_.PageViewCount -ne '' } | Measure-Object -Property PageViewCount -Sum).Sum
            
            Write-Host "`n=== Summary Statistics ===" -ForegroundColor Cyan
            Write-Host "Total Storage Used: $([math]::Round($totalStorageMB / 1024, 2)) GB" -ForegroundColor White
            Write-Host "Total Files: $totalFiles" -ForegroundColor White
            Write-Host "Total Page Views (last 7 days): $totalPageViews" -ForegroundColor White
        }
        elseif ($UseGraphAPI) {
            $totalStorage = ($usageData | Measure-Object -Property StorageUsedInBytes -Sum).Sum
            $totalFiles = ($usageData | Measure-Object -Property FileCount -Sum).Sum
            $totalPageViews = ($usageData | Measure-Object -Property PageViewCount -Sum).Sum
            
            Write-Host "`n=== Summary Statistics ===" -ForegroundColor Cyan
            Write-Host "Total Storage Used: $([math]::Round($totalStorage / 1GB, 2)) GB" -ForegroundColor White
            Write-Host "Total Files: $totalFiles" -ForegroundColor White
            Write-Host "Total Page Views: $totalPageViews" -ForegroundColor White
        }
        else {
            $totalStorageMB = ($usageData | Measure-Object -Property StorageUsedMB -Sum).Sum
            Write-Host "`n=== Summary Statistics ===" -ForegroundColor Cyan
            Write-Host "Total Storage Used: $([math]::Round($totalStorageMB / 1024, 2)) GB" -ForegroundColor White
        }
    }
    else {
        Write-Warning "No usage data retrieved. No CSV file created."
    }
}
catch {
    Write-Error "An error occurred: $_"
    Write-Error $_.ScriptStackTrace
    exit 1
}
