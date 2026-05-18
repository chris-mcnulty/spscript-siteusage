<#
.SYNOPSIS
Inventory sensitivity labels across an M365 tenant:
 - Site container label (SensitivityLabel) for every SharePoint site (incl OneDrive + channel sites)
 - Default library sensitivity label (DefaultSensitivityLabelForLibrary) for every document library in each site
Writes streaming CSV output (safe for large tenants), includes retry/backoff, and supports resume.

WHY TWO MODULES?
- PnP bulk retrieval of site SensitivityLabel can be empty (known issue). Workaround: use Get-SPOSite. 【1-8a8ff5】【2-5e5285】【3-63f690】

AUTH
- Uses app-only certificate auth for BOTH:
  - Connect-SPOService (SPO Management Shell supports certificate auth parameters) 【8-95166b】
  - Connect-PnPOnline (PnP app-only cert)

REQUIREMENTS
- Microsoft.Online.SharePoint.PowerShell module
- PnP.PowerShell module
- An Entra ID app registration with appropriate SharePoint permissions (Sites.FullControl.All is typical for app-only SPO) 【8-95166b】
- Certificate (PFX) associated with the app

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantName,                 # e.g. "contoso" (used to build admin URL)

    [Parameter(Mandatory = $true)]
    [string]$ClientId,                   # Entra App (Application/Client) ID

    [Parameter(Mandatory = $true)]
    [string]$TenantId,                   # Entra Tenant ID (GUID)

    [Parameter(Mandatory = $true)]
    [string]$CertificatePath,            # Path to .pfx

    [Parameter(Mandatory = $true)]
    [SecureString]$CertificatePassword,  # SecureString password for .pfx

    [Parameter(Mandatory = $false)]
    [string]$OutputCsvPath = ".\M365_SensitivityLabel_Inventory.csv",

    [Parameter(Mandatory = $false)]
    [string]$ErrorCsvPath  = ".\M365_SensitivityLabel_Inventory_Errors.csv",

    # Default ON: include OneDrive personal sites (required for “include OneDrive by default”) 【4-bd0be9】
    [Parameter(Mandatory = $false)]
    [bool]$IncludeOneDrivePersonalSites = $true,

    # Default OFF: include hidden libraries (normally you don’t want them)
    [Parameter(Mandatory = $false)]
    [bool]$IncludeHiddenLibraries = $false,

    # Optional: Filter sites by URL contains (cheap scoping for test runs)
    [Parameter(Mandatory = $false)]
    [string]$UrlContainsFilter = "",

    # Resume: if CSV exists, skip rows already written
    [Parameter(Mandatory = $false)]
    [switch]$Resume,

    # Optional: resolve label IDs to names (best effort). Requires InfoProtectionPolicy permissions in Graph context. 【7-d07106】
    [Parameter(Mandatory = $false)]
    [bool]$ResolveLabelNames = $true,

    # Throttling safety: small delay between sites to reduce SPO throttling risk
    [Parameter(Mandatory = $false)]
    [int]$InterSiteDelayMs = 100
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --------------------------
# Helpers: CSV streaming + retries
# --------------------------
function Ensure-CsvHeader {
    param(
        [string]$Path,
        [string[]]$Headers
    )
    if (-not (Test-Path $Path)) {
        ($Headers -join ",") | Out-File -FilePath $Path -Encoding utf8
    }
}

function Append-CsvRow {
    param(
        [string]$Path,
        [hashtable]$Row,
        [string[]]$Headers
    )
    # Maintain column order and escape quotes
    $values = foreach ($h in $Headers) {
        $v = ""
        if ($Row.ContainsKey($h) -and $null -ne $Row[$h]) { $v = [string]$Row[$h] }
        $v = $v.Replace('"','""')
        '"' + $v + '"'
    }
    ($values -join ",") | Add-Content -Path $Path -Encoding utf8
}

function Invoke-WithRetry {
    param(
        [scriptblock]$Operation,
        [int]$MaxRetries = 6,
        [int]$InitialDelaySeconds = 2
    )

    $delay = $InitialDelaySeconds
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            return & $Operation
        }
        catch {
            if ($attempt -eq $MaxRetries) { throw }
            Start-Sleep -Seconds $delay
            $delay = [Math]::Min($delay * 2, 60)
        }
    }
}

# --------------------------
# Modules
# --------------------------
Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
#Import-Module PnP.PowerShell -ErrorAction Stop

$adminUrl = "https://$TenantName-admin.sharepoint.com"

# --------------------------
# Connect: SPO (site labels)
# Connect-SPOService supports certificate-based auth parameters (ClientId, TenantId, CertificatePath/Password). 【8-95166b】
# --------------------------
Write-Host "Connecting to SPO admin: $adminUrl" -ForegroundColor Cyan
Invoke-WithRetry -Operation {
    Connect-SPOService -Url $adminUrl -ClientId $ClientId -TenantId $TenantId -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword
}

# --------------------------
# Connect: PnP admin (libraries + optional label-name lookup)
# --------------------------
Write-Host "Connecting to PnP admin: $adminUrl" -ForegroundColor Cyan
$adminConn = Invoke-WithRetry -Operation {
    Connect-PnPOnline -Url $adminUrl -ClientId $ClientId -Tenant $TenantId -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -ReturnConnection
}

# --------------------------
# Label name lookup (best effort)
# Get-PnPAvailableSensitivityLabel is the supported way to list tenant labels in PnP. 【7-d07106】
# --------------------------
$labelLookup = @{}
if ($ResolveLabelNames) {
    try {
        $labels = Get-PnPAvailableSensitivityLabel -Connection $adminConn
        foreach ($l in $labels) {
            $id = $l.Id.ToString().ToLower()
            $name = $null
            if ($l.PSObject.Properties.Name -contains "DisplayName") { $name = $l.DisplayName }
            if (-not $name -and ($l.PSObject.Properties.Name -contains "Name")) { $name = $l.Name }
            if ($name) { $labelLookup[$id] = $name }
        }
    }
    catch {
        Write-Host "Warning: Could not read label catalog via Get-PnPAvailableSensitivityLabel; exporting IDs only." -ForegroundColor Yellow
    }
}

function Resolve-LabelName {
    param([string]$LabelId)
    if ([string]::IsNullOrEmpty($LabelId)) { return "" }
    $k = $LabelId.ToLower()
    if ($labelLookup.ContainsKey($k)) { return $labelLookup[$k] }
    return ""
}

# --------------------------
# Output headers
# --------------------------
$headers = @(
    "Timestamp",
    "Scope",                 # Site | Library
    "SiteUrl",
    "SiteTitle",
    "SiteTemplate",
    "SiteSensitivityLabelId",
    "SiteSensitivityLabelName",
    "LibraryTitle",
    "LibraryId",
    "LibraryServerRelativeUrl",
    "LibraryDefaultLabelId",
    "LibraryDefaultLabelName",
    "Notes"
)

$errorHeaders = @(
    "Timestamp","Scope","SiteUrl","LibraryTitle","Operation","Error"
)

Ensure-CsvHeader -Path $OutputCsvPath -Headers $headers
Ensure-CsvHeader -Path $ErrorCsvPath  -Headers $errorHeaders

# --------------------------
# Resume support: build a set of already-exported keys
# Key = Scope|SiteUrl|LibraryId (LibraryId empty for Site rows)
# --------------------------
$exportedKeys = New-Object 'System.Collections.Generic.HashSet[string]'
if ($Resume -and (Test-Path $OutputCsvPath)) {
    Write-Host "Resume enabled: loading existing keys from $OutputCsvPath ..." -ForegroundColor Cyan
    try {
        $existing = Import-Csv -Path $OutputCsvPath
        foreach ($r in $existing) {
            $k = "{0}|{1}|{2}" -f $r.Scope, $r.SiteUrl, $r.LibraryId
            [void]$exportedKeys.Add($k)
        }
        Write-Host ("Loaded {0} exported keys" -f $exportedKeys.Count) -ForegroundColor Cyan
    }
    catch {
        Write-Host "Warning: could not parse existing CSV for resume; continuing without skip set." -ForegroundColor Yellow
    }
}

# --------------------------
# Get sites (includes OneDrive personal sites by default) 【4-bd0be9】
# SiteSensitivityLabel comes from Get-SPOSite output (and is called out as included in internal usage script notes). 【5-dc078a】
# --------------------------
Write-Host "Retrieving sites via Get-SPOSite ..." -ForegroundColor Cyan
$sites = Invoke-WithRetry -Operation {
    Get-SPOSite -Limit ALL -IncludePersonalSite $IncludeOneDrivePersonalSites
}

if (-not [string]::IsNullOrEmpty($UrlContainsFilter)) {
    $sites = $sites | Where-Object { $_.Url -like "*$UrlContainsFilter*" }
}

Write-Host ("Sites retrieved: {0}" -f ($sites | Measure-Object).Count) -ForegroundColor Cyan

# --------------------------
# Main loop
# --------------------------
foreach ($s in $sites) {
    $siteUrl = $s.Url
    $siteTitle = ""
    if ($s.PSObject.Properties.Name -contains "Title") { $siteTitle = $s.Title }

    $siteTemplate = ""
    if ($s.PSObject.Properties.Name -contains "Template") { $siteTemplate = $s.Template }

    $siteLabelId = ""
    if ($s.PSObject.Properties.Name -contains "SensitivityLabel") { $siteLabelId = [string]$s.SensitivityLabel }

    $siteLabelName = Resolve-LabelName -LabelId $siteLabelId

    # --- Write site row (skip if resume has it)
    $siteKey = "Site|$siteUrl|"
    if (-not $exportedKeys.Contains($siteKey)) {
        $row = @{
            Timestamp               = (Get-Date).ToString("s")
            Scope                   = "Site"
            SiteUrl                 = $siteUrl
            SiteTitle               = $siteTitle
            SiteTemplate            = $siteTemplate
            SiteSensitivityLabelId  = $siteLabelId
            SiteSensitivityLabelName= $siteLabelName
            LibraryTitle            = ""
            LibraryId               = ""
            LibraryServerRelativeUrl= ""
            LibraryDefaultLabelId   = ""
            LibraryDefaultLabelName = ""
            Notes                   = ""
        }
        Append-CsvRow -Path $OutputCsvPath -Row $row -Headers $headers
        [void]$exportedKeys.Add($siteKey)
    }

    # --- Enumerate libraries (PnP)
    try {
        Start-Sleep -Milliseconds $InterSiteDelayMs

        $siteConn = Invoke-WithRetry -Operation {
            Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Tenant $TenantId -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -ReturnConnection
        }

        $lists = Invoke-WithRetry -Operation {
            # DefaultSensitivityLabelForLibrary is the doc-library default sensitivity label setting in PnP list model. 【5-44e944】
            Get-PnPList -Connection $siteConn -Includes "Id","Title","BaseTemplate","Hidden","RootFolder","DefaultSensitivityLabelForLibrary"
        }

        $docLibs = $lists | Where-Object { $_.BaseTemplate -eq 101 }

        if (-not $IncludeHiddenLibraries) {
            $docLibs = $docLibs | Where-Object { -not $_.Hidden }
        }

        foreach ($lib in $docLibs) {
            $libId = [string]$lib.Id
            $libKey = "Library|$siteUrl|$libId"
            if ($exportedKeys.Contains($libKey)) { continue }

            $libTitle = [string]$lib.Title

            $serverRel = ""
            if ($lib.PSObject.Properties.Name -contains "RootFolder" -and $null -ne $lib.RootFolder) {
                if ($lib.RootFolder.PSObject.Properties.Name -contains "ServerRelativeUrl") {
                    $serverRel = [string]$lib.RootFolder.ServerRelativeUrl
                }
            }

            $libLabelId = ""
            if ($lib.PSObject.Properties.Name -contains "DefaultSensitivityLabelForLibrary") {
                $libLabelId = [string]$lib.DefaultSensitivityLabelForLibrary
            }
            $libLabelName = Resolve-LabelName -LabelId $libLabelId

            $row = @{
                Timestamp               = (Get-Date).ToString("s")
                Scope                   = "Library"
                SiteUrl                 = $siteUrl
                SiteTitle               = $siteTitle
                SiteTemplate            = $siteTemplate
                SiteSensitivityLabelId  = $siteLabelId
                SiteSensitivityLabelName= $siteLabelName
                LibraryTitle            = $libTitle
                LibraryId               = $libId
                LibraryServerRelativeUrl= $serverRel
                LibraryDefaultLabelId   = $libLabelId
                LibraryDefaultLabelName = $libLabelName
                Notes                   = "DefaultSensitivityLabelForLibrary"
            }
            Append-CsvRow -Path $OutputCsvPath -Row $row -Headers $headers
            [void]$exportedKeys.Add($libKey)
        }
    }
    catch {
        $errRow = @{
            Timestamp   = (Get-Date).ToString("s")
            Scope       = "LibraryEnumeration"
            SiteUrl     = $siteUrl
            LibraryTitle= ""
            Operation   = "EnumerateLibraries"
            Error       = $_.Exception.Message
        }
        Append-CsvRow -Path $ErrorCsvPath -Row $errRow -Headers $errorHeaders
        continue
    }
}

Write-Host "Done. Output: $OutputCsvPath" -ForegroundColor Green
Write-Host "Errors: $ErrorCsvPath" -ForegroundColor Yellow