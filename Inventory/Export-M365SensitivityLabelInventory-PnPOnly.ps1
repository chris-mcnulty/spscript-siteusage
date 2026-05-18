<#
.SYNOPSIS
PnP.PowerShell-only inventory of:
  - Site container sensitivity label (SensitivityLabel) (best-effort; per-site Identity fallback)
  - Library default sensitivity label (DefaultSensitivityLabelForLibrary) for document libraries
Outputs streaming CSV with resume support.

WHY PnP-ONLY?
Avoids assembly conflicts between Microsoft.Online.SharePoint.PowerShell and PnP.PowerShell.

NOTES ON SITE SENSITIVITY LABEL
Get-PnPTenantSite can return empty SensitivityLabel in bulk; querying per-site with -Identity is often required. 【4-9829aa】

REFERENCES
- Get-PnPTenantSite supports -IncludeOneDriveSites and -Identity 【3-c08909】
- Get-PnPList supports -Includes for extra properties 【2-db1d5c】
- Set-PnPList documents DefaultSensitivityLabelForLibrary (this script reads it from list objects) 【1-4f7220】
- Get-PnPAvailableSensitivityLabel provides label catalog (optional mapping) 【5-7e46c2】
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$TenantName,   # e.g. "synozur" (used to build https://<tenant>-admin.sharepoint.com)

    [Parameter(Mandatory=$true)]
    [string]$ClientId,     # app registration (client) id

    [Parameter(Mandatory=$true)]
    [string]$Tenant,       # tenant id (GUID) OR tenant domain (e.g. synozur.onmicrosoft.com)

    [Parameter(Mandatory=$true)]
    [string]$CertificatePath,  # path to .pfx

    [Parameter(Mandatory=$true)]
    [SecureString]$CertificatePassword,

    [Parameter(Mandatory=$false)]
    [string]$OutputCsvPath = ".\M365_SensitivityLabel_Inventory.csv",

    [Parameter(Mandatory=$false)]
    [string]$ErrorCsvPath = ".\M365_SensitivityLabel_Inventory_Errors.csv",

    [Parameter(Mandatory=$false)]
    [bool]$IncludeOneDriveSites = $true, # default ON 【3-c08909】

    [Parameter(Mandatory=$false)]
    [bool]$IncludeHiddenLibraries = $false,

    [Parameter(Mandatory=$false)]
    [bool]$ResolveLabelNames = $true, # uses Get-PnPAvailableSensitivityLabel 【5-7e46c2】

    [Parameter(Mandatory=$false)]
    [switch]$Resume,

    [Parameter(Mandatory=$false)]
    [int]$InterSiteDelayMs = 100
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Import-Module PnP.PowerShell -ErrorAction Stop

# --------------------------
# Helpers: CSV streaming
# --------------------------
function Ensure-CsvHeader {
    param([string]$Path, [string[]]$Headers)
    if (-not (Test-Path $Path)) { ($Headers -join ",") | Out-File -FilePath $Path -Encoding utf8 }
}

function Append-CsvRow {
    param([string]$Path, [hashtable]$Row, [string[]]$Headers)
    $values = foreach ($h in $Headers) {
        $v = ""
        if ($Row.ContainsKey($h) -and $null -ne $Row[$h]) { $v = [string]$Row[$h] }
        $v = $v.Replace('"','""')
        '"' + $v + '"'
    }
    ($values -join ",") | Add-Content -Path $Path -Encoding utf8
}

function Invoke-WithRetry {
    param([scriptblock]$Operation, [int]$MaxRetries = 6, [int]$InitialDelaySeconds = 2)
    $delay = $InitialDelaySeconds
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try { return & $Operation }
        catch {
            if ($attempt -eq $MaxRetries) { throw }
            Start-Sleep -Seconds $delay
            $delay = [Math]::Min($delay * 2, 60)
        }
    }
}

# --------------------------
# Output headers
# --------------------------
$headers = @(
    "Timestamp",
    "Scope",                       # Site | Library
    "SiteUrl",
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
# Resume support
# Key = Scope|SiteUrl|LibraryId (LibraryId empty for Site rows)
# --------------------------
$exportedKeys = New-Object 'System.Collections.Generic.HashSet[string]'
if ($Resume -and (Test-Path $OutputCsvPath)) {
    try {
        $existing = Import-Csv -Path $OutputCsvPath
        foreach ($r in $existing) {
            $k = "{0}|{1}|{2}" -f $r.Scope, $r.SiteUrl, $r.LibraryId
            [void]$exportedKeys.Add($k)
        }
    } catch {
        # If parsing fails, continue without resume set
    }
}

# --------------------------
# Connect to admin
# --------------------------
$adminUrl = "https://$TenantName-admin.sharepoint.com"

Write-Host "Connecting to PnP admin: $adminUrl" -ForegroundColor Cyan
$adminConn = Invoke-WithRetry {
    Connect-PnPOnline `
        -Url $adminUrl `
        -ClientId $ClientId `
        -Tenant $Tenant `
        -CertificatePath $CertificatePath `
        -CertificatePassword $CertificatePassword `
        -ReturnConnection
}

# --------------------------
# Optional: build label lookup (GUID -> display name)
# --------------------------
$labelLookup = @{}
if ($ResolveLabelNames) {
    try {
        $labels = Get-PnPAvailableSensitivityLabel -Connection $adminConn  # 【5-7e46c2】
        foreach ($l in $labels) {
            $id = $l.Id.ToString().ToLower()
            $name = $null
            if ($l.PSObject.Properties.Name -contains "DisplayName") { $name = $l.DisplayName }
            if (-not $name -and ($l.PSObject.Properties.Name -contains "Name")) { $name = $l.Name }
            if ($name) { $labelLookup[$id] = $name }
        }
    } catch {
        Write-Host "Label name mapping unavailable; exporting label IDs only." -ForegroundColor Yellow
    }
}

function Resolve-LabelName {
    param([string]$LabelId)
    if ([string]::IsNullOrWhiteSpace($LabelId)) { return "" }
    $k = $LabelId.ToLower()
    if ($labelLookup.ContainsKey($k)) { return $labelLookup[$k] }
    return ""
}

# --------------------------
# Get all tenant sites (include OneDrive sites by default) 【3-c08909】
# --------------------------
Write-Host "Retrieving tenant sites..." -ForegroundColor Cyan
$sites = Invoke-WithRetry {
    Get-PnPTenantSite -Connection $adminConn -IncludeOneDriveSites:$IncludeOneDriveSites
}

Write-Host ("Sites found: {0}" -f $sites.Count) -ForegroundColor Cyan

foreach ($s in $sites) {
    $siteUrl = $s.Url

    # ---- Get site sensitivity label (best-effort)
    # Known issue: SensitivityLabel can be empty in bulk; per-site Identity often works. 【4-9829aa】
    $siteLabelId = ""
    try {
        $detail = Invoke-WithRetry { Get-PnPTenantSite -Connection $adminConn -Identity $siteUrl }
        if ($detail.PSObject.Properties.Name -contains "SensitivityLabel") {
            $siteLabelId = [string]$detail.SensitivityLabel
        }
    } catch {
        # leave blank, record in error file
        Append-CsvRow -Path $ErrorCsvPath -Headers $errorHeaders -Row @{
            Timestamp    = (Get-Date).ToString("s")
            Scope        = "Site"
            SiteUrl      = $siteUrl
            LibraryTitle = ""
            Operation    = "Get-PnPTenantSite -Identity (SensitivityLabel)"
            Error        = $_.Exception.Message
        }
    }

    $siteLabelName = Resolve-LabelName -LabelId $siteLabelId

    # ---- write site row (resume-aware)
    $siteKey = "Site|$siteUrl|"
    if (-not $exportedKeys.Contains($siteKey)) {
        Append-CsvRow -Path $OutputCsvPath -Headers $headers -Row @{
            Timestamp               = (Get-Date).ToString("s")
            Scope                   = "Site"
            SiteUrl                 = $siteUrl
            SiteSensitivityLabelId  = $siteLabelId
            SiteSensitivityLabelName= $siteLabelName
            LibraryTitle            = ""
            LibraryId               = ""
            LibraryServerRelativeUrl= ""
            LibraryDefaultLabelId   = ""
            LibraryDefaultLabelName = ""
            Notes                   = ""
        }
        [void]$exportedKeys.Add($siteKey)
    }

    # ---- enumerate libraries
    try {
        Start-Sleep -Milliseconds $InterSiteDelayMs

        $siteConn = Invoke-WithRetry {
            Connect-PnPOnline `
                -Url $siteUrl `
                -ClientId $ClientId `
                -Tenant $Tenant `
                -CertificatePath $CertificatePath `
                -CertificatePassword $CertificatePassword `
                -ReturnConnection
        }

        # Get-PnPList supports -Includes to pull extra properties 【2-db1d5c】
        $lists = Invoke-WithRetry {
            Get-PnPList -Connection $siteConn -Includes "Id","Title","BaseTemplate","Hidden","RootFolder","DefaultSensitivityLabelForLibrary"
        }

        # Document libraries have BaseTemplate 101 (standard SharePoint convention; used widely in PnP examples)
        $docLibs = $lists | Where-Object { $_.BaseTemplate -eq 101 }
        if (-not $IncludeHiddenLibraries) {
            $docLibs = $docLibs | Where-Object { -not $_.Hidden }
        }

        foreach ($lib in $docLibs) {
            $libId = [string]$lib.Id
            $libKey = "Library|$siteUrl|$libId"
            if ($exportedKeys.Contains($libKey)) { continue }

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

            Append-CsvRow -Path $OutputCsvPath -Headers $headers -Row @{
                Timestamp               = (Get-Date).ToString("s")
                Scope                   = "Library"
                SiteUrl                 = $siteUrl
                SiteSensitivityLabelId  = $siteLabelId
                SiteSensitivityLabelName= $siteLabelName
                LibraryTitle            = [string]$lib.Title
                LibraryId               = $libId
                LibraryServerRelativeUrl= $serverRel
                LibraryDefaultLabelId   = $libLabelId
                LibraryDefaultLabelName = $libLabelName
                Notes                   = "DefaultSensitivityLabelForLibrary"
            }
            [void]$exportedKeys.Add($libKey)
        }
    } catch {
        Append-CsvRow -Path $ErrorCsvPath -Headers $errorHeaders -Row @{
            Timestamp    = (Get-Date).ToString("s")
            Scope        = "LibraryEnumeration"
            SiteUrl      = $siteUrl
            LibraryTitle = ""
            Operation    = "Enumerate libraries"
            Error        = $_.Exception.Message
        }
        continue
    }
}

Write-Host "Done." -ForegroundColor Green
Write-Host "Output: $OutputCsvPath" -ForegroundColor Green
Write-Host "Errors: $ErrorCsvPath" -ForegroundColor Yellow