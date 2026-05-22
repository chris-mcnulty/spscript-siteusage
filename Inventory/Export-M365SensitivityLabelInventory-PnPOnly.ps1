[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$TenantName,

    [Parameter(Mandatory=$true)]
    [string]$ClientId,

    [Parameter(Mandatory=$true)]
    [string]$Tenant,

    [Parameter(Mandatory=$true)]
    [string]$CertificatePath,

    [Parameter(Mandatory=$true)]
    [SecureString]$CertificatePassword,

    [string]$OutputCsvPath = ".\M365_SensitivityLabel_Inventory.csv",
    [string]$ErrorCsvPath = ".\M365_SensitivityLabel_Inventory_Errors.csv",

    [bool]$IncludeOneDriveSites = $false,  # ✅ EXCLUDED BY DEFAULT

    [switch]$Resume,

    # UPN of a licensed user whose published-label set is used to resolve GUIDs -> display names
    # in the inventory CSV. Required when running app-only because Get-PnPAvailableSensitivityLabel
    # has no user context otherwise. Needs Graph InformationProtectionPolicy.Read.All +
    # User.Read.All (Application) on this app.
    [string]$LabelOwnerUpn
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Import-Module PnP.PowerShell -ErrorAction Stop

# --------------------------
# CSV Setup
# --------------------------
$headers = @(
    "Timestamp","Scope","SiteUrl",
    "SiteSensitivityLabelId","SiteSensitivityLabelName",
    "LibraryTitle","LibraryId","LibraryServerRelativeUrl",
    "LibraryDefaultLabelId","LibraryDefaultLabelName"
)

if (!(Test-Path $OutputCsvPath)) {
    ($headers -join ",") | Out-File $OutputCsvPath
}

# Resume logic
$exportedKeys = @{}
if ($Resume -and (Test-Path $OutputCsvPath)) {
    Import-Csv $OutputCsvPath | ForEach-Object {
        $key = "$($_.Scope)|$($_.SiteUrl)|$($_.LibraryId)"
        $exportedKeys[$key] = $true
    }
}

# --------------------------
# Connect
# --------------------------
$adminUrl = "https://$TenantName-admin.sharepoint.com"
Write-Host "Connecting to $adminUrl..." -ForegroundColor Cyan

$adminConn = Connect-PnPOnline `
    -Url $adminUrl `
    -ClientId $ClientId `
    -Tenant $Tenant `
    -CertificatePath $CertificatePath `
    -CertificatePassword $CertificatePassword `
    -ReturnConnection

# --------------------------
# Try to resolve labels
# --------------------------
function Get-LabelDisplayProp($l) {
    if ($null -eq $l) { return "" }
    foreach ($prop in @("DisplayName","Name","displayName","name")) {
        if ($l.PSObject.Properties.Name -contains $prop) {
            $v = [string]$l.$prop
            if ($v) { return $v }
        }
    }
    return ""
}

$labelMap = @{}
try {
    if ($LabelOwnerUpn) {
        $labels = Get-PnPAvailableSensitivityLabel -Connection $adminConn -User $LabelOwnerUpn
    } else {
        $labels = Get-PnPAvailableSensitivityLabel -Connection $adminConn
    }
    foreach ($l in $labels) {
        if ($l.PSObject.Properties.Name -contains "Id") {
            $labelMap[([string]$l.Id).ToLower()] = (Get-LabelDisplayProp $l)
        }
    }
    Write-Host ("Loaded {0} label name(s) for resolution." -f $labelMap.Count) -ForegroundColor DarkGray
} catch {
    Write-Host ("Label name mapping unavailable; exporting label IDs only. ({0})" -f $_.Exception.Message) -ForegroundColor Yellow
    if (-not $LabelOwnerUpn) {
        Write-Host "Tip: pass -LabelOwnerUpn <user@domain> so app-only auth has a label-policy scope to read." -ForegroundColor Yellow
    }
}

function Resolve-LabelName($id) {
    if ([string]::IsNullOrWhiteSpace($id)) { return "" }
    $k = ([string]$id).ToLower()
    if ($labelMap.ContainsKey($k)) { return $labelMap[$k] }
    return ""
}

# --------------------------
# Get Sites
# --------------------------
Write-Host "Retrieving sites..." -ForegroundColor Cyan

$sites = Get-PnPTenantSite -Connection $adminConn -IncludeOneDriveSites:$IncludeOneDriveSites

$totalSites = $sites.Count
$siteIndex = 0
$libraryCount = 0

Write-Host "Sites found: $totalSites" -ForegroundColor Cyan

# --------------------------
# Process Sites
# --------------------------
foreach ($site in $sites) {

    $siteIndex++
    Write-Host ""
    Write-Host ("[{0}/{1}] Processing site: {2}" -f $siteIndex, $totalSites, $site.Url) -ForegroundColor Cyan

    # Get label
    $siteLabelId = ""
    try {
        $detail = Get-PnPTenantSite -Connection $adminConn -Identity $site.Url
        $siteLabelId = $detail.SensitivityLabel
    } catch {}

    $siteLabelName = Resolve-LabelName $siteLabelId

    if ($siteLabelName) {
        Write-Host ("  Site Label: {0}" -f $siteLabelName) -ForegroundColor DarkCyan
    } elseif ($siteLabelId) {
        Write-Host ("  Site Label ID: {0}" -f $siteLabelId) -ForegroundColor DarkYellow
    }

    # Write site row
    $siteKey = "Site|$($site.Url)|"
    if (!$exportedKeys.ContainsKey($siteKey)) {

        Add-Content $OutputCsvPath (
            '"' + (Get-Date -Format s) + '","Site","' + $site.Url + '","' +
            $siteLabelId + '","' + $siteLabelName + '","","","","",""'
        )

        $exportedKeys[$siteKey] = $true
    }

    # Connect to site
    $siteConn = Connect-PnPOnline `
        -Url $site.Url `
        -ClientId $ClientId `
        -Tenant $Tenant `
        -CertificatePath $CertificatePath `
        -CertificatePassword $CertificatePassword `
        -ReturnConnection

    # Get libraries
    #$lists = Get-PnPList -Connection $siteConn -Includes "Id","Title","BaseTemplate","RootFolder","DefaultSensitivityLabelForLibrary"
    try {
            $lists = Get-PnPList -Connection $siteConn -Includes "Id","Title","BaseTemplate","RootFolder","DefaultSensitivityLabelForLibrary"
        }
        catch {
            Write-Host ("  ⚠️ Skipping site (no access / archived / restricted): {0}" -f $site.Url) -ForegroundColor Yellow

            Add-Content $ErrorCsvPath ( '"' + (Get-Date -Format s) + '","Site","' + $site.Url + '","","","","","","","ACCESS_DENIED_OR_ARCHIVED"' )

            continue
        }
    $libs = $lists | Where-Object { $_.BaseTemplate -eq 101 }

    foreach ($lib in $libs) {

        $libraryCount++
        Write-Host ("    [{0}] Library: {1}" -f $libraryCount, $lib.Title) -ForegroundColor Gray

        $libId = $lib.Id.ToString()
        $libKey = "Library|$($site.Url)|$libId"

        if ($exportedKeys.ContainsKey($libKey)) { continue }

        $libLabelId = $lib.DefaultSensitivityLabelForLibrary
        $libLabelName = Resolve-LabelName $libLabelId

        if ($libLabelName) {
            Write-Host ("      Default Label: {0}" -f $libLabelName) -ForegroundColor DarkGreen
        } elseif ($libLabelId) {
            Write-Host ("      Default Label ID: {0}" -f $libLabelId) -ForegroundColor DarkYellow
        }

        $url = $lib.RootFolder.ServerRelativeUrl

        Add-Content $OutputCsvPath (
            '"' + (Get-Date -Format s) + '","Library","' + $site.Url + '","' +
            $siteLabelId + '","' + $siteLabelName + '","' +
            $lib.Title + '","' + $libId + '","' + $url + '","' +
            $libLabelId + '","' + $libLabelName + '"'
        )

        $exportedKeys[$libKey] = $true
    }

    if (($siteIndex % 10) -eq 0) {
        Write-Host ("--- Progress: {0}/{1} sites, {2} libraries ---" -f $siteIndex, $totalSites, $libraryCount) -ForegroundColor Green
    }
}

# --------------------------
# Final Summary
# --------------------------
Write-Host ""
Write-Host "====================================" -ForegroundColor Green
Write-Host "Inventory complete"
Write-Host ("Sites processed: {0}" -f $siteIndex)
Write-Host ("Libraries scanned: {0}" -f $libraryCount)
Write-Host ("Output: {0}" -f $OutputCsvPath)
Write-Host "===================================="
``