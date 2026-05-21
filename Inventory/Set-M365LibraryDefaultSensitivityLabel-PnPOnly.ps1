[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
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

    # The sensitivity label GUID to apply as the library default.
    [Parameter(Mandatory=$true)]
    [string]$DefaultLabelId,

    # CSV audit log of every library evaluated and what was done.
    [string]$OutputCsvPath = ".\M365_LibraryDefaultLabel_Apply.csv",
    [string]$ErrorCsvPath  = ".\M365_LibraryDefaultLabel_Apply_Errors.csv",

    [bool]$IncludeOneDriveSites = $false,
    [bool]$IncludeHiddenLibraries = $false,

    # If $true, replace an existing different label. If $false, only set when the library has no default label.
    [switch]$OverwriteExisting,

    # Skip libraries already recorded in $OutputCsvPath as Action=Set or Action=AlreadySet.
    [switch]$Resume,

    # Optional filter: only process site URLs matching this wildcard (e.g. "*/sites/Finance*").
    [string]$SiteUrlLike
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Import-Module PnP.PowerShell -ErrorAction Stop

# --------------------------
# Validate label GUID
# --------------------------
$labelGuid = [Guid]::Empty
if (-not [Guid]::TryParse($DefaultLabelId, [ref]$labelGuid)) {
    throw "DefaultLabelId '$DefaultLabelId' is not a valid GUID."
}
$DefaultLabelId = $labelGuid.ToString()

# --------------------------
# CSV Setup
# --------------------------
$headers = @(
    "Timestamp","SiteUrl","LibraryTitle","LibraryId","LibraryServerRelativeUrl",
    "PreviousLabelId","PreviousLabelName","TargetLabelId","TargetLabelName",
    "Action","Detail"
)
$errorHeaders = @("Timestamp","Scope","SiteUrl","LibraryTitle","Operation","Error")

if (!(Test-Path $OutputCsvPath)) {
    ($headers -join ",") | Out-File $OutputCsvPath
}
if (!(Test-Path $ErrorCsvPath)) {
    ($errorHeaders -join ",") | Out-File $ErrorCsvPath
}

function CsvEscape([object]$v) {
    if ($null -eq $v) { return '""' }
    $s = [string]$v
    $s = $s.Replace('"','""')
    return '"' + $s + '"'
}

function Write-CsvRow($path, $values) {
    $line = ($values | ForEach-Object { CsvEscape $_ }) -join ","
    Add-Content -Path $path -Value $line
}

# Resume logic — skip libraries we've already reached a terminal state for.
# Only Set / AlreadySet are terminal. WouldSet (from -WhatIf) and
# SkippedExistingLabel (depends on -OverwriteExisting) must NOT be cached,
# so a follow-up apply or -OverwriteExisting run re-evaluates them.
$processedKeys = @{}
if ($Resume -and (Test-Path $OutputCsvPath)) {
    Import-Csv $OutputCsvPath | ForEach-Object {
        if ($_.Action -in @("Set","AlreadySet")) {
            $k = "$($_.SiteUrl)|$($_.LibraryId)"
            $processedKeys[$k] = $true
        }
    }
}

# --------------------------
# Connect to admin
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
# Resolve label names (best-effort)
# --------------------------
$labelMap = @{}
try {
    $labels = Get-PnPAvailableSensitivityLabel -Connection $adminConn
    foreach ($l in $labels) {
        $labelMap[$l.Id.ToString().ToLower()] = $l.DisplayName
    }
} catch {
    Write-Host "Label name mapping unavailable; will log GUIDs only." -ForegroundColor Yellow
}

function Resolve-LabelName([string]$id) {
    if ([string]::IsNullOrWhiteSpace($id)) { return "" }
    $k = $id.ToLower()
    if ($labelMap.ContainsKey($k)) { return $labelMap[$k] }
    return ""
}

$targetLabelName = Resolve-LabelName $DefaultLabelId
if ($targetLabelName) {
    Write-Host ("Target default label: {0}  ({1})" -f $targetLabelName, $DefaultLabelId) -ForegroundColor Green
} else {
    Write-Host ("Target default label ID: {0}" -f $DefaultLabelId) -ForegroundColor Green
}

# --------------------------
# Enumerate sites
# --------------------------
Write-Host "Retrieving sites..." -ForegroundColor Cyan
$sites = Get-PnPTenantSite -Connection $adminConn -IncludeOneDriveSites:$IncludeOneDriveSites

if ($SiteUrlLike) {
    $sites = $sites | Where-Object { $_.Url -like $SiteUrlLike }
}

$totalSites = $sites.Count
$siteIndex  = 0
$libsSeen   = 0
$libsSet    = 0
$libsSkipped= 0

Write-Host ("Sites to process: {0}" -f $totalSites) -ForegroundColor Cyan

# --------------------------
# Process sites
# --------------------------
foreach ($site in $sites) {

    $siteIndex++
    Write-Host ""
    Write-Host ("[{0}/{1}] {2}" -f $siteIndex, $totalSites, $site.Url) -ForegroundColor Cyan

    try {
        $siteConn = Connect-PnPOnline `
            -Url $site.Url `
            -ClientId $ClientId `
            -Tenant $Tenant `
            -CertificatePath $CertificatePath `
            -CertificatePassword $CertificatePassword `
            -ReturnConnection
    }
    catch {
        Write-Host ("  Cannot connect to site: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        Write-CsvRow $ErrorCsvPath @((Get-Date -Format s), "Site", $site.Url, "", "ConnectSite", $_.Exception.Message)
        continue
    }

    try {
        $lists = Get-PnPList -Connection $siteConn -Includes "Id","Title","Hidden","BaseTemplate","RootFolder","DefaultSensitivityLabelForLibrary"
    }
    catch {
        Write-Host ("  Skipping site (no access / archived / restricted)") -ForegroundColor Yellow
        Write-CsvRow $ErrorCsvPath @((Get-Date -Format s), "Site", $site.Url, "", "EnumerateLibraries", $_.Exception.Message)
        continue
    }

    $docLibs = $lists | Where-Object { $_.BaseTemplate -eq 101 }
    if (-not $IncludeHiddenLibraries) {
        $docLibs = $docLibs | Where-Object { -not $_.Hidden }
    }

    foreach ($lib in $docLibs) {

        $libsSeen++
        $libId    = $lib.Id.ToString()
        $libTitle = [string]$lib.Title
        $libUrl   = ""
        if ($lib.RootFolder -and $lib.RootFolder.ServerRelativeUrl) {
            $libUrl = [string]$lib.RootFolder.ServerRelativeUrl
        }

        $key = "$($site.Url)|$libId"
        if ($processedKeys.ContainsKey($key)) {
            Write-Host ("    [skip-resume] {0}" -f $libTitle) -ForegroundColor DarkGray
            continue
        }

        $previousLabelId = ""
        if ($lib.PSObject.Properties.Name -contains "DefaultSensitivityLabelForLibrary") {
            $previousLabelId = [string]$lib.DefaultSensitivityLabelForLibrary
        }
        $previousLabelName = Resolve-LabelName $previousLabelId

        $action = ""
        $detail = ""

        $alreadyMatches = ($previousLabelId -and ($previousLabelId.ToLower() -eq $DefaultLabelId.ToLower()))
        $hasDifferent   = ($previousLabelId -and -not $alreadyMatches)

        if ($alreadyMatches) {
            $action = "AlreadySet"
            $libsSkipped++
            Write-Host ("    [=] {0}  (already has target label)" -f $libTitle) -ForegroundColor DarkGreen
        }
        elseif ($hasDifferent -and -not $OverwriteExisting) {
            $action = "SkippedExistingLabel"
            $detail = "Library has a different label; pass -OverwriteExisting to replace."
            $libsSkipped++
            Write-Host ("    [!] {0}  (has different label: {1})" -f $libTitle, ($previousLabelName ? $previousLabelName : $previousLabelId)) -ForegroundColor Yellow
        }
        else {
            # Either empty, or different + OverwriteExisting
            if ($PSCmdlet.ShouldProcess("$($site.Url) :: $libTitle", "Set DefaultSensitivityLabelForLibrary -> $DefaultLabelId")) {
                try {
                    Set-PnPList -Connection $siteConn -Identity $lib.Id -DefaultSensitivityLabelForLibrary $DefaultLabelId | Out-Null
                    $action = "Set"
                    $libsSet++
                    Write-Host ("    [+] {0}  -> {1}" -f $libTitle, ($targetLabelName ? $targetLabelName : $DefaultLabelId)) -ForegroundColor Green
                }
                catch {
                    $action = "Error"
                    $detail = $_.Exception.Message
                    Write-Host ("    [x] {0}  ERROR: {1}" -f $libTitle, $_.Exception.Message) -ForegroundColor Red
                    Write-CsvRow $ErrorCsvPath @((Get-Date -Format s), "Library", $site.Url, $libTitle, "SetDefaultSensitivityLabelForLibrary", $_.Exception.Message)
                }
            }
            else {
                $action = "WouldSet"
                $detail = "WhatIf"
                Write-Host ("    [?] {0}  would set default label" -f $libTitle) -ForegroundColor DarkCyan
            }
        }

        Write-CsvRow $OutputCsvPath @(
            (Get-Date -Format s),
            $site.Url,
            $libTitle,
            $libId,
            $libUrl,
            $previousLabelId,
            $previousLabelName,
            $DefaultLabelId,
            $targetLabelName,
            $action,
            $detail
        )
    }

    if (($siteIndex % 10) -eq 0) {
        Write-Host ("--- Progress: {0}/{1} sites, libs seen={2}, set={3}, skipped={4} ---" -f $siteIndex, $totalSites, $libsSeen, $libsSet, $libsSkipped) -ForegroundColor Green
    }
}

# --------------------------
# Final Summary
# --------------------------
Write-Host ""
Write-Host "====================================" -ForegroundColor Green
Write-Host "Apply complete"
Write-Host ("Sites processed     : {0}" -f $siteIndex)
Write-Host ("Libraries evaluated : {0}" -f $libsSeen)
Write-Host ("Libraries set       : {0}" -f $libsSet)
Write-Host ("Libraries skipped   : {0}" -f $libsSkipped)
Write-Host ("Audit log           : {0}" -f $OutputCsvPath)
Write-Host ("Error log           : {0}" -f $ErrorCsvPath)
Write-Host "===================================="
