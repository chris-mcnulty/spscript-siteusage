<#
.SYNOPSIS
Applies a single Microsoft Purview sensitivity label as the default to every
UNLABELED file across the tenant. Files that already carry any label are
left alone.

.DESCRIPTION
Iterates every site, every document library, and every file. For each file
that has no current sensitivity label, the script assigns the supplied
default label via the Microsoft Graph endpoint

  POST /drives/{drive-id}/items/{item-id}/assignSensitivityLabel

Files that already have *any* sensitivity label (matching or different) are
skipped — this script intentionally never overwrites. Use
Apply-M365FileSensitivityLabel-PnPOnly.ps1 with -OverwriteExisting when you
need to replace existing labels.

The same SPO-Inventory Entra app + certificate used for the inventory and
the library-default scripts authenticates here. The required Graph
Application permissions are:

  - Files.ReadWrite.All            (enumerate driveItems, assign labels)
  - Sites.Read.All                 (resolve site -> drive)
  - InformationProtectionPolicy.Read.All   (resolve label names; optional)

.NOTES
Throttling: 429 / 503 responses honor Retry-After; otherwise exponential
            backoff (2,4,8,... capped at 60s) up to 5 attempts.
Resume:     re-running with -Resume skips files already in a terminal state
            (Labeled, AlreadyLabeled, SkippedExistingLabel) in $OutputCsvPath.
            WouldLabel rows from -WhatIf are intentionally re-evaluated.
#>
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium', DefaultParameterSetName='ById')]
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

    # The default sensitivity label GUID to apply to unlabeled files.
    # Mutually exclusive with -DefaultLabelName.
    [Parameter(Mandatory=$true, ParameterSetName='ById')]
    [string]$DefaultLabelId,

    # The default sensitivity label display name to apply. Requires
    # -LabelOwnerUpn so the script can resolve name -> GUID via
    # Get-PnPAvailableSensitivityLabel. Mutually exclusive with -DefaultLabelId.
    [Parameter(Mandatory=$true, ParameterSetName='ByName')]
    [string]$DefaultLabelName,

    # standard | privileged | auto
    [ValidateSet("standard","privileged","auto")]
    [string]$AssignmentMethod = "standard",

    [string]$OutputCsvPath = ".\M365_FileDefaultLabel_Apply.csv",
    [string]$ErrorCsvPath  = ".\M365_FileDefaultLabel_Apply_Errors.csv",

    [bool]$IncludeOneDriveSites   = $false,
    [bool]$IncludeHiddenLibraries = $false,

    [switch]$Resume,

    [string]$SiteUrlLike,

    # UPN of a licensed user whose published-label set is used to resolve GUIDs -> display names
    # in the audit CSV / console output. Required when running app-only because
    # Get-PnPAvailableSensitivityLabel has no user context otherwise. Needs Graph
    # InformationProtectionPolicy.Read.All + User.Read.All (Application) on this app.
    [string]$LabelOwnerUpn,

    # Cap per library to throttle very large rollouts (0 = no cap).
    [int]$MaxFilesPerLibrary = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Import-Module PnP.PowerShell -ErrorAction Stop

# --------------------------
# Validate label GUID (only when supplied directly — name mode resolves to GUID later)
# --------------------------
if ($PSCmdlet.ParameterSetName -eq 'ById') {
    $labelGuid = [Guid]::Empty
    if (-not [Guid]::TryParse($DefaultLabelId, [ref]$labelGuid)) {
        throw "DefaultLabelId '$DefaultLabelId' is not a valid GUID."
    }
    $DefaultLabelId = $labelGuid.ToString()
}
elseif ($PSCmdlet.ParameterSetName -eq 'ByName' -and -not $LabelOwnerUpn) {
    throw "-DefaultLabelName requires -LabelOwnerUpn so the label catalog can be loaded and the name resolved."
}

# --------------------------
# CSV setup
# --------------------------
$headers = @(
    "Timestamp","SiteUrl","LibraryTitle","DriveId","ItemId","ItemPath",
    "PreviousLabelId","PreviousLabelName","TargetLabelId","TargetLabelName","Action","Detail"
)
$errorHeaders = @("Timestamp","Scope","SiteUrl","LibraryTitle","Operation","Error")

if (!(Test-Path $OutputCsvPath)) { ($headers      -join ",") | Out-File $OutputCsvPath -WhatIf:$false -Confirm:$false }
if (!(Test-Path $ErrorCsvPath))  { ($errorHeaders -join ",") | Out-File $ErrorCsvPath  -WhatIf:$false -Confirm:$false }

function CsvEscape([object]$v) {
    if ($null -eq $v) { return '""' }
    $s = [string]$v
    return '"' + $s.Replace('"','""') + '"'
}
function Write-CsvRow($path, $values) {
    $line = ($values | ForEach-Object { CsvEscape $_ }) -join ","
    # Audit writes must bypass -WhatIf / -Confirm so the CSV records WouldLabel rows during preview runs.
    Add-Content -Path $path -Value $line -WhatIf:$false -Confirm:$false
}

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

# Resume cache: keyed by DriveId|ItemId. Both Labeled/AlreadyLabeled AND
# SkippedExistingLabel are terminal here, because this script intentionally
# never overwrites — a file with any existing label is decided and stays
# decided across runs. Only WouldLabel (from -WhatIf) is re-evaluated.
$processedItems = @{}
if ($Resume -and (Test-Path $OutputCsvPath)) {
    Import-Csv $OutputCsvPath | ForEach-Object {
        if ($_.Action -in @("Labeled","AlreadyLabeled","SkippedExistingLabel")) {
            $processedItems["$($_.DriveId)|$($_.ItemId)"] = $true
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
# Label name map (best-effort)
# --------------------------
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
    Write-Host ("Label name mapping unavailable; logging GUIDs only. ({0})" -f $_.Exception.Message) -ForegroundColor Yellow
    if (-not $LabelOwnerUpn) {
        Write-Host "Tip: pass -LabelOwnerUpn <user@domain> so app-only auth has a label-policy scope to read." -ForegroundColor Yellow
    }
}

# --------------------------
# Resolve -DefaultLabelName -> GUID, or pre-flight-validate -DefaultLabelId against the catalog
# --------------------------
function Format-LabelCatalog($map) {
    if ($map.Count -eq 0) { return "  (catalog empty)" }
    ($map.GetEnumerator() | Sort-Object Value | ForEach-Object { "  - $($_.Value)  ($($_.Key))" }) -join "`n"
}

if ($PSCmdlet.ParameterSetName -eq 'ByName') {
    if ($labelMap.Count -eq 0) {
        throw "Could not load the label catalog, so -DefaultLabelName '$DefaultLabelName' cannot be resolved. Check Graph permissions (InformationProtectionPolicy.Read.All, User.Read.All) and that -LabelOwnerUpn '$LabelOwnerUpn' is a licensed user in this tenant."
    }
    $match = $labelMap.GetEnumerator() | Where-Object { $_.Value -ieq $DefaultLabelName } | Select-Object -First 1
    if (-not $match) {
        throw "Label name '$DefaultLabelName' not found in tenant catalog.`nAvailable labels:`n$(Format-LabelCatalog $labelMap)"
    }
    $DefaultLabelId = [string]$match.Key
    Write-Host ("Resolved -DefaultLabelName '{0}' -> {1}" -f $DefaultLabelName, $DefaultLabelId) -ForegroundColor Green
}
elseif ($labelMap.Count -gt 0) {
    if (-not $labelMap.ContainsKey($DefaultLabelId.ToLower())) {
        throw "DefaultLabelId '$DefaultLabelId' is not in the tenant label catalog. Aborting before any writes.`nAvailable labels:`n$(Format-LabelCatalog $labelMap)"
    }
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
# Graph helpers (use admin connection; PnP routes Graph calls via the same app identity)
# --------------------------
function Invoke-Graph {
    param(
        [string]$Url,
        [string]$Method = "GET",
        [object]$Body
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            if ($Body -ne $null -and $Method -ne "GET") {
                $json = if ($Body -is [string]) { $Body } else { ($Body | ConvertTo-Json -Depth 10 -Compress) }
                return Invoke-PnPGraphMethod -Connection $adminConn -Url $Url -Method $Method -Content $json -ContentType "application/json"
            }
            else {
                return Invoke-PnPGraphMethod -Connection $adminConn -Url $Url -Method $Method
            }
        }
        catch {
            $msg = $_.Exception.Message
            $isThrottle = $msg -match "429|503|throttle|Too Many Requests|Service Unavailable"
            if ($isThrottle -and $attempt -lt 5) {
                $wait = $null
                try {
                    $resp = $_.Exception.PSObject.Properties['Response']
                    if ($resp -and $resp.Value -and $resp.Value.Headers) {
                        $ra = $resp.Value.Headers['Retry-After']
                        if (-not $ra) { $ra = $resp.Value.Headers.RetryAfter }
                        if ($ra) {
                            $raStr = [string]$ra
                            $sec = 0
                            if ([int]::TryParse($raStr, [ref]$sec) -and $sec -gt 0) {
                                $wait = [Math]::Min(300, $sec)
                            }
                        }
                    }
                } catch { }
                if (-not $wait) {
                    if ($msg -match "Retry-After[^\d]*([0-9]+)") {
                        $wait = [Math]::Min(300, [int]$Matches[1])
                    }
                }
                if (-not $wait) {
                    $wait = [int][Math]::Min(60, [Math]::Pow(2, $attempt))
                }
                Write-Host ("    [throttle] sleeping {0}s (attempt {1}/5)..." -f $wait, $attempt) -ForegroundColor DarkYellow
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

function Get-GraphSiteId([string]$siteUrl) {
    $u = [Uri]$siteUrl
    $path = $u.AbsolutePath.TrimEnd('/')
    if ([string]::IsNullOrEmpty($path)) {
        $endpoint = "v1.0/sites/$($u.Host)"
    } else {
        $endpoint = "v1.0/sites/$($u.Host):$path"
    }
    $res = Invoke-Graph -Url $endpoint
    return $res.id
}

function Get-LibraryDrive($graphSiteId, $listId) {
    $endpoint = "v1.0/sites/$graphSiteId/lists/$listId/drive"
    return Invoke-Graph -Url $endpoint
}

function Get-DriveFilesRecursive($driveId) {
    $queue = New-Object System.Collections.Generic.Queue[string]
    $queue.Enqueue("root")
    while ($queue.Count -gt 0) {
        $node = $queue.Dequeue()
        $url = "v1.0/drives/$driveId/items/$node/children?`$top=200&`$select=id,name,folder,file,parentReference,size,sensitivityLabel"
        while ($url) {
            $page = Invoke-Graph -Url $url
            foreach ($item in $page.value) {
                if ($item.PSObject.Properties.Name -contains "folder" -and $null -ne $item.folder) {
                    $queue.Enqueue($item.id)
                }
                elseif ($item.PSObject.Properties.Name -contains "file" -and $null -ne $item.file) {
                    Write-Output $item
                }
            }
            $url = $null
            if ($page.PSObject.Properties.Name -contains "@odata.nextLink" -and $page.'@odata.nextLink') {
                $next = [string]$page.'@odata.nextLink'
                $url = $next -replace "^https?://graph\.microsoft\.com/", ""
            }
        }
    }
}

function Get-ItemLabelId($item) {
    if ($item.PSObject.Properties.Name -contains "sensitivityLabel" -and $null -ne $item.sensitivityLabel) {
        if ($item.sensitivityLabel.PSObject.Properties.Name -contains "id") {
            return [string]$item.sensitivityLabel.id
        }
    }
    return ""
}

function Apply-Label($driveId, $itemId, $labelId, $method) {
    $body = @{
        sensitivityLabelId = $labelId
        assignmentMethod   = $method
    }
    $url = "v1.0/drives/$driveId/items/$itemId/assignSensitivityLabel"
    Invoke-Graph -Url $url -Method POST -Body $body | Out-Null
}

# --------------------------
# Enumerate sites
# --------------------------
Write-Host "Retrieving sites..." -ForegroundColor Cyan
$sites = Get-PnPTenantSite -Connection $adminConn -IncludeOneDriveSites:$IncludeOneDriveSites
if ($SiteUrlLike) { $sites = $sites | Where-Object { $_.Url -like $SiteUrlLike } }

# Force an array so .Count is valid even when the filter narrows to a single site (or none).
$sites = @($sites)
$totalSites   = $sites.Count
$siteIndex    = 0
$filesSeen    = 0
$filesLabeled = 0
$filesSkippedExisting = 0

Write-Host ("Sites to process: {0}" -f $totalSites) -ForegroundColor Cyan

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
    } catch {
        Write-CsvRow $ErrorCsvPath @((Get-Date -Format s),"Site",$site.Url,"","ConnectSite",$_.Exception.Message)
        Write-Host ("  Cannot connect: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        continue
    }

    try {
        $lists = Get-PnPList -Connection $siteConn -Includes "Id","Title","Hidden","BaseTemplate","RootFolder"
    } catch {
        Write-CsvRow $ErrorCsvPath @((Get-Date -Format s),"Site",$site.Url,"","EnumerateLibraries",$_.Exception.Message)
        Write-Host "  Skipping site (no access / archived / restricted)" -ForegroundColor Yellow
        continue
    }

    $docLibs = $lists | Where-Object { $_.BaseTemplate -eq 101 }
    if (-not $IncludeHiddenLibraries) { $docLibs = $docLibs | Where-Object { -not $_.Hidden } }

    $graphSiteId = $null
    try { $graphSiteId = Get-GraphSiteId $site.Url } catch {
        Write-CsvRow $ErrorCsvPath @((Get-Date -Format s),"Site",$site.Url,"","ResolveGraphSite",$_.Exception.Message)
        Write-Host ("  Cannot resolve Graph site: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        continue
    }

    foreach ($lib in $docLibs) {

        $libId    = $lib.Id.ToString()
        $libTitle = [string]$lib.Title

        Write-Host ("  Library: {0}  ->  {1}" -f $libTitle, ($targetLabelName ? $targetLabelName : $DefaultLabelId)) -ForegroundColor Cyan

        $drive = $null
        try { $drive = Get-LibraryDrive $graphSiteId $libId } catch {
            Write-CsvRow $ErrorCsvPath @((Get-Date -Format s),"Library",$site.Url,$libTitle,"ResolveDrive",$_.Exception.Message)
            Write-Host ("    Cannot resolve drive: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
            continue
        }
        $driveId = $drive.id

        $perLibCount = 0

        try {
            foreach ($item in Get-DriveFilesRecursive $driveId) {

                $filesSeen++
                $perLibCount++

                if ($MaxFilesPerLibrary -gt 0 -and $perLibCount -gt $MaxFilesPerLibrary) {
                    Write-Host ("    [cap] reached MaxFilesPerLibrary={0}" -f $MaxFilesPerLibrary) -ForegroundColor DarkYellow
                    break
                }

                $itemId   = [string]$item.id
                $itemName = [string]$item.name
                $itemPath = ""
                if ($item.PSObject.Properties.Name -contains "parentReference" -and $null -ne $item.parentReference -and
                    $item.parentReference.PSObject.Properties.Name -contains "path") {
                    $itemPath = ([string]$item.parentReference.path) + "/" + $itemName
                } else {
                    $itemPath = $itemName
                }

                $key = "$driveId|$itemId"
                if ($processedItems.ContainsKey($key)) { continue }

                $prev = Get-ItemLabelId $item
                $prevName = Resolve-LabelName $prev
                $action = ""
                $detail = ""

                if ($prev) {
                    # File already has a label — never overwrite in this script.
                    if ($prev.ToLower() -eq $DefaultLabelId.ToLower()) {
                        $action = "AlreadyLabeled"
                    } else {
                        $action = "SkippedExistingLabel"
                        $detail = "File already labeled with a different label; this script never overwrites. Use Apply-M365FileSensitivityLabel-PnPOnly.ps1 -OverwriteExisting to replace."
                    }
                    $filesSkippedExisting++
                }
                else {
                    if ($PSCmdlet.ShouldProcess("$itemPath", "Assign default sensitivity label $DefaultLabelId")) {
                        try {
                            Apply-Label $driveId $itemId $DefaultLabelId $AssignmentMethod
                            $action = "Labeled"
                            $filesLabeled++
                        } catch {
                            $action = "Error"
                            $detail = $_.Exception.Message
                            Write-CsvRow $ErrorCsvPath @((Get-Date -Format s),"File",$site.Url,$libTitle,"AssignSensitivityLabel",$_.Exception.Message)
                        }
                    } else {
                        $action = "WouldLabel"
                        $detail = "WhatIf"
                    }
                }

                Write-CsvRow $OutputCsvPath @(
                    (Get-Date -Format s),
                    $site.Url,
                    $libTitle,
                    $driveId,
                    $itemId,
                    $itemPath,
                    $prev,
                    $prevName,
                    $DefaultLabelId,
                    $targetLabelName,
                    $action,
                    $detail
                )

                if (($filesSeen % 100) -eq 0) {
                    Write-Host ("    ... files seen={0}, labeled={1}, skipped={2}" -f $filesSeen, $filesLabeled, $filesSkippedExisting) -ForegroundColor DarkGreen
                }
            }
        }
        catch {
            Write-CsvRow $ErrorCsvPath @((Get-Date -Format s),"Library",$site.Url,$libTitle,"EnumerateFiles",$_.Exception.Message)
            Write-Host ("    Enumeration error: {0}" -f $_.Exception.Message) -ForegroundColor Red
            continue
        }
    }

    if (($siteIndex % 10) -eq 0) {
        Write-Host ("--- Progress: {0}/{1} sites, files seen={2}, labeled={3}, skipped={4} ---" -f $siteIndex, $totalSites, $filesSeen, $filesLabeled, $filesSkippedExisting) -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "====================================" -ForegroundColor Green
Write-Host "Apply complete"
Write-Host ("Sites processed         : {0}" -f $siteIndex)
Write-Host ("Files seen              : {0}" -f $filesSeen)
Write-Host ("Files labeled (default) : {0}" -f $filesLabeled)
Write-Host ("Files skipped (labeled) : {0}" -f $filesSkippedExisting)
Write-Host ("Audit log               : {0}" -f $OutputCsvPath)
Write-Host ("Error log               : {0}" -f $ErrorCsvPath)
Write-Host "===================================="
