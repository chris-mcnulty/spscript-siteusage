<#
.SYNOPSIS
Applies a Microsoft Purview sensitivity label to existing files in every
SharePoint document library, using the same SPO-Inventory Entra app +
certificate as the inventory and library-default scripts.

.DESCRIPTION
By default, for each document library this script reads the library's
DefaultSensitivityLabelForLibrary and applies that label to every file
in the library that is not already labeled with it. Pass -LabelId to
force a single label across all libraries instead.

Labels are assigned via the Microsoft Graph endpoint
  POST /drives/{drive-id}/items/{item-id}/assignSensitivityLabel
called through Invoke-PnPGraphMethod, so the same Entra app must also
have the following GRAPH Application permissions granted and
admin-consented:

  - Files.ReadWrite.All            (to enumerate driveItems and assign labels)
  - Sites.Read.All                 (to resolve site -> drive)
  - InformationProtectionPolicy.Read.All   (to resolve label names; optional)

Microsoft's recommended approach for backfilling sensitivity labels to
existing content at scale is a Purview auto-labeling policy. Use this
script when that is not viable (small tenants, targeted libraries,
incremental rollout, etc.).

.NOTES
Throttling: 429 / 503 responses are retried. The Retry-After response
            header is honored when surfaced by the underlying call;
            otherwise the script falls back to exponential backoff
            (2,4,8,... capped at 60s) for up to 5 attempts.
Resume:     re-running with -Resume skips files already in a terminal
            state (Labeled or AlreadyLabeled) in $OutputCsvPath.
            Non-terminal rows (WouldLabel from -WhatIf,
            SkippedExistingLabel from a run without -OverwriteExisting)
            are intentionally re-evaluated on the next run.
#>
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

    # If supplied, this label is applied to every file in every library.
    # If omitted, each file inherits the library's current
    # DefaultSensitivityLabelForLibrary (libraries with no default are skipped).
    # Mutually exclusive with -LabelName.
    [string]$LabelId,

    # Alternative to -LabelId: specify the label by its display name. Requires -LabelOwnerUpn
    # so the script can resolve the name via Get-PnPAvailableSensitivityLabel. Mutually
    # exclusive with -LabelId.
    [string]$LabelName,

    # standard | privileged | auto
    [ValidateSet("standard","privileged","auto")]
    [string]$AssignmentMethod = "standard",

    # If $true, replace a file's existing different label. Default: only label files that are unlabeled.
    [switch]$OverwriteExisting,

    [string]$OutputCsvPath = ".\M365_FileLabel_Apply.csv",
    [string]$ErrorCsvPath  = ".\M365_FileLabel_Apply_Errors.csv",

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
# Validate forced label (if any) — -LabelName resolves to a GUID after the catalog loads
# --------------------------
if ($LabelId -and $LabelName) {
    throw "Specify -LabelId or -LabelName, not both."
}
if ($LabelName -and -not $LabelOwnerUpn) {
    throw "-LabelName requires -LabelOwnerUpn so the label catalog can be loaded and the name resolved."
}
if ($LabelId) {
    $g = [Guid]::Empty
    if (-not [Guid]::TryParse($LabelId, [ref]$g)) {
        throw "LabelId '$LabelId' is not a valid GUID."
    }
    $LabelId = $g.ToString()
}

# --------------------------
# CSV setup
# --------------------------
$headers = @(
    "Timestamp","SiteUrl","LibraryTitle","DriveId","ItemId","ItemPath",
    "PreviousLabelId","TargetLabelId","TargetLabelName","Action","Detail"
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

# Resume cache: keyed by DriveId|ItemId. Only terminal actions are cached so
# that a re-run with -OverwriteExisting (or a real run following a -WhatIf
# dry run) re-evaluates non-terminal rows like WouldLabel / SkippedExistingLabel.
$processedItems = @{}
if ($Resume -and (Test-Path $OutputCsvPath)) {
    Import-Csv $OutputCsvPath | ForEach-Object {
        if ($_.Action -in @("Labeled","AlreadyLabeled")) {
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
# Resolve -LabelName -> GUID, or pre-flight-validate -LabelId against the catalog
# --------------------------
function Format-LabelCatalog($map) {
    if ($map.Count -eq 0) { return "  (catalog empty)" }
    ($map.GetEnumerator() | Sort-Object Value | ForEach-Object { "  - $($_.Value)  ($($_.Key))" }) -join "`n"
}

if ($LabelName) {
    if ($labelMap.Count -eq 0) {
        throw "Could not load the label catalog, so -LabelName '$LabelName' cannot be resolved. Check Graph permissions (InformationProtectionPolicy.Read.All, User.Read.All) and that -LabelOwnerUpn '$LabelOwnerUpn' is a licensed user in this tenant."
    }
    $match = $labelMap.GetEnumerator() | Where-Object { $_.Value -ieq $LabelName } | Select-Object -First 1
    if (-not $match) {
        throw "Label name '$LabelName' not found in tenant catalog.`nAvailable labels:`n$(Format-LabelCatalog $labelMap)"
    }
    $LabelId = [string]$match.Key
    Write-Host ("Resolved -LabelName '{0}' -> {1}" -f $LabelName, $LabelId) -ForegroundColor Green
}
elseif ($LabelId -and $labelMap.Count -gt 0) {
    # We have both a forced ID and a catalog -> fail fast on bogus GUIDs before any writes.
    if (-not $labelMap.ContainsKey($LabelId.ToLower())) {
        throw "LabelId '$LabelId' is not in the tenant label catalog. Aborting before any writes.`nAvailable labels:`n$(Format-LabelCatalog $labelMap)"
    }
}

function Resolve-LabelName([string]$id) {
    if ([string]::IsNullOrWhiteSpace($id)) { return "" }
    $k = $id.ToLower()
    if ($labelMap.ContainsKey($k)) { return $labelMap[$k] }
    return ""
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
                # Honor Retry-After when present on the response; otherwise back off exponentially.
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
                    # Fallback: also try to parse "Retry-After: N" from the exception text itself.
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
    # /sites/{siteId}/lists/{listId}/drive returns the document library drive
    $endpoint = "v1.0/sites/$graphSiteId/lists/$listId/drive"
    return Invoke-Graph -Url $endpoint
}

function Get-DriveFilesRecursive($driveId) {
    # BFS over folders
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
                # Strip Graph base URL prefix so Invoke-PnPGraphMethod accepts the relative form
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
$totalSites = $sites.Count
$siteIndex  = 0
$filesSeen  = 0
$filesLabeled = 0
$filesSkipped = 0

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
        $lists = Get-PnPList -Connection $siteConn -Includes "Id","Title","Hidden","BaseTemplate","RootFolder","DefaultSensitivityLabelForLibrary"
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

        # Resolve target label for this library
        $libDefaultLabel = ""
        if ($lib.PSObject.Properties.Name -contains "DefaultSensitivityLabelForLibrary") {
            $libDefaultLabel = [string]$lib.DefaultSensitivityLabelForLibrary
        }

        $targetLabel = if ($LabelId) { $LabelId } else { $libDefaultLabel }

        if (-not $targetLabel) {
            Write-Host ("    [skip-no-label] {0}" -f $libTitle) -ForegroundColor DarkGray
            continue
        }

        $targetLabelName = Resolve-LabelName $targetLabel

        Write-Host ("  Library: {0}  ->  {1}" -f $libTitle, ($targetLabelName ? $targetLabelName : $targetLabel)) -ForegroundColor Cyan

        # Resolve drive
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
                $action = ""
                $detail = ""

                $alreadyMatches = ($prev -and ($prev.ToLower() -eq $targetLabel.ToLower()))
                $hasDifferent   = ($prev -and -not $alreadyMatches)

                if ($alreadyMatches) {
                    $action = "AlreadyLabeled"
                    $filesSkipped++
                }
                elseif ($hasDifferent -and -not $OverwriteExisting) {
                    $action = "SkippedExistingLabel"
                    $detail = "File already labeled with a different label; pass -OverwriteExisting to replace."
                    $filesSkipped++
                }
                else {
                    if ($PSCmdlet.ShouldProcess("$itemPath", "Assign sensitivity label $targetLabel")) {
                        try {
                            Apply-Label $driveId $itemId $targetLabel $AssignmentMethod
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
                    $targetLabel,
                    $targetLabelName,
                    $action,
                    $detail
                )

                if (($filesSeen % 100) -eq 0) {
                    Write-Host ("    ... files seen={0}, labeled={1}, skipped={2}" -f $filesSeen, $filesLabeled, $filesSkipped) -ForegroundColor DarkGreen
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
        Write-Host ("--- Progress: {0}/{1} sites, files seen={2}, labeled={3}, skipped={4} ---" -f $siteIndex, $totalSites, $filesSeen, $filesLabeled, $filesSkipped) -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "====================================" -ForegroundColor Green
Write-Host "Apply complete"
Write-Host ("Sites processed : {0}" -f $siteIndex)
Write-Host ("Files seen      : {0}" -f $filesSeen)
Write-Host ("Files labeled   : {0}" -f $filesLabeled)
Write-Host ("Files skipped   : {0}" -f $filesSkipped)
Write-Host ("Audit log       : {0}" -f $OutputCsvPath)
Write-Host ("Error log       : {0}" -f $ErrorCsvPath)
Write-Host "===================================="
