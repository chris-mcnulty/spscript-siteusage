<#
.SYNOPSIS
    Gets the GUID of a sensitivity label by display name via Security & Compliance PowerShell.
    Returns all labels in the tenant regardless of publishing policy scoping.
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$LabelName
)

# Ensure required module is loaded
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
}
Import-Module ExchangeOnlineManagement

# Connect to Security & Compliance PowerShell
try {
    Get-Label -ErrorAction Stop | Out-Null
    Write-Host "Already connected to IPPS." -ForegroundColor Green
} catch {
    Write-Host "Connecting to Security & Compliance PowerShell..." -ForegroundColor Yellow
    Connect-IPPSSession
}

# Get all labels
$labels = Get-Label

if (-not $labels -or $labels.Count -eq 0) {
    Write-Warning "No labels found in tenant."
    exit 1
}

Write-Host "Retrieved $($labels.Count) labels from tenant." -ForegroundColor Green

# Find the matching label
$match = $labels | Where-Object { $_.DisplayName -eq $LabelName }

if ($match) {
    Write-Host ""
    Write-Host "MATCH FOUND" -ForegroundColor Green
    Write-Host "DisplayName: $($match.DisplayName)"
    Write-Host "GUID: $($match.Guid)"
    Write-Host "ImmutableId: $($match.ImmutableId)"
    Write-Host "Priority: $($match.Priority)"
    Write-Host "ContentType: $($match.ContentType)"
    return $match.Guid
} else {
    Write-Warning "No label found with display name '$LabelName'"
    Write-Host ""
    Write-Host "Available labels in tenant:" -ForegroundColor Yellow
    $labels | Sort-Object DisplayName | ForEach-Object {
        Write-Host "  '$($_.DisplayName)' -> $($_.Guid)"
    }
    
    Write-Host ""
    Write-Host "Character-level comparison of your input:" -ForegroundColor Yellow
    $LabelName.ToCharArray() | ForEach-Object { 
        Write-Host ("  '{0}' = U+{1:X4}" -f $_, [int]$_)
    }
}