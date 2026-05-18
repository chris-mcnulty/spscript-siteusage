# =====================================
# Connect to SharePoint Online Admin
# =====================================

# Replace with your tenant name
$AdminUrl = "https://synozur-admin.sharepoint.com"

Connect-SPOService -Url $AdminUrl


# =====================================
# Check current AIP integration status
# =====================================

Write-Host "Checking EnableAIPIntegration status..." -ForegroundColor Cyan

$tenant = Get-SPOTenant
$status = $tenant.EnableAIPIntegration

Write-Host "EnableAIPIntegration current value:" $status -ForegroundColor Yellow


# =====================================
# Enable if not already enabled
# =====================================

if ($status -eq $false) {
    Write-Host "Enabling EnableAIPIntegration..." -ForegroundColor Cyan
    
    Set-SPOTenant -EnableAIPIntegration $true
    
    # Re-check after setting
    $tenant = Get-SPOTenant
    Write-Host "New EnableAIPIntegration value:" $tenant.EnableAIPIntegration -ForegroundColor Green
}
else {
    Write-Host "EnableAIPIntegration is already enabled. No changes made." -ForegroundColor Green
}


# =====================================
# Optional: Output related labeling state
# =====================================

Write-Host ""
Write-Host "Additional tenant labeling-related properties:" -ForegroundColor Cyan

Get-SPOTenant | Select `
    EnableAIPIntegration,
    DisplayName,
    # helpful context checks (won’t always correlate but useful)
    ConditionalAccessPolicy,
    DisableCustomAppAuthentication
