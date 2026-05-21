<#
.SYNOPSIS
Creates an Entra app registration + service principal, generates a self-signed certificate (PFX+CER),
uploads the cert to the app (keyCredentials), grants the full application-permission set used by every
script in this folder (inventory, library-default label, file label, site container label), and outputs
the exact values needed to run the scripts.

GRANTS (Application, admin-consented):
  SharePoint Online:
    - Sites.FullControl.All                        (inventory enumeration + Set-PnPList)
  Microsoft Graph:
    - Sites.FullControl.All                        (Set-PnPSite -SensitivityLabel)
    - Files.ReadWrite.All                          (file-level assignSensitivityLabel)
    - InformationProtectionPolicy.Read.All         (resolve label GUIDs -> display names)
    - User.Read.All                                (scope label policy lookup to a user under app-only)
    - Group.ReadWrite.All                          (sync container label to backing M365 group; pass -SkipGroupReadWrite to omit)

REQUIRES
- PowerShell 7+ recommended
- Microsoft.Graph PowerShell module
  Install-Module Microsoft.Graph -Scope CurrentUser

NOTES
- You will sign in once interactively to Microsoft Graph with sufficient rights to create apps and assign app roles.
- SharePoint Online Management Shell supports cert-based app-only auth for Connect-SPOService using -ClientId -TenantId -CertificatePath -CertificatePassword. 【3-dbeb06】
- Graph supports programmatically adding certs to apps. 【2-db332b】
- Graph supports programmatically creating apps and service principals. 【1-3ef0ec】
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$AppName,

    # Your tenant’s GUID (Directory (tenant) ID)
    [Parameter(Mandatory=$true)]
    [string]$TenantId,

    # Where to write the .pfx and .cer
    [Parameter(Mandatory=$false)]
    [string]$OutDir = ".\cert",

    # PFX password to protect the private key export
    [Parameter(Mandatory=$true)]
    [SecureString]$PfxPassword,

    # Cert validity
    [Parameter(Mandatory=$false)]
    [int]$ValidYears = 2,

    # Optional: attempt to assign Entra directory role "SharePoint Administrator" to the service principal
    [Parameter(Mandatory=$false)]
    [switch]$AssignSharePointAdminRole,

    # Skip granting Graph Group.ReadWrite.All. Container label assignment still works for the site
    # itself, but won't synchronize to the backing M365 Group object for Teams/Group-connected sites.
    [Parameter(Mandatory=$false)]
    [switch]$SkipGroupReadWrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------- Prep ----------
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

# Microsoft Graph connection scopes:
# - create/update apps: Application.ReadWrite.All
# - assign app roles: AppRoleAssignment.ReadWrite.All
# - optionally assign directory roles: RoleManagement.ReadWrite.Directory
$scopes = @(
    "Application.ReadWrite.All",
    "AppRoleAssignment.ReadWrite.All"
)
if ($AssignSharePointAdminRole) {
    $scopes += "RoleManagement.ReadWrite.Directory"
}

Import-Module Microsoft.Graph -ErrorAction Stop
Select-MgProfile -Name "v1.0" | Out-Null

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -TenantId $TenantId -Scopes $scopes | Out-Null

# ---------- 1) Create app registration ----------
Write-Host "Creating application: $AppName" -ForegroundColor Cyan
$app = New-MgApplication -DisplayName $AppName

# ---------- 2) Create service principal ----------
Write-Host "Creating service principal..." -ForegroundColor Cyan
$sp = New-MgServicePrincipal -AppId $app.AppId

# ---------- 3) Generate certificate + export ----------
# Creates cert in CurrentUser\My and exports:
# - CER (public key) for reference
# - PFX (private+public) for your scripts
Write-Host "Generating self-signed certificate..." -ForegroundColor Cyan
$cert = New-SelfSignedCertificate `
    -Subject ("CN=" + $AppName) `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -HashAlgorithm "SHA256" `
    -NotAfter (Get-Date).AddYears($ValidYears)

$cerPath = Join-Path $OutDir ($AppName + ".cer")
$pfxPath = Join-Path $OutDir ($AppName + ".pfx")

Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $PfxPassword | Out-Null

# ---------- 4) Upload certificate to the app (keyCredentials) ----------
# Graph expects the public cert bytes in keyCredentials.key (base64).
Write-Host "Uploading certificate public key to the app registration..." -ForegroundColor Cyan
$cerBytes = [System.IO.File]::ReadAllBytes($cerPath)
$now = Get-Date
$keyCred = @{
    type  = "AsymmetricX509Cert"
    usage = "Verify"
    key   = $cerBytes
    displayName = ($AppName + " cert")
    startDateTime = $now.ToString("o")
    endDateTime   = $cert.NotAfter.ToString("o")
}

# Append to existing key credentials
$appCurrent = Get-MgApplication -ApplicationId $app.Id
$existing = @()
if ($appCurrent.KeyCredentials) { $existing = @($appCurrent.KeyCredentials) }

Update-MgApplication -ApplicationId $app.Id -KeyCredentials ($existing + $keyCred)

# ---------- 5) Grant application permissions (admin-consented) ----------
# Resource service principals are located by their well-known appIds:
#   00000003-0000-0ff1-ce00-000000000000 — Office 365 SharePoint Online
#   00000003-0000-0000-c000-000000000000 — Microsoft Graph
$permissions = @(
    @{ ResourceAppId = "00000003-0000-0ff1-ce00-000000000000"; ResourceName = "SharePoint Online";  Role = "Sites.FullControl.All" }
    @{ ResourceAppId = "00000003-0000-0000-c000-000000000000"; ResourceName = "Microsoft Graph";    Role = "Sites.FullControl.All" }
    @{ ResourceAppId = "00000003-0000-0000-c000-000000000000"; ResourceName = "Microsoft Graph";    Role = "Files.ReadWrite.All" }
    @{ ResourceAppId = "00000003-0000-0000-c000-000000000000"; ResourceName = "Microsoft Graph";    Role = "InformationProtectionPolicy.Read.All" }
    @{ ResourceAppId = "00000003-0000-0000-c000-000000000000"; ResourceName = "Microsoft Graph";    Role = "User.Read.All" }
)
if (-not $SkipGroupReadWrite) {
    $permissions += @{ ResourceAppId = "00000003-0000-0000-c000-000000000000"; ResourceName = "Microsoft Graph"; Role = "Group.ReadWrite.All" }
}

Write-Host "Granting application permissions..." -ForegroundColor Cyan

$resourceSpCache = @{}
$existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id

foreach ($perm in $permissions) {
    $resAppId = $perm.ResourceAppId

    if (-not $resourceSpCache.ContainsKey($resAppId)) {
        $resSp = Get-MgServicePrincipal -Filter "appId eq '$resAppId'"
        if (-not $resSp) { throw "Could not find service principal for $($perm.ResourceName) (appId $resAppId) in this tenant." }
        $resourceSpCache[$resAppId] = $resSp
    }
    $resSp = $resourceSpCache[$resAppId]

    $role = $resSp.AppRoles |
        Where-Object { $_.Value -eq $perm.Role -and ($_.AllowedMemberTypes -contains "Application") } |
        Select-Object -First 1
    if (-not $role) { throw "Could not find AppRole '$($perm.Role)' on $($perm.ResourceName) service principal." }

    $already = $existingAssignments | Where-Object { $_.ResourceId -eq $resSp.Id -and $_.AppRoleId -eq $role.Id }
    if ($already) {
        Write-Host ("  [=] {0} :: {1} (already granted)" -f $perm.ResourceName, $perm.Role) -ForegroundColor DarkGray
        continue
    }

    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $sp.Id `
        -PrincipalId $sp.Id `
        -ResourceId $resSp.Id `
        -AppRoleId $role.Id | Out-Null
    Write-Host ("  [+] {0} :: {1}" -f $perm.ResourceName, $perm.Role) -ForegroundColor Green
}

# ---------- 6) Optional: assign Entra role "SharePoint Administrator" ----------
if ($AssignSharePointAdminRole) {
    Write-Host "Attempting to assign Entra role: SharePoint Administrator (tenant scope)..." -ForegroundColor Cyan

    $roleDef = Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq 'SharePoint Administrator'" | Select-Object -First 1
    if (-not $roleDef) {
        Write-Host "Warning: Could not find role definition 'SharePoint Administrator'. Skipping." -ForegroundColor Yellow
    }
    else {
        New-MgRoleManagementDirectoryRoleAssignment -BodyParameter @{
            principalId      = $sp.Id
            roleDefinitionId = $roleDef.Id
            directoryScopeId = "/"
        } | Out-Null
    }
}

# ---------- Output: values you need for inventory run ----------
Write-Host "`n=== COPY THESE VALUES INTO YOUR INVENTORY RUN ===" -ForegroundColor Green

$result = [pscustomobject]@{
    ApplicationName = $AppName
    ClientId        = $app.AppId
    TenantId        = $TenantId
    CertificatePath = (Resolve-Path $pfxPath).Path
    CertificateCer  = (Resolve-Path $cerPath).Path
    Notes           = "Use TenantName from your SharePoint admin URL prefix (e.g., contoso from https://contoso-admin.sharepoint.com)."
}

$result | Format-List

Write-Host "`nExample inventory run:" -ForegroundColor Cyan
Write-Host '$pwd = Read-Host "PFX password" -AsSecureString'
Write-Host ('.\Export-M365SensitivityLabelInventory.ps1 -TenantName "<tenantPrefix>" -ClientId "' + $app.AppId + '" -TenantId "' + $TenantId + '" -CertificatePath "' + (Resolve-Path $pfxPath).Path + '" -CertificatePassword $pwd -Resume')