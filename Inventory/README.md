# Executive summary (what you need to run this)

You are creating a **non-human identity (app)** that can:

* Read all SharePoint sites
* Read sensitivity labels
* Authenticate with a certificate (no prompts)

That requires:

* ✅ Entra App registration
* ✅ Certificate (PFX file)
* ✅ Permissions granted + admin consent

***

# Where each value comes from

## 1) `TenantName`

👉 Your SharePoint tenant prefix

**How to find it:**

* Go to SharePoint Admin Center
* URL will look like:
  ```
  https://contoso-admin.sharepoint.com
  ```

👉 Your value = `contoso`

***

## 2) `ClientId`

👉 The App Registration ID (GUID)

**How to get it:**

* Go to:  
  **Microsoft Entra Admin Center → App registrations**
* Create new app OR open existing
* Copy:

```
Application (client) ID
```

Example:

```
00000000-0000-0000-0000-000000000000
```

***

## 3) `TenantId`

👉 Your Entra tenant ID (GUID)

**How to get it:**

* Same screen as above

Copy:

```
Directory (tenant) ID
```

Example:

```
11111111-1111-1111-1111-111111111111
```

***

## 4) `CertificatePath`

👉 The file path to your **.pfx certificate**

Example:

```
C:\certs\spo-inventory.pfx
```

***

## 5) `$pwd` / `CertificatePassword`

👉 The password you set when creating the PFX

Used here:

```powershell
$pwd = Read-Host "PFX password" -AsSecureString
```

***

# If you don’t have a certificate yet (most likely)

Here’s the **fastest working approach** (local dev or pilot):

***

## Create a certificate (PowerShell)

```powershell
$cert = New-SelfSignedCertificate `
  -Subject "CN=spo-inventory" `
  -CertStoreLocation "Cert:\CurrentUser\My" `
  -KeyExportPolicy Exportable `
  -KeySpec Signature `
  -KeyLength 2048 `
  -HashAlgorithm sha256

Export-PfxCertificate `
  -Cert $cert `
  -FilePath "C:\certs\spo-inventory.pfx" `
  -Password (ConvertTo-SecureString "YourStrongPassword!" -AsPlainText -Force)
```

👉 Now you have:

* `.pfx` file → your `CertificatePath`
* password → your `$pwd`

***

## Upload certificate to app registration

* Go to:
  **Entra → App registrations → Your App → Certificates & secrets**
* Upload:
  * The **public key (.cer) OR export from cert store**

***

## Export public cert (if needed)

```powershell
Export-Certificate `
  -Cert $cert `
  -FilePath "C:\certs\spo-inventory.cer"
```

Upload that `.cer` to Entra.

***

# Permissions you MUST set (or script fails silently)

In your App Registration:

## API Permissions

Add:

### ✅ SharePoint

```
Sites.FullControl.All   (Application)
```

Then:
👉 Click **Grant admin consent**

***

## Role assignment (often missed)

* Go to:
  **Entra → Roles and administrators**
* Add your app to:
  ```
  SharePoint Administrator
  ```

👉 Without this, `Get-SPOSite` will fail in real tenants.

***

# What your final working values look like

```powershell
$pwd = Read-Host "PFX password" -AsSecureString

.\Export-M365SensitivityLabelInventory.ps1 `
  -TenantName "contoso" `
  -ClientId "b183408f-123b-48df-b886-21003add50c1" `
  -TenantId  "icewolfch.onmicrosoft.com" `
  -CertificatePath "C:\certs\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -Resume
```

***

# What your actual working values look like for Synozur

```powershell
$pwd = Read-Host "PFX password" -AsSecureString

.\Export-M365SensitivityLabelInventory.ps1 `
  -TenantName "synozur" `
  -ClientId "df7b6a64-8d68-4e75-b15d-2cffc07cb554" `
  -TenantId  "synozur.onmicrosoft.com" `
  -CertificatePath ".\cert\spo-inventory.pfx" `
  -CertificatePassword $pwd `
  -Resume
```

***

# Practical guidance (from your use case)

Given what you’re doing (tracking labeling across thousands of sites for Copilot governance):

* This app becomes your **core governance identity**
* You will reuse it for:
  * label enforcement
  * inventory scans
  * adaptive scope validation
  * Copilot readiness scoring

👉 This is basically the “mini Zenith engine” identity you’ll run repeatedly.

***

# Common failure modes (so you don’t burn time)

* ❌ Missing admin consent on API permissions
* ❌ Not using **Application** permissions (using Delegated instead)
* ❌ Certificate uploaded incorrectly (wrong format)
* ❌ App not added to SharePoint Administrator role
* ❌ Wrong tenant format (`TenantId` vs tenant domain mismatch)

***

# Bottom line

You are wiring up three things:

* Identity → `ClientId`, `TenantId`
* Trust → certificate (.pfx + password)
* Authority → SharePoint permissions + admin consent

Once those are correct, your script will:

* run unattended
* scale to 2k+ sites
* and become production-stable


