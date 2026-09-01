# 365_Adminscript

🛠 An advanced PowerShell tool for unified Microsoft 365 administration (Entra ID, Exchange Online, Intune).

## Key Features

- 🔐 Graph & Exchange authentication via device code flow
- 🧾 Full audit export to Excel (Users, MFA, Devices, Licenses, Domains, OneDrive)
- 📊 Intelligent Graph API throttling management
- 🧼 Secure logging that masks sensitive UPNs + portable AES key encryption (key file in separate `Sherl0ck_Secrets` directory with explicit NTFS ACL)
- 📁 MSAL authentication conflict handling (Graph <> Exchange)
- 🛡 PowerShell modules pinned and verified via official PSGallery (MinimumVersion + SHA-512 hash comparison)
- 🧱 Modular and testable architecture with Pester (4 modules + entry point + 4 test files)
- 👤 Identity & Security audit (MFA status, Conditional Access policies, OAuth applications, RBAC)

## OAuth Audit Modes (H3)

- **ReadOnly** (default) — scopes restricted to read-only operations, minimizing the attack surface
- **ReadWrite** — reserved for future write operations (currently warns: no write functions implemented)

```powershell
.\Scripts\Sherl0ck_v4.1.ps1                          # ReadOnly mode (default)
.\Scripts\Sherl0ck_v4.1.ps1 -AuditMode ReadWrite   # ReadWrite mode
```

## Prerequisites

- PowerShell 5.1 or higher
- Microsoft 365 administrator account
- Access to the following scopes:
  - Policy.Read.All
  - User.Read.All
  - Organization.Read.All
  - Domain.Read.All
  - Device.Read.All
  - Directory.Read.All
  - Reports.Read.All
  - Sites.Read.All
  - Files.Read.All
  - AuditLog.Read.All
  - RoleManagement.Read.Directory
  - Application.Read.All

## Installation

Clone the repository:
```bash
git clone https://github.com/g33ky00/365_Adminscript.git
cd 365_Adminscript
```

Run the main script:
```powershell
.\Scripts\Sherl0ck_v4.1.ps1
```

Required PowerShell modules are downloaded automatically from the official PowerShell Gallery:
- `Microsoft.Graph.Authentication` (MinimumVersion: 1.9.3)
- `ExchangeOnlineManagement` (MinimumVersion: 3.2.0)
- `ImportExcel` (MinimumVersion: 7.8.0)

To skip automatic module installation (air-gapped/secure environments):
```powershell
.\Scripts\Sherl0ck_v4.1.ps1 -SkipModuleInstall
```

## Main Menu

```
[ 1 ] IDENTITY & SECURITY   (MFA, Conditional Access, OAuth apps, RBAC)
[ 2 ] EXCHANGE ONLINE        (Mailboxes, Quotas, Redirects)
[ 3 ] M365 AUDIT             (Complete collection + Excel Export)
[ 4 ] SESSION LOGS          (Logs & errors)
[ 0 ] DISCONNECT & QUIT
```

## Project Structure

```
365_Adminscript/
├── Scripts/
│   └── Sherl0ck_v4.1.ps1          # Main entry point (no functions defined, imports modules)
├── Modules/
│   ├── Sherl0ck.Utils.psm1        # Utilities: Convert-EXOSizeToGB, Convert-BytesToGB, Invoke-SafeOpen, Get-UniqueFilePath
│   ├── Sherl0ck.UI.psm1           # User interface: Mask-SensitiveData, Add-SessionLog, Show-SessionLogs
│   ├── Sherl0ck.Auth.psm1         # Graph/Exchange auth: Connect-O365Core, Connect-O365Exchange, Verify-TrustedModule, Invoke-BrowserPrivate, Get-REQUIRED_MODULES
│   └── Sherl0ck.Audit.psm1        # Audit: Get-GraphData, Export-OneDriveUsage, Export-FullAuditExcel, Show-MenuAudit, Get-MFAStatus, Get-ConditionalAccessPolicies, Get-OAuthApplications, Get-RoleBasedAccess, Export-IdentitySecurityExcel
├── Config/
│   └── config.json.example        # Configuration template (SkipModuleInstall, ModuleSource)
├── Tests/
│   ├── Sherl0ck.Tests.Utils.ps1    # Pester tests for 4 Utils functions
│   ├── Sherl0ck.Tests.UI.ps1       # Pester tests for 3 UI functions
│   ├── Sherl0ck.Tests.Auth.ps1     # Pester tests for 4 Auth functions
│   └── Sherl0ck.Tests.Audit.ps1    # Pester tests for 5 Identity + 4 core Audit functions
├── README.md
└── LICENSE
```

## Tests

Use Pester to validate each module independently:
```powershell
Invoke-Pester -Path ./Tests/
```

## Applied Corrections (Sherl0ck v4.1)

- **H1** : MSAL conflict handling (Graph <> Exchange) — protects against authentication context collisions ✅
- **H2** : Secure logging — masks sensitive UPNs via `Mask-SensitiveData` ✅
- **H3** : OAuth scope separation — ReadOnly/ReadWrite modes with dynamic `AuditMode` parameter ✅
  - **Point 6** : ReadWrite privileged scopes **removed** — unused by any function. ReadWrite mode still accepted as parameter but warns that no write operations are available.
- **H4** : Module pinning with `MinimumVersion` + PSGallery source verification via `Verify-TrustedModule` ✅
- **M1** : UPN masking (`***`) + **portable AES key encryption** (key file in separate `%LOCALAPPDATA%\Sherl0ck_Secrets\` directory with explicit NTFS ACL — key never co-located with logs) via `ConvertFrom-SecureString -Key` ✅
- **M2** : `Get-UniqueFilePath` — checks existence before write + increments filename (`_1`, `_2`, ...) ✅
- **M3** : `Find-Module` source + version verification + **real SHA-512 hash comparison** via `Save-Module` + `Get-FileHash` vs `PackageHash` + `-SkipModuleInstall` switch ✅
- **L1** : Modular refactoring into PowerShell modules (Utils, UI, Auth, Audit) ✅
- **L2** : Extended unit tests (Pester) — test stubs for all 4 modules ✅
- **L3** : Simulation mode `-WhatIf` ❌ To do
- **FIX-3** : Browser fallback (Edge→Firefox→Chrome) + explicit error if none available ✅
- **FIX-6** : ReadWrite privileged scopes removed — no function consumes them ✅
- **Point 8** : Identity & Security audit stubs (MFA, CA, OAuth, RBAC) + extended menu [1] ✅
- **Point 4** : Graph $batch (POST /v1.0/$batch, 20 req/batch) for Export-OneDriveUsage — reduces API calls by ~20x ✅
- **Point 2** : Set-StrictMode -Off → -Version 3.0 (enables variable/property strict checking) ✅
- **Point 1** : DPAPI-only encryption → portable AES key (key file in separate `Sherl0ck_Secrets` directory + explicit NTFS ACL) ✅
- **FIX-7** : Silent error loss in Export-OneDriveUsage → Add-SessionLog + ODStats entry ✅
- **Audit findings** : 15 empty catch blocks replaced with `Add-SessionLog` (error tracing) ✅
- **I1** : Configuration externalization (`Config/config.json.example`) ✅
- **I2** : Multilingual support (fr-FR / en-US) ❌ To do
- **Point 9** : REQUIRED_MODULE_VERSIONS converted to readonly hash via `Set-Variable -Option ReadOnly` + `Get-REQUIRED_MODULES()` getter function (no longer exported as mutable variable) ✅
- **I3** : Integrated documentation (help XML comments) ✅

## License

MIT — See LICENSE for details.

## Credits

Developed by g33ky00 — based on the Sherl0ck v4.1 tool.
