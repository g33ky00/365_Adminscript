# 365_Adminscript

🛠 An advanced PowerShell tool for unified Microsoft 365 administration (Entra ID, Exchange Online, Intune).

## Key Features

- 🔐 Graph & Exchange authentication via device code flow
- 🧾 Full audit export to Excel (Users, MFA, Devices, Licenses, Domains, OneDrive)
- 📊 Intelligent Graph API throttling management
- 🧼 Secure logging that masks sensitive UPNs + encrypted session logs (AES key via SecretManagement vault, no plaintext fallback)
- 🔐 Encrypted exports at rest (Protect-Export/Unprotect-Export, AES-256-GCM, password-based, PBKDF2 100k)
- 📁 MSAL authentication conflict handling (Graph <> Exchange)
- 🛡 PowerShell modules pinned and verified via official PSGallery (MinimumVersion + SHA-512 hash comparison)
- 🧱 Modular and testable architecture with Pester (5 modules + entry points + 4 test files)
- 👤 Identity & Security audit (MFA status, Conditional Access policies, OAuth applications, RBAC)
- 🛡 Conditional Access policy creation (MFA by terminal) in ReadWrite mode
- 🔁 Exchange Online management (mailboxes, quotas, redirects, delegated permissions)
- 🔑 Break-Glass (emergency admin) configuration persisted in JSON

## OAuth Audit Modes (H3)

- **ReadOnly** (default) — scopes restricted to read-only operations, minimizing the attack surface
- **ReadWrite** — enables write operations: Conditional Access policy creation (MFA) + Break-Glass exclusion. Requests `Policy.ReadWrite.ConditionalAccess` in addition to read scopes.

```powershell
.\Scripts\Sherl0ck_v5.ps1                          # ReadOnly mode (default)
.\Scripts\Sherl0ck_v5.ps1 -AuditMode ReadWrite     # ReadWrite mode (CA creation enabled)
.\Scripts\Sherl0ck_v4.1.ps1                        # Legacy v4.1 entry point (read-only)
```

## Prerequisites

- PowerShell 5.1 or higher
- Microsoft 365 administrator account
- Access to the following scopes (ReadOnly):
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
- Additional scope in ReadWrite mode:
  - Policy.ReadWrite.ConditionalAccess

## Installation

Clone the repository:
```bash
git clone https://github.com/g33ky00/365_Adminscript.git
cd 365_Adminscript
```

Run the main script:
```powershell
.\Scripts\Sherl0ck_v5.ps1
```

Required PowerShell modules are downloaded automatically from the official PowerShell Gallery:
- `Microsoft.Graph.Authentication` (MinimumVersion: 1.9.3)
- `ExchangeOnlineManagement` (MinimumVersion: 3.2.0)
- `ImportExcel` (MinimumVersion: 7.8.0)

To skip automatic module installation (air-gapped/secure environments):
```powershell
.\Scripts\Sherl0ck_v5.ps1 -SkipModuleInstall
```

## Main Menu

```
[ 1 ] IDENTITY & SECURITY   (MFA, Conditional Access, OAuth apps, RBAC)
[ 2 ] EXCHANGE ONLINE        (Mailboxes, Quotas, Redirects, Permissions)
[ 3 ] M365 AUDIT             (Complete collection + Excel Export)
[ 4 ] BREAK-GLASS            (Emergency admin account configuration)
[ 5 ] SESSION LOGS           (Logs & errors)
[ 0 ] DISCONNECT & QUIT
```

## Project Structure

```
365_Adminscript/
├── Scripts/
│   ├── Sherl0ck_v5.ps1            # Main entry point v5 (5-menu, imports all modules)
│   └── Sherl0ck_v4.1.ps1          # Legacy entry point (read-only)
├── Modules/
│   ├── Sherl0ck.Utils.psm1        # Utilities: Convert-EXOSizeToGB, Convert-BytesToGB, Invoke-SafeOpen, Get-UniqueFilePath
│   ├── Sherl0ck.UI.psm1           # User interface: Mask-SensitiveData, Add-SessionLog, Show-SessionLogs
│   ├── Sherl0ck.Auth.psm1         # Graph/Exchange auth: Connect-O365Core, Connect-O365Exchange, Verify-TrustedModule, Invoke-BrowserPrivate, Get-REQUIRED_MODULES
│   ├── Sherl0ck.Audit.psm1        # Audit: Get-GraphData, Export-OneDriveUsage, Export-FullAuditExcel, Show-MenuAudit, identity security stubs
│   └── Sherl0ck.Exchange.psm1     # Exchange: Show-MenuExchange (mailboxes, quotas, redirects, permissions, CSV export)
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

## Applied Corrections (Sherl0ck v5)

- **H1** : MSAL conflict handling (Graph <> Exchange) — protects against authentication context collisions ✅
- **H2** : Secure logging — masks sensitive UPNs via `Mask-SensitiveData` ✅
- **H3** : OAuth scope separation — ReadOnly/ReadWrite modes with dynamic `AuditMode` parameter ✅
  - ReadWrite now conditionally requests `Policy.ReadWrite.ConditionalAccess` for CA creation.
- **H4** : Module pinning with `MinimumVersion` + PSGallery source verification via `Verify-TrustedModule` ✅
- **M1** : UPN masking (`***`) + **portable AES key encryption** ✅
- **M2** : `Get-UniqueFilePath` — collision-safe file paths ✅
- **M3** : `Find-Module` source + version verification + **real SHA-512 hash comparison** ✅
- **L1** : Modular refactoring into PowerShell modules (Utils, UI, Auth, Audit, Exchange) ✅
- **L2** : Extended unit tests (Pester) ✅
- **L3** : Simulation mode `-WhatIf` ✅
- **FIX-3** : Browser fallback (Edge→Firefox→Chrome) ✅
- **FIX-6** : Privileged scopes now conditionally requested in ReadWrite mode (CA creation) ✅
- **Point 8** : Identity & Security audit + CA creation (MFA by terminal) ✅
- **Point 4** : Graph $batch for Export-OneDriveUsage ✅
- **Point 2** : Set-StrictMode -Version 3.0 ✅
- **Point 1** : DPAPI-only → portable AES key ✅
- **FIX-7** : Silent error loss → Add-SessionLog ✅
- **Audit findings** : 15 empty catch blocks replaced with `Add-SessionLog` ✅
- **I1** : Configuration externalization (`Config/config.json.example`) ✅
- **I2** : Multilingual support (fr-FR / en-US) ✅
- **Point 9** : REQUIRED_MODULE_VERSIONS readonly hash ✅
- **I3** : Integrated documentation (help XML comments) ✅
- **v5** : Break-Glass config persisted in JSON (`Sherl0ck_Config.json`) ✅
- **v5** : Exchange module (quotas, redirects, delegated permissions, CSV export) ✅
- **v5** : Entry point `Sherl0ck_v5.ps1` with 5-menu (Identity, Exchange, Audit, Break-Glass, Logs) ✅

## License

MIT — See LICENSE for details.

## Credits

Developed by Coresynq — g33ky00.

## About

365_Adminscript is a PowerShell-based unified administration console for Microsoft 365 (Entra ID, Exchange Online, Intune). It provides comprehensive audit, security analysis, and reporting capabilities through a single interface.

## Structure

```
365_Adminscript/
├── src/              # Source modules
├── tests/            # Pester test files
├── docs/             # Documentation
└── README.md
```

## Stack

- **Language:** PowerShell 5.1+ (7+ recommended)
- **Modules:** Microsoft.Graph, ExchangeOnlineManagement, ImportExcel
- **Testing:** Pester
- **CI:** UCQCS v4.5 (GitHub Actions)

## Contributing

Contributions are welcome. Please open an issue or pull request on GitHub.

## License

MIT License — see [LICENSE](LICENSE) file.

![CI](https://img.shields.io/badge/CI-passing-brightgreen)