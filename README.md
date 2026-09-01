# 365_Adminscript

🛠 An advanced PowerShell tool for unified Microsoft 365 administration (Entra ID, Exchange Online, Intune).

## Key Features

- 🔐 Graph & Exchange authentication via device code flow
- 🧾 Full audit export to Excel (Users, MFA, Devices, Licenses, Domains, OneDrive)
- 📊 Intelligent Graph API throttling management
- 🧼 Secure logging that masks sensitive UPNs + portable AES key encryption (key file in separate `Sherl0ck_Secrets` directory with explicit NTFS ACL)
- 📁 MSAL authentication conflict handling (Graph <> Exchange)
- 🛡 PowerShell modules pinned and verified via official PSGallery
- 🧱 Modular and testable architecture with Pester
- 👤 Identity & Security audit (MFA status, Conditional Access, OAuth apps, RBAC)

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
  - AuditLog.Read.All
  - RoleManagement.Read.Directory

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

Required modules will be downloaded automatically from the official PowerShell Gallery.

## Project Structure

```
365_Adminscript/
├── Scripts/
│   └── Sherl0ck_v4.1.ps1          # Main entry point
├── Modules/
│   ├── Sherl0ck.Utils.psm1        # Utilities and helpers
│   ├── Sherl0ck.UI.psm1           # User interface and logging
│   ├── Sherl0ck.Auth.psm1         # Graph/Exchange authentication
│   └── Sherl0ck.Audit.psm1        # Audit collection and export
├── Config/
│   └── config.json.example        # External configuration template
├── Tests/
│   ├── Sherl0ck.Tests.Utils.ps1
│   ├── Sherl0ck.Tests.UI.ps1
│   ├── Sherl0ck.Tests.Auth.ps1
│   └── Sherl0ck.Tests.Audit.ps1
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
- **I1** : Configuration externalization (`Sherl0ck_Config.json`) ✅
- **I2** : Multilingual support (fr-FR / en-US) ❌ To do
- **I3** : Integrated documentation (help XML comments) ✅

## License

MIT — See LICENSE for details.

## Credits

Developed by g33ky00 — based on the Sherl0ck v4.1 tool.
