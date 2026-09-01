# 365_Adminscript

🛠 Un outil PowerShell avancé pour l'administration et l'audit unifié de Microsoft 365 (Entra ID, Exchange Online, Intune).

## Fonctionnalités clés

- 🔐 Authentification Graph & Exchange en mode device code
- 🧾 Export complet des audits vers Excel (Utilisateurs, MFA, Appareils, Licences, Domaines, OneDrive)
- 📊 Gestion intelligente du throttling Graph API
- 🧼 Journalisation sécurisée masquant les UPNs sensibles
- 📁 Gestion des conflits d'authentification MSAL (Graph <> Exchange)
- 🛡 Modules PowerShell épinglés & vérifiés via PSGallery officielle
- 🧱 Architecture modulaire et testable avec Pester

## Modes d'audit OAuth (H3)

- **ReadOnly** (défaut) : scopes restreints à la lecture seule — minimise le surface d'attaque
- **ReadWrite** : ajoute `Policy.ReadWrite.ConditionalAccess` et `User.ReadWrite.All` pour les opérations d'écriture critiques

```powershell
.\Scripts\Sherl0ck_v4.1.ps1                          # Mode ReadOnly par défaut
.\Scripts\Sherl0ck_v4.1.ps1 -AuditMode ReadWrite   # Mode ReadWrite
```

## Prérequis

- PowerShell 5.1 ou supérieur
- Compte administrateur Microsoft 365
- Accès aux scopes suivants :
  - Policy.Read.All
  - User.Read.All
  - Organization.Read.All
  - AuditLog.Read.All
  - RoleManagement.Read.Directory

## Installation

Clonez le dépôt :
```bash
git clone https://github.com/g33ky00/365_Adminscript.git
cd 365_Adminscript
```

Exécutez le script principal :
```powershell
.\Scripts\Sherl0ck_v4.1.ps1
```

Les modules requis seront téléchargés automatiquement depuis la galerie PowerShell officielle.

## Structure du projet

```
365_Adminscript/
├── Scripts/
│   └── Sherl0ck_v4.1.ps1          # Point d'entrée principal
├── Modules/
│   ├── Sherl0ck.Utils.psm1        # Utilitaires et helpers
│   ├── Sherl0ck.UI.psm1           # Interface utilisateur et logs
│   ├── Sherl0ck.Auth.psm1         # Authentification Graph/EXO
│   └── Sherl0ck.Audit.psm1        # Collecte et export audit
├── Config/
│   └── config.json.example        # Modèle de configuration externe
├── Tests/
│   ├── Sherl0ck.Tests.Utils.ps1
│   ├── Sherl0ck.Tests.UI.ps1
│   ├── Sherl0ck.Tests.Auth.ps1
│   └── Sherl0ck.Tests.Audit.ps1
├── README.md
└── LICENSE
```

## Tests

Utilise Pester pour valider chaque module indépendamment :
```powershell
Invoke-Pester -Path ./Tests/
```

## Corrections appliquées (Sherl0ck v4.1)

- **H1** : MSAL conflict handling (Graph <> Exchange) — protection contre les collisions de contexte d'authentification ✅
- **H2** : Secure logging — masquage automatique des UPNs sensibles via `Mask-SensitiveData` ✅
- **H3** : Scope OAuth separation — modes ReadOnly/ReadWrite avec paramètre dynamique ✅
- **H4** : Module pinning via PSGallery officielle — vérification des modules requis ⚠️ Partiel
- **M1** : Masquage UPNs (`***`) + chiffrement logs via `ConvertFrom-SecureString` ✅
- **M2** : Vérifier existence avant écriture + incrémenter nom si nécessaire ❌ À faire
- **M3** : Vérifier somme de contrôle / source fiable + Ajouter -SkipModuleInstall ❌ À faire
- **L1** : Refactorisation modulaire en modules PowerShell ❌ À faire
- **L2** : Tests unitaires étendus (Pester) ❌ À faire
- **L3** : Mode simulation -WhatIf ❌ À faire
- **I1** : Externalisation de configuration (Sherl0ck_Config.json) ✅
- **I2** : Support multilingue (fr-FR / en-US) ❌ À faire
- **I3** : Documentation intégrée (commentaires help XML) ✅

## Licence

MIT - Voir LICENSE pour plus d'informations.

## Crédits

Développé par g33ky00 — basé sur l'outil Sherl0ck v4.1.
