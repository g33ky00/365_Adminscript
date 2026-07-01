#Requires -Version 5.1
<#
.SYNOPSIS
    Sherl0ck v4.0 - Administration Unifiée M365 (Entra ID, Exchange, Intune)
    Audit complet avec export Excel multi-onglets (ImportExcel).

.DESCRIPTION
    Collecte exhaustive :
    - Utilisateurs (activité, MFA, licences, sync AD)
    - Administrateurs et rôles Entra ID
    - Appareils (Entra ID + Intune MDM fusionnés)
    - Boîtes Exchange (redirections, archives, litiges)
    - Licences (noms lisibles)
    - Accès Conditionnel
    - Applications/OAuth
    - OneDrive / SharePoint / Teams
    - Domaines vérifiés

.NOTES
    Modules requis : Microsoft.Graph, ExchangeOnlineManagement
    Module optionnel : ImportExcel (installé automatiquement si absent)
    Scopes Graph requis : voir Connect-O365Core
#>

[CmdletBinding()]
param()

Set-StrictMode -Off

# ==============================================================================
# VARIABLES GLOBALES
# ==============================================================================
$GLOBAL:POLICY_STATE        = "enabledForReportingButNotEnforced"
$GLOBAL:TARGET_TENANT       = ""
$GLOBAL:TENANT_NAME         = ""
$GLOBAL:ADMIN_UPN           = ""
$GLOBAL:GRAPH_CONNECTED     = $false
$GLOBAL:EXO_CONNECTED       = $false
$GLOBAL:SESSION_LOGS        = @()
$GLOBAL:BREAKGLASS_UPN      = ""

$ConfigPath                 = Join-Path $PSScriptRoot "Sherl0ck_Config.json"

$GLOBAL:LOG_DIR             = Join-Path $env:LOCALAPPDATA "Sherl0ck_Logs"
if (-not (Test-Path $GLOBAL:LOG_DIR)) { New-Item -ItemType Directory -Force -Path $GLOBAL:LOG_DIR | Out-Null }
$GLOBAL:LOG_FILE            = Join-Path $GLOBAL:LOG_DIR "Session_$(Get-Date -Format 'yyyyMMdd').log"

$GLOBAL:EDGE_TEMP_DIR       = Join-Path $env:TEMP "Sherl0ck_Edge_TempProfile"
$GLOBAL:AUDIT_DIR           = Join-Path $env:USERPROFILE "Documents\Sherl0ck_Audits"
if (-not (Test-Path $GLOBAL:AUDIT_DIR)) { New-Item -ItemType Directory -Force -Path $GLOBAL:AUDIT_DIR | Out-Null }

# ==============================================================================
# TABLE DE CORRESPONDANCE SKU → NOM LISIBLE
# ==============================================================================
$GLOBAL:SKU_NAMES = @{
    "SPE_E3"                       = "Microsoft 365 E3"
    "SPE_E5"                       = "Microsoft 365 E5"
    "SPE_F1"                       = "Microsoft 365 F1"
    "O365_BUSINESS_PREMIUM"        = "Microsoft 365 Business Premium"
    "O365_BUSINESS_ESSENTIALS"     = "Microsoft 365 Business Basic"
    "O365_BUSINESS"                = "Microsoft 365 Apps for Business"
    "STANDARDPACK"                 = "Office 365 E1"
    "STANDARDWOFFPACK"             = "Office 365 E2"
    "ENTERPRISEPACK"               = "Office 365 E3"
    "ENTERPRISEPREMIUM"            = "Office 365 E5"
    "DESKLESSPACK"                 = "Office 365 F3"
    "EXCHANGESTANDARD"             = "Exchange Online Plan 1"
    "EXCHANGEENTERPRISE"           = "Exchange Online Plan 2"
    "EXCHANGEDESKLESS"             = "Exchange Online Kiosk"
    "INTUNE_A"                     = "Intune Plan 1"
    "INTUNE_A_D365"                = "Intune for Education"
    "EMS"                          = "Enterprise Mobility + Security E3"
    "EMSPREMIUM"                   = "Enterprise Mobility + Security E5"
    "AAD_PREMIUM"                  = "Azure AD Premium P1"
    "AAD_PREMIUM_P2"               = "Azure AD Premium P2"
    "POWER_BI_STANDARD"            = "Power BI (gratuit)"
    "POWER_BI_PRO"                 = "Power BI Pro"
    "POWER_BI_PREMIUM_PER_USER"    = "Power BI Premium Per User"
    "PROJECTPREMIUM"               = "Project Plan 5"
    "PROJECTPROFESSIONAL"          = "Project Plan 3"
    "VISIOCLIENT"                  = "Visio Plan 2"
    "VISIO_PLAN1_NAT"              = "Visio Plan 1"
    "STREAM"                       = "Microsoft Stream"
    "TEAMS_EXPLORATORY"            = "Teams Exploratory"
    "MCOEV"                        = "Teams Phone Standard"
    "MCOPSTN1"                     = "Calling Plan Domestique"
    "MCOPSTN2"                     = "Calling Plan International"
    "RMSBASIC"                     = "Azure RMS Basic"
    "RIGHTSMANAGEMENT"             = "Azure Information Protection Plan 1"
    "DEFENDER_ENDPOINT_P1"         = "Defender for Endpoint Plan 1"
    "DEFENDER_ENDPOINT_P2"         = "Defender for Endpoint Plan 2"
    "ATP_ENTERPRISE"               = "Defender for Office 365 Plan 1"
    "THREAT_INTELLIGENCE"          = "Defender for Office 365 Plan 2"
    "CRMSTANDARD"                  = "Dynamics 365 Customer Engagement"
    "WINDOWS_STORE"                = "Windows Store for Business"
    "WIN10_PRO_ENT_SUB"            = "Windows 10/11 Enterprise E3"
    "WIN_DEF_ATP"                  = "Microsoft Defender for Endpoint"
}

# ==============================================================================
# UTILITAIRES
# ==============================================================================

function Convert-EXOSizeToGB {
    param($SizeObj)
    if (-not $SizeObj) { return 0 }
    $str = $SizeObj.ToString()
    if ($str -match "\((?<bytes>\d+)\s*bytes\)") { return [math]::Round([double]$matches['bytes'] / 1GB, 2) }
    if ($str -match "(?<val>[\d\.,]+)\s*GB")      { return [math]::Round([double]($matches['val'] -replace ',', '.'), 2) }
    if ($str -match "(?<val>[\d\.,]+)\s*MB")      { return [math]::Round([double]($matches['val'] -replace ',', '.') / 1024, 2) }
    return 0
}

function Convert-BytesToGB {
    param([long]$Bytes)
    if ($Bytes -le 0) { return 0 }
    return [math]::Round($Bytes / 1GB, 2)
}

function Resolve-SkuName {
    param([string]$SkuPartNumber)
    if ($GLOBAL:SKU_NAMES.ContainsKey($SkuPartNumber)) { return $GLOBAL:SKU_NAMES[$SkuPartNumber] }
    return $SkuPartNumber
}

function Invoke-GraphPagedRequest {
    param(
        [string]$Uri,
        [hashtable]$Headers = @{ ConsistencyLevel = "eventual" }
    )
    $Results = @()
    $Query = $Uri
    while ($Query) {
        $Response = Invoke-MgGraphRequest -Method GET -Uri $Query -Headers $Headers -ErrorAction Stop
        if ($Response.value) { $Results += $Response.value }
        $Query = $Response.'@odata.nextLink'
    }
    return $Results
}

# ==============================================================================
# MOTEUR DE JOURNALISATION
# ==============================================================================

function Add-SessionLog {
    param([string]$Level, [string]$Message, [string]$Details = "")
    $Entry = "[$(Get-Date -Format 'HH:mm:ss')] [$Level] $Message"
    if ($Details) { $Entry += " | $Details" }
    $GLOBAL:SESSION_LOGS += $Entry
    try { Add-Content -Path $GLOBAL:LOG_FILE -Value $Entry -Encoding UTF8 } catch {}
}

function Show-SessionLogs {
    Write-Host "`n--- JOURNAUX DE SESSION ---" -ForegroundColor Cyan
    Write-Host " Dossier : $GLOBAL:LOG_DIR" -ForegroundColor Yellow
    if ($GLOBAL:SESSION_LOGS.Count -eq 0) {
        Write-Host " [INFO] Aucun événement enregistré." -ForegroundColor DarkGray
    } else {
        foreach ($Log in $GLOBAL:SESSION_LOGS) {
            $Color = switch -Regex ($Log) {
                "\[ERREUR\]"    { "Red"   }
                "\[ATTENTION\]" { "Yellow" }
                "\[ACTION\]"    { "Green"  }
                default         { "DarkGray" }
            }
            Write-Host " $Log" -ForegroundColor $Color
        }
    }
    Write-Host "`n[O] Ouvrir le dossier | [0] Retour" -ForegroundColor Cyan
    $Choice = Read-Host "Action"
    if ($Choice -match "^[Oo]$") { Invoke-Item $GLOBAL:LOG_DIR }
}

# ==============================================================================
# CONFIGURATION JSON
# ==============================================================================

function Load-Configuration {
    if (Test-Path $ConfigPath) {
        try {
            $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            $GLOBAL:BREAKGLASS_UPN = $Config.BreakGlassUPN
        } catch { Add-SessionLog "ERREUR" "Config corrompue" $_.Exception.Message }
    }
}

function Save-Configuration {
    @{ BreakGlassUPN = $GLOBAL:BREAKGLASS_UPN } | ConvertTo-Json |
        Set-Content -Path $ConfigPath -Encoding UTF8
    Write-Host " [OK] Configuration sauvegardée." -ForegroundColor Green
    Add-SessionLog "ACTION" "Break-Glass mis à jour ($GLOBAL:BREAKGLASS_UPN)"
}

function Show-MenuBreakGlass {
    Write-Host "`n--- SÉCURITÉ BREAK-GLASS ---" -ForegroundColor Yellow
    Write-Host " Configuration actuelle : $(if($GLOBAL:BREAKGLASS_UPN){$GLOBAL:BREAKGLASS_UPN}else{'AUCUNE'})" -ForegroundColor Cyan
    $NewUPN = Read-Host "UPN du compte de secours (0 pour annuler)"
    if ($NewUPN -eq "0" -or $NewUPN -eq "") { return }
    if ($NewUPN -match "@") {
        $GLOBAL:BREAKGLASS_UPN = $NewUPN
        Save-Configuration
    } else {
        Write-Warning "Format UPN invalide."
        Add-SessionLog "ATTENTION" "Saisie Break-Glass invalide" "UPN=$NewUPN"
    }
}

# ==============================================================================
# AUTHENTIFICATION — GRAPH
# ==============================================================================

# ==============================================================================
# AUTHENTIFICATION — GRAPH
# ==============================================================================

# ==============================================================================
# AUTHENTIFICATION — GRAPH
# ==============================================================================

function Connect-O365Core {
    if ($GLOBAL:GRAPH_CONNECTED) { return }
    if ($GLOBAL:EXO_CONNECTED) {
        Write-Host "`n[ERREUR] Conflit mémoire MSAL (module Exchange actif)." -ForegroundColor Red
        Write-Host " Redémarrez la console pour basculer sur le moteur Identity." -ForegroundColor Yellow
        Read-Host "Entrée pour retourner au menu"
        return
    }

    Write-Host "`n[INIT] Démarrage du moteur Identity (Graph)..." -ForegroundColor Cyan

    # --- VÉRIFICATION ET AUTO-INSTALLATION DU SOCLE GRAPH ---
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host "`n[INSTALLATION] Module 'Microsoft.Graph.Authentication' introuvable sur ce poste." -ForegroundColor Red
        $InstConfirm = Read-Host "Installer le socle Microsoft Graph maintenant (Scope: CurrentUser) ? (O/N)"
        if ($InstConfirm -match "^[Oo]$") {
            Write-Host " -> Installation en cours... (Patientez, cela peut prendre 1 à 2 minutes)" -ForegroundColor Cyan
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            try { 
                Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop 
            } catch { 
                Write-Host "`n[ERREUR] L'installation a échoué. Détails : $($_.Exception.Message)" -ForegroundColor Red
                Add-SessionLog "ERREUR" "Install Graph" "$($_.Exception.Message)"
                Read-Host "Appuyez sur Entrée pour continuer"
                return 
            }
        } else { return }
    }
    
    try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop } catch {}
    # ---------------------------------------------------------------

    try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction SilentlyContinue | Out-Null } catch {}
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

    do { $GLOBAL:ADMIN_UPN = Read-Host "UPN Administrateur Global" } until ($GLOBAL:ADMIN_UPN -match "@")
    $GLOBAL:TARGET_TENANT = ($GLOBAL:ADMIN_UPN -split "@")[1]

    $Scopes = @(
        "Policy.ReadWrite.ConditionalAccess"
        "Policy.Read.All"
        "User.ReadWrite.All"
        "Organization.Read.All"
        "Domain.Read.All"
        "Device.Read.All"
        "Directory.Read.All"
        "Reports.Read.All"
        "Sites.Read.All"
        "Files.Read.All"
        "AuditLog.Read.All"
        "DeviceManagementManagedDevices.Read.All"
        "Application.Read.All"
        "DelegatedPermissionGrant.Read.All"
    )

    Write-Host "`n[AUTH] Ouverture Edge en navigation privée..." -ForegroundColor Yellow
    try { Start-Process msedge.exe -ArgumentList "--inprivate --new-window --user-data-dir=`"$GLOBAL:EDGE_TEMP_DIR`" https://login.microsoft.com/device" -ErrorAction SilentlyContinue } catch {}

    try {
        Connect-MgGraph -Scopes $Scopes -TenantId $GLOBAL:TARGET_TENANT -UseDeviceAuthentication -ContextScope Process -NoWelcome
    } catch {
        Write-Host "`n[ERREUR] Connexion Graph échouée." -ForegroundColor Red
        Write-Host "Détails: $($_.Exception.Message)" -ForegroundColor Red
        Add-SessionLog "ERREUR" "Connexion Graph" $_.Exception.Message
        Read-Host "Appuyez sur Entrée pour continuer"
        return
    }

    $Ctx = Get-MgContext
    if ($Ctx.Account -ne $GLOBAL:ADMIN_UPN) {
        Write-Host "`n[ERREUR] Identité connectée ($($Ctx.Account)) ≠ cible ($GLOBAL:ADMIN_UPN)." -ForegroundColor Red
        Add-SessionLog "ERREUR" "Mismatch identité Graph" "$($Ctx.Account) vs $GLOBAL:ADMIN_UPN"
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Read-Host "Appuyez sur Entrée pour continuer"
        return
    }

    # CORRECTION : Remplacement de la Cmdlet par une requête REST via le module d'Authentication natif
    try {
        $OrgRequest = Invoke-MgGraphRequest -Method GET -Uri "v1.0/organization" -ErrorAction Stop
        $GLOBAL:TENANT_NAME = $OrgRequest.value[0].displayName
    } catch {
        $GLOBAL:TENANT_NAME = $GLOBAL:TARGET_TENANT
    }

    $GLOBAL:GRAPH_CONNECTED = $true
    Add-SessionLog "ACTION" "Connexion Graph réussie ($GLOBAL:ADMIN_UPN)"
    Write-Host "[OK] Connecté : $($GLOBAL:TENANT_NAME)" -ForegroundColor Green
    Start-Sleep -Seconds 1
}

# ==============================================================================
# AUTHENTIFICATION — EXCHANGE
# ==============================================================================

function Connect-O365Exchange {
    if ($GLOBAL:EXO_CONNECTED) { return }

    Write-Host "`n[INIT] Démarrage du moteur Messagerie (Exchange)..." -ForegroundColor Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if (-not $GLOBAL:ADMIN_UPN) {
        do { $GLOBAL:ADMIN_UPN = Read-Host "UPN Administrateur Global" } until ($GLOBAL:ADMIN_UPN -match "@")
        $GLOBAL:TARGET_TENANT = ($GLOBAL:ADMIN_UPN -split "@")[1]
    }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "[INSTALL] Module 'ExchangeOnlineManagement' absent." -ForegroundColor Red
        $Confirm = Read-Host "Installer maintenant ? (O/N)"
        if ($Confirm -match "^[Oo]$") {
            try { Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop }
            catch { Add-SessionLog "ERREUR" "Install EXO" $_.Exception.Message; return }
        } else { return }
    }

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -UserPrincipalName $GLOBAL:ADMIN_UPN -ShowProgress $false -InformationAction SilentlyContinue -ErrorAction Stop
        $GLOBAL:EXO_CONNECTED = $true
        Add-SessionLog "ACTION" "Connexion Exchange réussie ($GLOBAL:ADMIN_UPN)"
        Write-Host "[OK] Module Messagerie prêt." -ForegroundColor Green
    } catch {
        $GLOBAL:EXO_CONNECTED = $false
        Add-SessionLog "ERREUR" "Connexion Exchange" $_.Exception.Message
        Write-Host "`n[ERREUR] Connexion Exchange échouée." -ForegroundColor Red
        if ($_.Exception.Message -match "WithBroker") {
            Write-Host " [DIAGNOSTIC] Conflit MSAL détecté." -ForegroundColor Yellow
            $Restart = Read-Host " Redémarrer la console ? (O/N)"
            if ($Restart -match "^[Oo]$") {
                Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
                Stop-Process -Id $PID -Force
            }
        } else {
            Read-Host " Entrée pour retourner au menu"
        }
    }
}

# ==============================================================================
# COLLECTEURS DE DONNÉES (Get-* — retournent des objets, pas de CSV)
# ==============================================================================

function Get-LicenseData {
    Write-Host "   -> Licences..." -NoNewline -ForegroundColor DarkGray
    try {
        $SubscribedSkus = Invoke-MgGraphRequest -Uri "v1.0/subscribedSkus" -Method GET -ErrorAction Stop
        $Result = foreach ($Sku in $SubscribedSkus.value) {
            [PSCustomObject]@{
                SkuPartNumber = $Sku.skuPartNumber
                NomLisible    = Resolve-SkuName $Sku.skuPartNumber
                Total         = $Sku.prepaidUnits.enabled
                Utilises      = $Sku.consumedUnits
                Restants      = $Sku.prepaidUnits.enabled - $Sku.consumedUnits
                Suspendu      = $Sku.prepaidUnits.suspended
                SkuId         = $Sku.skuId
            }
        }
        Write-Host " [OK] ($($Result.Count))" -ForegroundColor Green
        return $Result
    } catch {
        Add-SessionLog "ERREUR" "Get-LicenseData" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return @()
    }
}

function Get-UserData {
    Write-Host "   -> Utilisateurs..." -NoNewline -ForegroundColor DarkGray
    try {
        # Construction d'un index SKU id→nom pour jointure licences
        $SkuMap = @{}
        try {
            $SkuResp = Invoke-MgGraphRequest -Uri "v1.0/subscribedSkus" -Method GET -ErrorAction Stop
            foreach ($S in $SkuResp.value) { $SkuMap[$S.skuId] = Resolve-SkuName $S.skuPartNumber }
        } catch {}

        $Fields = @(
            "id","displayName","userPrincipalName","mail","accountEnabled",
            "createdDateTime","lastPasswordChangeDateTime",
            "jobTitle","department","officeLocation","mobilePhone",
            "usageLocation","assignedLicenses","userType",
            "onPremisesSyncEnabled","onPremisesLastSyncDateTime",
            "passwordPolicies","proxyAddresses","signInActivity"
        ) -join ","

        $Users = Invoke-GraphPagedRequest -Uri "v1.0/users?`$select=$Fields&`$top=999"

        $Result = foreach ($U in $Users) {
            # Résolution licences
            $LicNames = @()
            if ($U.assignedLicenses) {
                foreach ($Lic in $U.assignedLicenses) {
                    $Name = if ($SkuMap.ContainsKey($Lic.skuId)) { $SkuMap[$Lic.skuId] } else { $Lic.skuId }
                    $LicNames += $Name
                }
            }

            # Activité de connexion (nécessite AAD P1 + AuditLog.Read.All)
            $LastSignIn              = ""
            $LastNonInteractiveSignIn = ""
            if ($U.signInActivity) {
                $LastSignIn               = $U.signInActivity.lastSignInDateTime
                $LastNonInteractiveSignIn = $U.signInActivity.lastNonInteractiveSignInDateTime
            }

            # Calcul inactivité
            $InactiveDays = ""
            if ($LastSignIn) {
                $InactiveDays = [math]::Round(((Get-Date) - [datetime]$LastSignIn).TotalDays, 0)
            }

            [PSCustomObject]@{
                Nom                   = $U.displayName
                UPN                   = $U.userPrincipalName
                Email                 = $U.mail
                Actif                 = $U.accountEnabled
                Type                  = $U.userType           # Member / Guest
                Poste                 = $U.jobTitle
                Departement           = $U.department
                Bureau                = $U.officeLocation
                Mobile                = $U.mobilePhone
                Pays                  = $U.usageLocation
                Licences              = if ($LicNames.Count -gt 0) { $LicNames -join " | " } else { "Aucune" }
                Nb_Licences           = $LicNames.Count
                Derniere_Connexion    = $LastSignIn
                Derniere_Co_Passive   = $LastNonInteractiveSignIn
                Jours_Inactivite      = $InactiveDays
                Sync_OnPrem           = $U.onPremisesSyncEnabled
                Derniere_Sync_AD      = $U.onPremisesLastSyncDateTime
                Mdp_Expire_Jamais     = ($U.passwordPolicies -match "DisablePasswordExpiration")
                Dernier_Changement_Mdp = $U.lastPasswordChangeDateTime
                Alias_SMTP            = if ($U.proxyAddresses) { ($U.proxyAddresses -join " | ") } else { "" }
                Date_Creation         = $U.createdDateTime
            }
        }
        Write-Host " [OK] ($($Result.Count))" -ForegroundColor Green
        return $Result
    } catch {
        Add-SessionLog "ERREUR" "Get-UserData" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return @()
    }
}

function Get-AdminData {
    Write-Host "   -> Administrateurs et rôles..." -NoNewline -ForegroundColor DarkGray
    try {
        $Result = @()
        $Roles = Invoke-MgGraphRequest -Uri "v1.0/directoryRoles" -Method GET -ErrorAction Stop

        foreach ($Role in $Roles.value) {
            try {
                $Members = Invoke-MgGraphRequest -Uri "v1.0/directoryRoles/$($Role.id)/members" -Method GET -ErrorAction SilentlyContinue
                foreach ($M in $Members.value) {
                    $Result += [PSCustomObject]@{
                        Role          = $Role.displayName
                        Description   = $Role.description
                        Membre        = $M.displayName
                        UPN           = $M.userPrincipalName
                        Type          = ($M.'@odata.type' -replace '#microsoft.graph.', '')
                        CompteActif   = $M.accountEnabled
                        BreakGlass    = ($M.userPrincipalName -eq $GLOBAL:BREAKGLASS_UPN)
                        RoleId        = $Role.id
                    }
                }
            } catch {}
        }

        Write-Host " [OK] ($($Result.Count))" -ForegroundColor Green

        # Alerte Global Admin
        $GAs = $Result | Where-Object { $_.Role -eq "Global Administrator" -and -not $_.BreakGlass }
        if ($GAs.Count -gt 2) {
            Write-Host "   [ATTENTION] $($GAs.Count) Global Admins actifs hors Break-Glass !" -ForegroundColor Yellow
            Add-SessionLog "ATTENTION" "Multiples Global Admins" ($GAs.UPN -join ", ")
        }

        return $Result
    } catch {
        Add-SessionLog "ERREUR" "Get-AdminData" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return @()
    }
}

function Get-MFAData {
    Write-Host "   -> Statuts MFA..." -NoNewline -ForegroundColor DarkGray
    try {
        $Raw = Invoke-GraphPagedRequest -Uri "v1.0/reports/authenticationMethods/userRegistrationDetails"
        $Result = foreach ($U in $Raw) {
            [PSCustomObject]@{
                Utilisateur              = $U.userDisplayName
                UPN                      = $U.userPrincipalName
                MFA_Enregistre           = $U.isMfaRegistered
                MFA_Capable              = $U.isMfaCapable
                SSPR_Enregistre          = $U.isSsprRegistered
                Compte_Adminsitratif     = $U.isAdmin
                Methodes                 = if ($U.methodsRegistered) { $U.methodsRegistered -join " | " } else { "Aucune" }
                Methode_Defaut           = $U.defaultMfaMethod
            }
        }
        Write-Host " [OK] ($($Result.Count))" -ForegroundColor Green
        return $Result
    } catch {
        Add-SessionLog "ERREUR" "Get-MFAData" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return @()
    }
}

function Get-DeviceData {
    Write-Host "   -> Appareils (Entra ID + Intune)..." -NoNewline -ForegroundColor DarkGray
    try {
        # --- SOURCE 1 : Entra ID ---
        $EntraSelect = @(
            "displayName","deviceId","operatingSystem","operatingSystemVersion",
            "approximateLastSignInDateTime","isCompliant","trustType",
            "accountEnabled","enrollmentType","mdmAppId","manufacturer","model"
        ) -join ","

        $EntraDevices = Invoke-GraphPagedRequest -Uri "v1.0/devices?`$select=$EntraSelect&`$top=999"

        # --- SOURCE 2 : Intune MDM ---
        $IntuneDevices = @()
        try {
            $IntuneSelect = @(
                "deviceName","id","operatingSystem","osVersion","lastSyncDateTime",
                "complianceState","managementAgent","enrolledDateTime",
                "userPrincipalName","azureADDeviceId","manufacturer","model",
                "serialNumber","totalStorageSpaceInBytes","freeStorageSpaceInBytes",
                "isEncrypted","deviceEnrollmentType"
            ) -join ","
            $IntuneDevices = Invoke-GraphPagedRequest -Uri "v1.0/deviceManagement/managedDevices?`$select=$IntuneSelect&`$top=999" -Headers @{}
        } catch {
            Add-SessionLog "ATTENTION" "Intune MDM inaccessible (scope ou licence manquant)" $_.Exception.Message
        }

        # Index Intune par azureADDeviceId
        $IntuneIndex = @{}
        foreach ($I in $IntuneDevices) {
            if ($I.azureADDeviceId) { $IntuneIndex[$I.azureADDeviceId] = $I }
        }

        $Report = @()

        # Entra ID enrichis avec Intune
        foreach ($D in $EntraDevices) {
            $Intune = $null
            if ($D.deviceId -and $IntuneIndex.ContainsKey($D.deviceId)) { $Intune = $IntuneIndex[$D.deviceId] }

            $Report += [PSCustomObject]@{
                Nom                  = $D.displayName
                DeviceId             = $D.deviceId
                Fabricant            = $D.manufacturer
                Modele               = $D.model
                OS                   = $D.operatingSystem
                Version_OS           = $D.operatingSystemVersion
                Type_Jonction        = $D.trustType        # AzureAd / ServerAd / Workplace
                Conforme_Entra       = $D.isCompliant
                Actif                = $D.accountEnabled
                Derniere_Connexion   = $D.approximateLastSignInDateTime
                UPN_Intune           = if ($Intune) { $Intune.userPrincipalName } else { "" }
                Agent_MDM            = if ($Intune) { $Intune.managementAgent } else { "" }
                Conformite_Intune    = if ($Intune) { $Intune.complianceState } else { "" }
                Chiffre              = if ($Intune) { $Intune.isEncrypted } else { "" }
                Numero_Serie         = if ($Intune) { $Intune.serialNumber } else { "" }
                Stockage_Total_Go    = if ($Intune -and $Intune.totalStorageSpaceInBytes) { Convert-BytesToGB $Intune.totalStorageSpaceInBytes } else { "" }
                Stockage_Libre_Go    = if ($Intune -and $Intune.freeStorageSpaceInBytes)  { Convert-BytesToGB $Intune.freeStorageSpaceInBytes }  else { "" }
                Derniere_Synchro_MDM = if ($Intune) { $Intune.lastSyncDateTime } else { "" }
                Date_Enrolement_MDM  = if ($Intune) { $Intune.enrolledDateTime } else { "" }
                Source               = if ($Intune) { "EntraID+Intune" } else { "EntraID" }
            }
        }

        # Appareils Intune sans objet Entra ID (BYOD sans join)
        foreach ($I in $IntuneDevices) {
            if (-not ($I.azureADDeviceId -and $IntuneIndex.ContainsKey($I.azureADDeviceId))) {
                $Report += [PSCustomObject]@{
                    Nom                  = $I.deviceName
                    DeviceId             = $I.azureADDeviceId
                    Fabricant            = $I.manufacturer
                    Modele               = $I.model
                    OS                   = $I.operatingSystem
                    Version_OS           = $I.osVersion
                    Type_Jonction        = "IntuneOnly"
                    Conforme_Entra       = ""
                    Actif                = ""
                    Derniere_Connexion   = ""
                    UPN_Intune           = $I.userPrincipalName
                    Agent_MDM            = $I.managementAgent
                    Conformite_Intune    = $I.complianceState
                    Chiffre              = $I.isEncrypted
                    Numero_Serie         = $I.serialNumber
                    Stockage_Total_Go    = if ($I.totalStorageSpaceInBytes) { Convert-BytesToGB $I.totalStorageSpaceInBytes } else { "" }
                    Stockage_Libre_Go    = if ($I.freeStorageSpaceInBytes)  { Convert-BytesToGB $I.freeStorageSpaceInBytes }  else { "" }
                    Derniere_Synchro_MDM = $I.lastSyncDateTime
                    Date_Enrolement_MDM  = $I.enrolledDateTime
                    Source               = "IntuneOnly"
                }
            }
        }

        Write-Host " [OK] (Entra: $($EntraDevices.Count) | Intune: $($IntuneDevices.Count) | Rapport: $($Report.Count))" -ForegroundColor Green
        return $Report
    } catch {
        Add-SessionLog "ERREUR" "Get-DeviceData" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return @()
    }
}

function Get-MailboxData {
    # Nécessite EXO connecté
    Write-Host "   -> Boîtes Exchange..." -NoNewline -ForegroundColor DarkGray
    if (-not $GLOBAL:EXO_CONNECTED) { Write-Host " [SKIP] EXO non connecté." -ForegroundColor Yellow; return @() }
    try {
        $Mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties * -ErrorAction Stop |
            Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress,
                RecipientTypeDetails, Database,
                IssueWarningQuota, ProhibitSendQuota, ProhibitSendReceiveQuota,
                ForwardingSmtpAddress, ForwardingAddress, DeliverToMailboxAndForward,
                LitigationHoldEnabled, ArchiveStatus, ArchiveQuota,
                HiddenFromAddressListsEnabled, RetentionPolicy,
                AuditEnabled, MaxSendSize, MaxReceiveSize

        $Result = foreach ($M in $Mailboxes) {
            $Stats = Get-EXOMailboxStatistics -Identity $M.UserPrincipalName -ErrorAction SilentlyContinue

            # Signaux d'alerte
            $AlertRedirection = ($null -ne $M.ForwardingSmtpAddress -and $M.ForwardingSmtpAddress -ne "")

            [PSCustomObject]@{
                Utilisateur              = $M.DisplayName
                UPN                      = $M.UserPrincipalName
                Email_Principal          = $M.PrimarySmtpAddress
                Type_Boite               = $M.RecipientTypeDetails
                Base_Donnees             = $M.Database
                Taille_Go                = if ($Stats) { Convert-EXOSizeToGB $Stats.TotalItemSize } else { "N/A" }
                Nombre_Objets            = if ($Stats) { $Stats.ItemCount } else { "N/A" }
                Quota_Avertissement      = $M.IssueWarningQuota
                Quota_Blocage_Envoi      = $M.ProhibitSendQuota
                Quota_Blocage_Total      = $M.ProhibitSendReceiveQuota
                Redirection_Externe      = $M.ForwardingSmtpAddress    # CRITIQUE : exfiltration potentielle
                Redirection_Interne      = $M.ForwardingAddress
                Livrer_ET_Rediriger      = $M.DeliverToMailboxAndForward
                ALERTE_Redirection       = $AlertRedirection
                Conservation_Legale      = $M.LitigationHoldEnabled
                Archive_Statut           = $M.ArchiveStatus
                Archive_Quota            = $M.ArchiveQuota
                Masque_GAL               = $M.HiddenFromAddressListsEnabled
                Politique_Retention      = $M.RetentionPolicy
                Audit_Actif              = $M.AuditEnabled
            }
        }

        $Redirections = ($Result | Where-Object { $_.ALERTE_Redirection }).Count
        Write-Host " [OK] ($($Result.Count) boîtes)" -ForegroundColor Green
        if ($Redirections -gt 0) {
            Write-Host "   [ATTENTION] $Redirections redirection(s) externe(s) détectée(s) !" -ForegroundColor Red
            Add-SessionLog "ATTENTION" "Redirections externes détectées" "$Redirections boîtes concernées"
        }
        return $Result
    } catch {
        Add-SessionLog "ERREUR" "Get-MailboxData" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return @()
    }
}

function Get-CAPolicyData {
    Write-Host "   -> Politiques d'accès conditionnel..." -NoNewline -ForegroundColor DarkGray
    try {
        $Policies = Invoke-MgGraphRequest -Uri "v1.0/identity/conditionalAccess/policies" -Method GET -ErrorAction Stop
        $Result = foreach ($P in $Policies.value) {
            [PSCustomObject]@{
                Nom              = $P.displayName
                Etat             = $P.state
                Description      = $P.description
                Utilisateurs     = if ($P.conditions.users.includeUsers)          { $P.conditions.users.includeUsers -join " | " } else { "" }
                Groupes          = if ($P.conditions.users.includeGroups)          { $P.conditions.users.includeGroups -join " | " } else { "" }
                Exclusions       = if ($P.conditions.users.excludeUsers)           { $P.conditions.users.excludeUsers -join " | " } else { "" }
                Applications     = if ($P.conditions.applications.includeApplications) { $P.conditions.applications.includeApplications -join " | " } else { "" }
                Plateformes      = if ($P.conditions.platforms.includePlatforms)   { $P.conditions.platforms.includePlatforms -join " | " } else { "" }
                Pays             = if ($P.conditions.locations.includeLocations)   { $P.conditions.locations.includeLocations -join " | " } else { "" }
                Controles        = if ($P.grantControls.builtInControls)           { $P.grantControls.builtInControls -join " | " } else { "" }
                Operateur        = if ($P.grantControls.operator)                  { $P.grantControls.operator } else { "" }
                Cree_Le          = $P.createdDateTime
                Modifie_Le       = $P.modifiedDateTime
                Id               = $P.id
            }
        }
        Write-Host " [OK] ($($Result.Count))" -ForegroundColor Green
        return $Result
    } catch {
        Add-SessionLog "ERREUR" "Get-CAPolicyData" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return @()
    }
}

function Get-AppData {
    Write-Host "   -> Applications et consentements OAuth..." -NoNewline -ForegroundColor DarkGray
    try {
        $Apps = Invoke-GraphPagedRequest -Uri "v1.0/servicePrincipals?`$select=id,displayName,appId,publisherName,replyUrls,accountEnabled,createdDateTime,appRoleAssignmentRequired,servicePrincipalType&`$top=999"

        # Permissions déléguées (consentements)
        $Grants = @()
        try {
            $Grants = Invoke-GraphPagedRequest -Uri "v1.0/oauth2PermissionGrants?`$top=999" -Headers @{}
        } catch {
            Add-SessionLog "ATTENTION" "OAuth grants inaccessibles" $_.Exception.Message
        }

        # Index grants par clientId
        $GrantIndex = @{}
        foreach ($G in $Grants) {
            if (-not $GrantIndex.ContainsKey($G.clientId)) { $GrantIndex[$G.clientId] = @() }
            $GrantIndex[$G.clientId] += $G
        }

        $Result = foreach ($App in $Apps) {
            $AppGrants    = if ($GrantIndex.ContainsKey($App.id)) { $GrantIndex[$App.id] } else { @() }
            $AdminConsent = ($AppGrants | Where-Object { $_.consentType -eq "AllPrincipals" }).Count -gt 0
            $AllScopes    = ($AppGrants | ForEach-Object { $_.scope } | Where-Object { $_ }) -join " "

            # Signaux d'alerte
            $HasDangerousScopes = ($AllScopes -match "Mail\.ReadWrite|Files\.ReadWrite\.All|Directory\.ReadWrite|User\.ReadWrite\.All|offline_access")

            [PSCustomObject]@{
                Application          = $App.displayName
                AppId                = $App.appId
                Editeur              = $App.publisherName
                Type                 = $App.servicePrincipalType
                Active               = $App.accountEnabled
                Consentement_Admin   = $AdminConsent
                Permissions          = $AllScopes
                ALERTE_Scope_Etendu  = $HasDangerousScopes
                Cree_Le              = $App.createdDateTime
            }
        }

        $Risky = ($Result | Where-Object { $_.ALERTE_Scope_Etendu }).Count
        Write-Host " [OK] ($($Apps.Count) apps)" -ForegroundColor Green
        if ($Risky -gt 0) {
            Write-Host "   [ATTENTION] $Risky app(s) avec scopes étendus (Mail/Files/Directory write) !" -ForegroundColor Yellow
            Add-SessionLog "ATTENTION" "Apps OAuth scopes dangereux" "$Risky applications"
        }
        return $Result
    } catch {
        Add-SessionLog "ERREUR" "Get-AppData" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return @()
    }
}

function Get-OneDriveData {
    Write-Host "   -> Volumétrie OneDrive..." -NoNewline -ForegroundColor DarkGray
    try {
        $Users = Invoke-GraphPagedRequest -Uri "v1.0/users?`$select=id,displayName,userPrincipalName&`$top=999"
        $Result = @()
        foreach ($U in $Users) {
            try {
                $Drive = Invoke-MgGraphRequest -Uri "v1.0/users/$($U.id)/drive" -Method GET -ErrorAction SilentlyContinue
                if ($Drive) {
                    $Result += [PSCustomObject]@{
                        Utilisateur      = $U.displayName
                        UPN              = $U.userPrincipalName
                        Utilise_Go       = if ($Drive.quota) { Convert-BytesToGB $Drive.quota.used }      else { "N/A" }
                        Total_Go         = if ($Drive.quota) { Convert-BytesToGB $Drive.quota.total }     else { "N/A" }
                        Restant_Go       = if ($Drive.quota) { Convert-BytesToGB $Drive.quota.remaining } else { "N/A" }
                        Taux_Occupation  = if ($Drive.quota -and $Drive.quota.total -gt 0) {
                                               [math]::Round($Drive.quota.used / $Drive.quota.total * 100, 1)
                                           } else { "N/A" }
                        DriveId          = $Drive.id
                    }
                }
            } catch {}
        }
        Write-Host " [OK] ($($Result.Count))" -ForegroundColor Green
        return $Result
    } catch {
        Add-SessionLog "ERREUR" "Get-OneDriveData" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return @()
    }
}

function Get-SharePointData {
    Write-Host "   -> Volumétrie SharePoint / Teams..." -NoNewline -ForegroundColor DarkGray
    try {
        $Sites = Invoke-GraphPagedRequest -Uri "v1.0/sites?`$top=999"
        $Result = @()
        foreach ($Site in $Sites) {
            try {
                $Drive = Invoke-MgGraphRequest -Uri "v1.0/sites/$($Site.id)/drive" -Method GET -ErrorAction SilentlyContinue
                $Result += [PSCustomObject]@{
                    Nom             = $Site.displayName
                    Type            = if ($Site.siteCollection -and $Site.siteCollection.hostname) { "SharePoint" } else { "Teams" }
                    URL             = $Site.webUrl
                    Utilise_Go      = if ($Drive -and $Drive.quota) { Convert-BytesToGB $Drive.quota.used }  else { "N/A" }
                    Total_Go        = if ($Drive -and $Drive.quota) { Convert-BytesToGB $Drive.quota.total } else { "N/A" }
                    SiteId          = $Site.id
                    Cree_Le         = $Site.createdDateTime
                    Derniere_Modif  = $Site.lastModifiedDateTime
                }
            } catch {}
        }
        Write-Host " [OK] ($($Result.Count))" -ForegroundColor Green
        return $Result
    } catch {
        Add-SessionLog "ERREUR" "Get-SharePointData" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return @()
    }
}

function Get-TeamsData {
    Write-Host "   -> Inventaire Teams..." -NoNewline -ForegroundColor DarkGray
    try {
        $Groups = Invoke-GraphPagedRequest -Uri "v1.0/groups?`$filter=resourceProvisioningOptions/any(s:s eq 'Team')&`$select=id,displayName,description,mail,visibility,createdDateTime,mailEnabled,groupTypes&`$top=999"
        $Result = foreach ($G in $Groups) {
            # Récupération du nombre de membres (appel secondaire, best-effort)
            $MemberCount = ""
            try {
                $MemResp = Invoke-MgGraphRequest -Uri "v1.0/groups/$($G.id)/members/`$count" -Method GET -Headers @{ ConsistencyLevel = "eventual" } -ErrorAction SilentlyContinue
                $MemberCount = $MemResp
            } catch {}

            [PSCustomObject]@{
                Equipe        = $G.displayName
                Description   = $G.description
                Email         = $G.mail
                Visibilite    = $G.visibility
                Nb_Membres    = $MemberCount
                Cree_Le       = $G.createdDateTime
                GroupId       = $G.id
            }
        }
        Write-Host " [OK] ($($Result.Count))" -ForegroundColor Green
        return $Result
    } catch {
        Add-SessionLog "ERREUR" "Get-TeamsData" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return @()
    }
}

function Get-DomainData {
    Write-Host "   -> Domaines vérifiés..." -NoNewline -ForegroundColor DarkGray
    try {
        $Domains = Invoke-MgGraphRequest -Uri "v1.0/domains" -Method GET -ErrorAction Stop
        $Result = foreach ($D in $Domains.value) {
            [PSCustomObject]@{
                Domaine           = $D.id
                Verifie           = $D.isVerified
                Defaut            = $D.isDefault
                Authentification  = $D.authenticationType  # Managed / Federated
                Support_Email     = ($D.supportedServices -contains "Email")
                Support_Teams     = ($D.supportedServices -contains "OfficeCommunicationsOnline")
                Support_SharePoint = ($D.supportedServices -contains "SharePoint")
                Admin             = $D.isAdminManaged
            }
        }
        Write-Host " [OK] ($($Result.Count))" -ForegroundColor Green

        $Federes = ($Result | Where-Object { $_.Authentification -eq "Federated" }).Count
        if ($Federes -gt 0) {
            Write-Host "   [INFO] $Federes domaine(s) fédéré(s) détecté(s)." -ForegroundColor Cyan
        }
        return $Result
    } catch {
        Add-SessionLog "ERREUR" "Get-DomainData" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return @()
    }
}

# ==============================================================================
# EXPORTS CSV INDIVIDUELS (wrappeurs sur Get-*)
# ==============================================================================

function Export-ToCSV {
    param([array]$Data, [string]$FileName, [string]$Label)
    if ($Data.Count -gt 0) {
        $Path = Join-Path $GLOBAL:AUDIT_DIR "$FileName`_$(Get-Date -Format yyyyMMdd_HHmm).csv"
        $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        Write-Host " -> Export CSV : $Path" -ForegroundColor DarkGray
        Add-SessionLog "ACTION" "Export $Label" $Path
        return $Path
    }
    return ""
}

function Export-LicenseInventory        { $D = Get-LicenseData;    Export-ToCSV $D "Audit_Licences"       "Licences" | Out-Null;       $D | Format-Table -AutoSize }
function Export-UserInventory           { $D = Get-UserData;       Export-ToCSV $D "Audit_Utilisateurs"   "Utilisateurs" | Out-Null;   Write-Host "Total: $($D.Count)" -ForegroundColor Yellow }
function Export-AdminInventory          { $D = Get-AdminData;      Export-ToCSV $D "Audit_Admins"         "Admins" | Out-Null;         $D | Format-Table Role,Membre,UPN,CompteActif,BreakGlass -AutoSize }
function Export-MFAStatus               { $D = Get-MFAData;        Export-ToCSV $D "Audit_MFA"            "MFA" | Out-Null;            $En = ($D|Where-Object{$_.MFA_Enregistre}).Count; Write-Host "MFA enregistré: $En/$($D.Count)" -ForegroundColor Yellow }
function Export-DeviceInventory         { $D = Get-DeviceData;     Export-ToCSV $D "Audit_Appareils"      "Appareils" | Out-Null;      Write-Host "Total: $($D.Count)" -ForegroundColor Yellow }
function Export-CAPolicies              { $D = Get-CAPolicyData;   Export-ToCSV $D "Audit_CA_Policies"    "CA Policies" | Out-Null;    $D | Format-Table Nom,Etat,Controles -AutoSize }
function Export-EnterpriseApps          { $D = Get-AppData;        Export-ToCSV $D "Audit_Apps_OAuth"     "Apps OAuth" | Out-Null;     Write-Host "Total: $($D.Count)" -ForegroundColor Yellow }
function Export-OneDriveUsage {
    Write-Host "`n[VOLUMÉTRIE] Analyse complète OneDrive (Tous utilisateurs)..." -ForegroundColor Cyan
    try {
        $Users = @()
        # On récupère tous les utilisateurs
        $UserQuery = "v1.0/users?`$select=id,displayName,userPrincipalName&`$top=999"
        while ($UserQuery) {
            $Response = Invoke-MgGraphRequest -Uri $UserQuery -Method GET
            $Users += $Response.value
            $UserQuery = $Response.'@odata.nextLink'
        }

        $ODStats = @()
        foreach ($U in $Users) {
            # On tente de récupérer le drive
            try {
                $Drive = Invoke-MgGraphRequest -Uri "v1.0/users/$($U.id)/drive" -Method GET -ErrorAction Stop
                $ODStats += [PSCustomObject]@{
                    Utilisateur       = $U.displayName
                    UPN               = $U.userPrincipalName
                    Statut            = "Actif"
                    Espace_utilise_Go = if ($Drive.quota) { Convert-BytesToGB $Drive.quota.used } else { 0 }
                    Espace_total_Go   = if ($Drive.quota) { Convert-BytesToGB $Drive.quota.total } else { 0 }
                }
            } catch {
                # Ici, on attrape le 404 : l'utilisateur existe mais n'a pas de OneDrive
                $ODStats += [PSCustomObject]@{
                    Utilisateur       = $U.displayName
                    UPN               = $U.userPrincipalName
                    Statut            = "Non provisionné (Jamais connecté)"
                    Espace_utilise_Go = 0
                    Espace_total_Go   = 0
                }
            }
        }

        if ($ODStats.Count -gt 0) {
            $Path = Join-Path $GLOBAL:AUDIT_DIR "Sherl0ck_Volumetrie_OneDrive_$(Get-Date -Format yyyyMMdd_HHmm).csv"
            $ODStats | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
            Write-Host " [OK] Rapport complet généré : $Path" -ForegroundColor Green
            Write-Host " (Inclus : $($ODStats.Count) utilisateurs au total)" -ForegroundColor Yellow
        }
    } catch {
        Add-SessionLog "ERREUR" "Volumétrie OneDrive" "$($_.Exception.Message)"
        Write-Host " [ÉCHEC] Erreur critique lors de la boucle : $($_.Exception.Message)" -ForegroundColor Red
    }
}
function Export-SharePointUsage         { $D = Get-SharePointData; Export-ToCSV $D "Audit_SharePoint"     "SharePoint" | Out-Null;     Write-Host "Total: $($D.Count)" -ForegroundColor Yellow }
function Export-TeamsInventory          { $D = Get-TeamsData;      Export-ToCSV $D "Audit_Teams"          "Teams" | Out-Null;          Write-Host "Total: $($D.Count) équipes" -ForegroundColor Yellow }
function Export-DomainInventory         { $D = Get-DomainData;     Export-ToCSV $D "Audit_Domaines"       "Domaines" | Out-Null;       $D | Format-Table -AutoSize }

function Export-MailboxInventory {
    Connect-O365Exchange
    if ($GLOBAL:EXO_CONNECTED) {
        $D = Get-MailboxData
        Export-ToCSV $D "Audit_Boites" "Boîtes Exchange" | Out-Null
        Write-Host "Total: $($D.Count) boîtes" -ForegroundColor Yellow
    }
}

# ==============================================================================
# EXPORT EXCEL MULTI-ONGLETS
# ==============================================================================

function Export-FullAuditExcel {
    Write-Host "`n[AUDIT EXCEL] Démarrage de la collecte complète..." -ForegroundColor Cyan

    # Vérification/installation ImportExcel
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "[INSTALL] Module 'ImportExcel' requis pour l'export Excel." -ForegroundColor Yellow
        $Confirm = Read-Host " Installer maintenant ? (O/N)"
        if ($Confirm -match "^[Oo]$") {
            try {
                Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
                Write-Host " [OK] Module ImportExcel installé." -ForegroundColor Green
            } catch {
                Write-Host " [ECHEC] Installez manuellement : Install-Module ImportExcel" -ForegroundColor Red
                Add-SessionLog "ERREUR" "Install ImportExcel" $_.Exception.Message
                return
            }
        } else {
            Write-Host " Export CSV de secours disponible via Menu 3 > [7]." -ForegroundColor Yellow
            return
        }
    }
    Import-Module ImportExcel -ErrorAction Stop

    $Date    = Get-Date -Format "yyyyMMdd_HHmm"
    $XlsPath = Join-Path $GLOBAL:AUDIT_DIR "Sherl0ck_Audit_$($GLOBAL:TENANT_NAME)_$Date.xlsx"

    Write-Host "`n[COLLECTE]" -ForegroundColor Cyan

    # Collecte de toutes les sources
    $DataLicenses    = Get-LicenseData
    $DataUsers       = Get-UserData
    $DataAdmins      = Get-AdminData
    $DataMFA         = Get-MFAData
    $DataDevices     = Get-DeviceData
    $DataCAP         = Get-CAPolicyData
    $DataApps        = Get-AppData
    $DataOneDrive    = Get-OneDriveData
    $DataSharePoint  = Get-SharePointData
    $DataTeams       = Get-TeamsData
    $DataDomains     = Get-DomainData

    # Données Exchange (optionnel — nécessite connexion EXO)
    $DataMailboxes   = @()
    if ($GLOBAL:EXO_CONNECTED) {
        $DataMailboxes = Get-MailboxData
    } else {
        Write-Host "   -> Boîtes Exchange [SKIP] — Connectez-vous via Menu 2 d'abord." -ForegroundColor Yellow
    }

    # --- Onglet Récapitulatif ---
    $MFARate = if ($DataMFA.Count -gt 0) {
        [math]::Round((($DataMFA | Where-Object { $_.MFA_Enregistre }).Count / $DataMFA.Count) * 100, 1)
    } else { "N/A" }

    $Summary = @(
        [PSCustomObject]@{ Indicateur = "Tenant";                        Valeur = $GLOBAL:TENANT_NAME }
        [PSCustomObject]@{ Indicateur = "Date de l'audit";               Valeur = (Get-Date -Format "dd/MM/yyyy HH:mm") }
        [PSCustomObject]@{ Indicateur = "Auditeur";                      Valeur = $GLOBAL:ADMIN_UPN }
        [PSCustomObject]@{ Indicateur = "Compte Break-Glass";            Valeur = if ($GLOBAL:BREAKGLASS_UPN) { $GLOBAL:BREAKGLASS_UPN } else { "NON DÉFINI" } }
        [PSCustomObject]@{ Indicateur = ""; Valeur = "" }
        [PSCustomObject]@{ Indicateur = "=== UTILISATEURS ===";         Valeur = "" }
        [PSCustomObject]@{ Indicateur = "Total utilisateurs";            Valeur = $DataUsers.Count }
        [PSCustomObject]@{ Indicateur = "Comptes actifs";                Valeur = ($DataUsers | Where-Object { $_.Actif }).Count }
        [PSCustomObject]@{ Indicateur = "Comptes désactivés";            Valeur = ($DataUsers | Where-Object { -not $_.Actif }).Count }
        [PSCustomObject]@{ Indicateur = "Comptes invités (Guest)";       Valeur = ($DataUsers | Where-Object { $_.Type -eq "Guest" }).Count }
        [PSCustomObject]@{ Indicateur = "Comptes sans licence";          Valeur = ($DataUsers | Where-Object { $_.Nb_Licences -eq 0 }).Count }
        [PSCustomObject]@{ Indicateur = "Comptes inactifs (+90j)";       Valeur = ($DataUsers | Where-Object { $_.Jours_Inactivite -ne "" -and [int]$_.Jours_Inactivite -ge 90 }).Count }
        [PSCustomObject]@{ Indicateur = "Mdp n'expire jamais";           Valeur = ($DataUsers | Where-Object { $_.Mdp_Expire_Jamais }).Count }
        [PSCustomObject]@{ Indicateur = "Sync AD hybride (on-prem)";     Valeur = ($DataUsers | Where-Object { $_.Sync_OnPrem }).Count }
        [PSCustomObject]@{ Indicateur = ""; Valeur = "" }
        [PSCustomObject]@{ Indicateur = "=== SÉCURITÉ ===";             Valeur = "" }
        [PSCustomObject]@{ Indicateur = "Taux MFA enregistré (%)";       Valeur = $MFARate }
        [PSCustomObject]@{ Indicateur = "Utilisateurs sans MFA";         Valeur = ($DataMFA | Where-Object { -not $_.MFA_Enregistre }).Count }
        [PSCustomObject]@{ Indicateur = "Comptes admin (rôles actifs)";  Valeur = ($DataAdmins | Select-Object UPN -Unique).Count }
        [PSCustomObject]@{ Indicateur = "Global Admins (hors BG)";       Valeur = ($DataAdmins | Where-Object { $_.Role -eq "Global Administrator" -and -not $_.BreakGlass }).Count }
        [PSCustomObject]@{ Indicateur = "Politiques CA actives";         Valeur = ($DataCAP | Where-Object { $_.Etat -eq "enabled" }).Count }
        [PSCustomObject]@{ Indicateur = "Apps OAuth consentement admin";  Valeur = ($DataApps | Where-Object { $_.Consentement_Admin }).Count }
        [PSCustomObject]@{ Indicateur = "Apps scopes étendus (ALERTE)";  Valeur = ($DataApps | Where-Object { $_.ALERTE_Scope_Etendu }).Count }
        [PSCustomObject]@{ Indicateur = ""; Valeur = "" }
        [PSCustomObject]@{ Indicateur = "=== MESSAGERIE ===";           Valeur = "" }
        [PSCustomObject]@{ Indicateur = "Total boîtes Exchange";          Valeur = $DataMailboxes.Count }
        [PSCustomObject]@{ Indicateur = "Redirections externes (ALERTE)"; Valeur = ($DataMailboxes | Where-Object { $_.ALERTE_Redirection }).Count }
        [PSCustomObject]@{ Indicateur = "Boîtes sous conservation légale"; Valeur = ($DataMailboxes | Where-Object { $_.Conservation_Legale }).Count }
        [PSCustomObject]@{ Indicateur = "Archives activées";              Valeur = ($DataMailboxes | Where-Object { $_.Archive_Statut -eq "Active" }).Count }
        [PSCustomObject]@{ Indicateur = ""; Valeur = "" }
        [PSCustomObject]@{ Indicateur = "=== INFRASTRUCTURE ===";       Valeur = "" }
        [PSCustomObject]@{ Indicateur = "Total appareils (Entra+Intune)"; Valeur = $DataDevices.Count }
        [PSCustomObject]@{ Indicateur = "Appareils non conformes";        Valeur = ($DataDevices | Where-Object { $_.Conforme_Entra -eq $false }).Count }
        [PSCustomObject]@{ Indicateur = "Appareils non chiffrés (Intune)"; Valeur = ($DataDevices | Where-Object { $_.Chiffre -eq $false -and $_.Chiffre -ne "" }).Count }
        [PSCustomObject]@{ Indicateur = "Total équipes Teams";             Valeur = $DataTeams.Count }
        [PSCustomObject]@{ Indicateur = "Total sites SharePoint";          Valeur = $DataSharePoint.Count }
        [PSCustomObject]@{ Indicateur = "Domaines vérifiés";               Valeur = ($DataDomains | Where-Object { $_.Verifie }).Count }
        [PSCustomObject]@{ Indicateur = "Domaines fédérés";                Valeur = ($DataDomains | Where-Object { $_.Authentification -eq "Federated" }).Count }
    )

    Write-Host "`n[GÉNÉRATION EXCEL]" -ForegroundColor Cyan

    $ExcelParams = @{
        Path        = $XlsPath
        AutoSize    = $true
        AutoFilter  = $true
        FreezeTopRow = $true
        BoldTopRow  = $true
    }

    # Onglet Récapitulatif en premier
    $Summary | Export-Excel @ExcelParams -WorksheetName "Récapitulatif" -TableStyle Medium2

    # Définition des onglets dans l'ordre souhaité
    $Sheets = [ordered]@{
        "Licences"          = $DataLicenses
        "Utilisateurs"      = $DataUsers
        "Admins_Roles"      = $DataAdmins
        "MFA"               = $DataMFA
        "Appareils"         = $DataDevices
        "Boites_Exchange"   = $DataMailboxes
        "CA_Policies"       = $DataCAP
        "Apps_OAuth"        = $DataApps
        "OneDrive"          = $DataOneDrive
        "SharePoint_Teams"  = $DataSharePoint
        "Equipes_Teams"     = $DataTeams
        "Domaines"          = $DataDomains
    }

    foreach ($Sheet in $Sheets.GetEnumerator()) {
        if ($Sheet.Value -and $Sheet.Value.Count -gt 0) {
            Write-Host "   -> Onglet '$($Sheet.Key)' ($($Sheet.Value.Count) lignes)..." -NoNewline -ForegroundColor DarkGray

            # Style par onglet
            $Style = switch ($Sheet.Key) {
                "Admins_Roles"     { "Medium6"  }  # Rouge
                "MFA"              { "Medium4"  }  # Bleu
                "Appareils"        { "Medium7"  }  # Vert
                "Boites_Exchange"  { "Medium9"  }  # Orange
                "CA_Policies"      { "Medium11" }  # Violet
                "Apps_OAuth"       { "Medium12" }  # Rouge foncé
                default            { "Medium15" }  # Bleu clair
            }

            $Sheet.Value | Export-Excel @ExcelParams -WorksheetName $Sheet.Key -TableStyle $Style -Append
            Write-Host " [OK]" -ForegroundColor Green
        } else {
            Write-Host "   -> Onglet '$($Sheet.Key)' [VIDE - ignoré]" -ForegroundColor DarkGray
        }
    }

    Write-Host "`n[OK] Rapport Excel généré :" -ForegroundColor Green
    Write-Host " $XlsPath" -ForegroundColor Yellow
    Add-SessionLog "ACTION" "Rapport Excel complet généré" $XlsPath
    Invoke-Item $XlsPath
}

# ==============================================================================
# EXPORT CSV COMPLET (fallback sans ImportExcel)
# ==============================================================================

function Export-FullAuditCSV {
    Write-Host "`n[AUDIT COMPLET CSV] Démarrage..." -ForegroundColor Cyan
    $Date      = Get-Date -Format "yyyyMMdd_HHmm"
    $ReportDir = Join-Path $GLOBAL:AUDIT_DIR "Sherl0ck_Rapport_$Date"
    if (-not (Test-Path $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir | Out-Null }

    Write-Host "`n[COLLECTE]" -ForegroundColor Cyan

    $Sheets = [ordered]@{
        "Audit_Licences"      = (Get-LicenseData)
        "Audit_Utilisateurs"  = (Get-UserData)
        "Audit_Admins"        = (Get-AdminData)
        "Audit_MFA"           = (Get-MFAData)
        "Audit_Appareils"     = (Get-DeviceData)
        "Audit_CA_Policies"   = (Get-CAPolicyData)
        "Audit_Apps_OAuth"    = (Get-AppData)
        "Audit_OneDrive"      = (Get-OneDriveData)
        "Audit_SharePoint"    = (Get-SharePointData)
        "Audit_Teams"         = (Get-TeamsData)
        "Audit_Domaines"      = (Get-DomainData)
    }

    if ($GLOBAL:EXO_CONNECTED) {
        $Sheets["Audit_Boites"] = Get-MailboxData
    }

    foreach ($S in $Sheets.GetEnumerator()) {
        if ($S.Value -and $S.Value.Count -gt 0) {
            $Path = Join-Path $ReportDir "$($S.Key).csv"
            $S.Value | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
            Write-Host "   -> $($S.Key).csv ($($S.Value.Count) lignes)" -ForegroundColor DarkGray
        }
    }

    # Index HTML
    $CsvFiles = Get-ChildItem $ReportDir -Filter "*.csv"
    $HTML = @"
<!DOCTYPE html><html><head><meta charset='UTF-8'>
<title>Sherl0ck Audit - $($GLOBAL:TENANT_NAME)</title>
<style>
  body{font-family:Segoe UI,sans-serif;margin:2rem;background:#f0f2f5}
  h1{color:#0078d4}h2{color:#333;border-bottom:2px solid #0078d4;padding-bottom:.4rem}
  .card{background:#fff;border-radius:8px;padding:1.2rem 1.5rem;margin:1rem 0;box-shadow:0 2px 6px rgba(0,0,0,.08)}
  a{color:#0078d4;text-decoration:none}a:hover{text-decoration:underline}
  .tag{display:inline-block;background:#e3f2fd;color:#0078d4;padding:2px 8px;border-radius:12px;font-size:.85em;margin:3px}
</style></head><body>
<h1>Rapport d'audit Microsoft 365</h1>
<div class='card'>
  <p><b>Tenant :</b> $($GLOBAL:TENANT_NAME) &nbsp;|&nbsp; <b>Auditeur :</b> $($GLOBAL:ADMIN_UPN)</p>
  <p><b>Date :</b> $(Get-Date -Format 'dd/MM/yyyy à HH:mm')</p>
</div>
<div class='card'><h2>Fichiers CSV</h2>
$(foreach ($F in $CsvFiles) { "<div><a href='$($F.Name)'>📄 $($F.Name)</a> <span class='tag'>$((Import-Csv $F.FullName).Count) lignes</span></div>" })
</div></body></html>
"@
    $HTML | Out-File "$ReportDir\index.html" -Encoding UTF8

    Write-Host "`n[OK] Rapport CSV généré : $ReportDir" -ForegroundColor Green
    Add-SessionLog "ACTION" "Rapport CSV complet généré" $ReportDir
    Invoke-Item $ReportDir
}

# ==============================================================================
# MODULE : IDENTITY & SECURITY (ACCÈS CONDITIONNEL + MFA)
# ==============================================================================

function Show-MenuIdentity {
    Connect-O365Core
    if (-not $GLOBAL:GRAPH_CONNECTED) { return }

    # Pré-chargement
    Write-Host "`n[DATA] Cartographie Entra ID..." -NoNewline
    try {
        $Skus = (Invoke-MgGraphRequest -Method GET -Uri "v1.0/subscribedSkus").value
        $EligibleSkuIds = @()
        foreach ($Sku in $Skus) {
            foreach ($Plan in $Sku.servicePlans) {
                if ($Plan.servicePlanName -match "AAD_PREMIUM") { $EligibleSkuIds += $Sku.skuId; break }
            }
        }

        $GlobalUsers = Invoke-GraphPagedRequest -Uri "v1.0/users?`$filter=endswith(userPrincipalName,'@$GLOBAL:TARGET_TENANT')&`$select=id,userPrincipalName,assignedLicenses" |
            ForEach-Object {
                $U = $_
                $Eligible = $false
                if ($U.assignedLicenses) {
                    foreach ($Lic in $U.assignedLicenses) {
                        if ($Lic.skuId -in $EligibleSkuIds) { $Eligible = $true; break }
                    }
                }
                [PSCustomObject]@{ Id = $U.id; UPN = $U.userPrincipalName; Eligible = $Eligible }
            }

        $BreakGlassId      = ($GlobalUsers | Where-Object { $_.UPN -eq $GLOBAL:BREAKGLASS_UPN }).Id
        $ExcludeUsersJson  = if ($BreakGlassId) { "`"$BreakGlassId`"" } else { "" }
        Write-Host " [OK] ($($GlobalUsers.Count) comptes)" -ForegroundColor Green
    } catch {
        Add-SessionLog "ERREUR" "Cartographie Identity" $_.Exception.Message
        Write-Host " [ECHEC]" -ForegroundColor Red
        return
    }

    $Quit = $false
    do {
        Write-Host "`n--- [ IDENTITY & SECURITY ] ---" -ForegroundColor Cyan
        Write-Host " [1] Lister les politiques CA"
        Write-Host " [2] Créer une politique MFA (segmentation terminal)"
        Write-Host " [3] Supprimer une politique"
        Write-Host " [4] Export CSV des politiques CA"
        Write-Host " [5] Export CSV des admins/rôles"
        Write-Host " [6] Export CSV des apps et consentements OAuth"
        Write-Host " [0] Retour"

        $Choice = Read-Host "Sélection"
        switch ($Choice) {
            "0" { $Quit = $true }
            "1" {
                $Policies = Invoke-MgGraphRequest -Uri "v1.0/identity/conditionalAccess/policies" -Method GET
                Write-Host "`n POLITIQUES D'ACCÈS CONDITIONNEL :" -ForegroundColor Cyan
                $i = 1
                foreach ($P in $Policies.value) {
                    $Color = switch ($P.state) { "enabled" { "Green" } "disabled" { "Red" } default { "Yellow" } }
                    Write-Host (" [{0:D2}] [{1}] {2}" -f $i, $P.state.ToUpper(), $P.displayName) -ForegroundColor $Color
                    $i++
                }
            }
            "2" { Invoke-CreateMFAPolicy -GlobalUsers $GlobalUsers -ExcludeUsersJson $ExcludeUsersJson }
            "3" { Invoke-DeleteCAPolicy }
            "4" { Connect-O365Core; if ($GLOBAL:GRAPH_CONNECTED) { Export-CAPolicies } }
            "5" { Connect-O365Core; if ($GLOBAL:GRAPH_CONNECTED) { Export-AdminInventory } }
            "6" { Connect-O365Core; if ($GLOBAL:GRAPH_CONNECTED) { Export-EnterpriseApps } }
        }
    } while (-not $Quit)
}

function Invoke-CreateMFAPolicy {
    param($GlobalUsers, $ExcludeUsersJson)

    Write-Host "`n INVENTAIRE DES COMPTES (vert = éligible AAD P1):" -ForegroundColor Cyan
    for ($i = 0; $i -lt $GlobalUsers.Count; $i++) {
        $Color = if ($GlobalUsers[$i].Eligible) { "Green" } else { "DarkGray" }
        Write-Host (" [{0:D2}] {1}" -f ($i + 1), $GlobalUsers[$i].UPN) -ForegroundColor $Color
    }

    $Selection = Read-Host "`n Cibles : index (ex: 1,3), UPN, 'ALL' ou 0 pour annuler"
    if ($Selection -eq "0" -or $Selection -eq "") { return }

    $Targets = @()
    if ($Selection -eq "ALL") { $Targets = $GlobalUsers | Where-Object { $_.Eligible } }
    elseif ($Selection -match "@") { $Targets = $GlobalUsers | Where-Object { $_.UPN -ieq $Selection.Trim() -and $_.Eligible } }
    elseif ($Selection -match "^\d+(,\d+)*$") {
        foreach ($Idx in ($Selection -split ',')) {
            $tIdx = [int]$Idx.Trim() - 1
            if ($tIdx -ge 0 -and $tIdx -lt $GlobalUsers.Count -and $GlobalUsers[$tIdx].Eligible) { $Targets += $GlobalUsers[$tIdx] }
        }
    }
    if ($Targets.Count -eq 0) { Write-Host " Aucun compte éligible sélectionné." -ForegroundColor Yellow; return }

    Write-Host "`n CIBLAGE PAR TERMINAL :" -ForegroundColor Yellow
    Write-Host " [1] TOUS les terminaux | [2] PC (Windows/macOS) | [3] MOBILE (iOS/Android)"
    $Term = Read-Host "Choix"
    $PlatformsJson = switch ($Term) {
        "2"     { '"includePlatforms": ["windows", "macOS"]' }
        "3"     { '"includePlatforms": ["android", "iOS"]' }
        default { '"includePlatforms": ["all"]' }
    }
    $Suffix = switch ($Term) { "2" { "PC" } "3" { "MOBILE" } default { "ALL" } }

    Write-Host "`n MODÈLES DE DURCISSEMENT :" -ForegroundColor Yellow
    Write-Host " [1] STANDARD (24h) | [2] STRICT (12h) | [3] PARANOÏAQUE (chaque connexion)"
    $Profile = Read-Host "Profil"
    $SessCtrl = switch ($Profile) {
        "2"     { '"signInFrequency": {"value": 12, "type": "hours", "isEnabled": true}' }
        "3"     { '"signInFrequency": {"value": null, "type": null, "frequencyInterval": "everyTime", "isEnabled": true}, "persistentBrowser": {"mode": "never", "isEnabled": true}' }
        default { '"signInFrequency": {"value": 1, "type": "days", "isEnabled": true}' }
    }

    $Comment    = Read-Host "Description / commentaire de la politique"
    $PolicyName = "SEC-MFA-$Suffix-$(Get-Date -Format 'yyyyMMdd-HHmm')"
    $IdsJson    = ($Targets.Id | ForEach-Object { "`"$_`"" }) -join ","

    $Body = @"
{
    "displayName": "$PolicyName",
    "description": "$Comment",
    "state": "$GLOBAL:POLICY_STATE",
    "conditions": {
        "applications": { "includeApplications": ["All"] },
        "users": { "includeUsers": [$IdsJson], "excludeUsers": [$ExcludeUsersJson] },
        "platforms": { $PlatformsJson },
        "clientAppTypes": ["browser", "mobileAppsAndDesktopClients"]
    },
    "grantControls": { "operator": "OR", "builtInControls": ["mfa"] },
    "sessionControls": { $SessCtrl }
}
"@
    try {
        Invoke-MgGraphRequest -Method POST -Uri "v1.0/identity/conditionalAccess/policies" -Body $Body -ContentType "application/json" | Out-Null
        Write-Host "[OK] Politique '$PolicyName' déployée (état: $GLOBAL:POLICY_STATE)." -ForegroundColor Green
        Add-SessionLog "ACTION" "Création MFA CA Policy" "$PolicyName | $Comment"
    } catch {
        Add-SessionLog "ERREUR" "Création CA Policy" $_.Exception.Message
        Write-Host "[ECHEC] $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Invoke-DeleteCAPolicy {
    $Policies = Invoke-MgGraphRequest -Uri "v1.0/identity/conditionalAccess/policies" -Method GET
    $i = 1
    $Map = @{}
    foreach ($P in $Policies.value) {
        Write-Host (" [{0:D2}] [{1}] {2}" -f $i, $P.state, $P.displayName)
        $Map[$i] = $P
        $i++
    }
    $Sel = [int](Read-Host "Index à supprimer (0 annuler)")
    if ($Sel -eq 0 -or -not $Map.ContainsKey($Sel)) { return }
    $Target = $Map[$Sel]
    $Confirm = Read-Host "Confirmer la suppression de '$($Target.displayName)' ? (O/N)"
    if ($Confirm -match "^[Oo]$") {
        try {
            Invoke-MgGraphRequest -Method DELETE -Uri "v1.0/identity/conditionalAccess/policies/$($Target.id)" | Out-Null
            Write-Host "[OK] Politique supprimée." -ForegroundColor Green
            Add-SessionLog "ACTION" "Suppression CA Policy" $Target.displayName
        } catch {
            Add-SessionLog "ERREUR" "Suppression CA Policy" $_.Exception.Message
            Write-Host "[ECHEC]" -ForegroundColor Red
        }
    }
}

# ==============================================================================
# MODULE : EXCHANGE ONLINE
# ==============================================================================

function Show-MenuExchange {
    Connect-O365Exchange
    if (-not $GLOBAL:EXO_CONNECTED) { return }

    $Quit = $false
    do {
        Write-Host "`n--- [ EXCHANGE ONLINE ] ---" -ForegroundColor Cyan
        Write-Host " [1] Inventaire des boîtes (complet avec alertes)"
        Write-Host " [2] Volumétrie et statistiques"
        Write-Host " [3] États de connexion actifs"
        Write-Host " [0] Retour"

        $Choice = Read-Host "Sélection"
        switch ($Choice) {
            "0" { $Quit = $true }
            "1" { Export-MailboxInventory }
            "2" {
                Write-Host "`n[VOLUMÉTRIE]" -ForegroundColor Cyan
                try {
                    $Mbx   = Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop
                    $Stats = foreach ($M in $Mbx) {
                        $S = Get-EXOMailboxStatistics -Identity $M.Identity -ErrorAction SilentlyContinue
                        [PSCustomObject]@{
                            Utilisateur = $M.DisplayName
                            UPN         = $M.UserPrincipalName
                            Type        = $M.RecipientTypeDetails
                            Taille_Go   = if ($S) { Convert-EXOSizeToGB $S.TotalItemSize } else { "N/A" }
                            Objets      = if ($S) { $S.ItemCount } else { "N/A" }
                            DB          = $M.Database
                        }
                    }
                    $Path = Join-Path $GLOBAL:AUDIT_DIR "Audit_Volumetrie_EXO_$(Get-Date -Format yyyyMMdd_HHmm).csv"
                    $Stats | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
                    Write-Host " [OK] -> $Path" -ForegroundColor Green
                    $TotalGB = ($Stats | Where-Object { $_.Taille_Go -ne "N/A" } | Measure-Object -Property Taille_Go -Sum).Sum
                    Write-Host " Volume total: $([math]::Round($TotalGB,2)) Go" -ForegroundColor Yellow
                } catch { Write-Host " [ECHEC]" -ForegroundColor Red }
            }
            "3" {
                Write-Host "`n[CONNEXIONS ACTIVES]" -ForegroundColor Cyan
                try {
                    Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop |
                        Get-EXOMailboxStatistics -ErrorAction SilentlyContinue |
                        Where-Object { $_.LastLoggedOnUserCount -gt 0 } |
                        Select-Object DisplayName, LastLoggedOnUserAccount, LastLogonTime |
                        Format-Table -AutoSize
                } catch { Write-Host " [ECHEC]" -ForegroundColor Red }
            }
        }
    } while (-not $Quit)
}

# ==============================================================================
# MODULE : AUDIT 365
# ==============================================================================

function Show-MenuAudit {
    Connect-O365Core
    if (-not $GLOBAL:GRAPH_CONNECTED) { return }

    $Quit = $false
    do {
        Write-Host "`n--- [ AUDIT & VOLUMÉTRIE 365 ] ---" -ForegroundColor Cyan
        Write-Host " [1]  Licences (SKU avec noms lisibles)"
        Write-Host " [2]  Utilisateurs complets (activité, MFA, licences, AD sync)"
        Write-Host " [3]  Administrateurs et rôles Entra ID"
        Write-Host " [4]  MFA global (méthodes et statuts)"
        Write-Host " [5]  Appareils (Entra ID + Intune fusionnés)"
        Write-Host " [6]  Accès Conditionnel (toutes les politiques)"
        Write-Host " [7]  Applications et consentements OAuth"
        Write-Host " [8]  OneDrive (volumétrie par utilisateur)"
        Write-Host " [9]  SharePoint / Teams (sites)"
        Write-Host " [10] Équipes Teams (inventaire groupes)"
        Write-Host " [11] Domaines vérifiés"
        Write-Host ""
        Write-Host " [E]  Export EXCEL multi-onglets (recommandé)" -ForegroundColor Green
        Write-Host " [C]  Export CSV complet (tous les modules)" -ForegroundColor Yellow
        Write-Host " [0]  Retour"

        $Choice = Read-Host "Sélection"
        switch ($Choice) {
            "1"  { Export-LicenseInventory }
            "2"  { Export-UserInventory }
            "3"  { Export-AdminInventory }
            "4"  { Export-MFAStatus }
            "5"  { Export-DeviceInventory }
            "6"  { Export-CAPolicies }
            "7"  { Export-EnterpriseApps }
            "8"  { Export-OneDriveUsage }
            "9"  { Export-SharePointUsage }
            "10" { Export-TeamsInventory }
            "11" { Export-DomainInventory }
            {$_ -match "^[Ee]$"} { Export-FullAuditExcel }
            {$_ -match "^[Cc]$"} { Export-FullAuditCSV }
            "0"  { $Quit = $true }
        }
    } while (-not $Quit)
}

# ==============================================================================
# POINT D'ENTRÉE PRINCIPAL
# ==============================================================================

Load-Configuration
$GlobalQuit = $false

do {
    Clear-Host
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "              SHERL0CK V4.0 — AUDIT M365 UNIFIÉ                " -ForegroundColor White
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " Tenant      : $(if($GLOBAL:TENANT_NAME){$GLOBAL:TENANT_NAME}else{'En attente de connexion...'})" -ForegroundColor Yellow
    Write-Host " Break-Glass : $(if($GLOBAL:BREAKGLASS_UPN){$GLOBAL:BREAKGLASS_UPN}else{'NON DÉFINI'})" -ForegroundColor $(if($GLOBAL:BREAKGLASS_UPN){"Green"}else{"Red"})
    Write-Host " Graph       : $(if($GLOBAL:GRAPH_CONNECTED){'Connecté'}else{'Déconnecté'})" -ForegroundColor $(if($GLOBAL:GRAPH_CONNECTED){"Green"}else{"DarkGray"})
    Write-Host " Exchange    : $(if($GLOBAL:EXO_CONNECTED){'Connecté'}else{'Déconnecté'})" -ForegroundColor $(if($GLOBAL:EXO_CONNECTED){"Green"}else{"DarkGray"})
    Write-Host " Rapports    : $GLOBAL:AUDIT_DIR" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host " [ 1 ] IDENTITY & SECURITY   (MFA, Accès Conditionnel, Apps OAuth)" -ForegroundColor White
    Write-Host " [ 2 ] EXCHANGE ONLINE       (Boîtes, Quotas, Redirections)" -ForegroundColor White
    Write-Host " [ 3 ] AUDIT 365             (Collecte complète + Export Excel/CSV)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " [ 4 ] BREAK-GLASS           (Compte de secours)" -ForegroundColor White
    Write-Host " [ 5 ] JOURNAUX DE SESSION   (Logs et erreurs)" -ForegroundColor White
    Write-Host ""
    Write-Host " [ 0 ] DÉCONNEXION & QUITTER" -ForegroundColor Red

    $MenuChoice = Read-Host "`n [>] Module"
    switch ($MenuChoice) {
        "1" { Show-MenuIdentity }
        "2" { Show-MenuExchange }
        "3" { Show-MenuAudit }
        "4" { Show-MenuBreakGlass }
        "5" { Show-SessionLogs }
        "0" { $GlobalQuit = $true }
    }
} while (-not $GlobalQuit)

# Déconnexion propre
try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
if ($GLOBAL:EXO_CONNECTED) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {} }
try { if (Test-Path $GLOBAL:EDGE_TEMP_DIR) { Remove-Item $GLOBAL:EDGE_TEMP_DIR -Recurse -Force -ErrorAction SilentlyContinue } } catch {}

Write-Host "`n[FIN] Fermeture sécurisée." -ForegroundColor Green
Start-Sleep -Seconds 1
Clear-Host
