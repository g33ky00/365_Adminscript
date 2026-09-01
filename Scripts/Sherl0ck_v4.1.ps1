#Requires -Version 5.1
<#
.SYNOPSIS
    Sherl0ck v4.1 - Administration Unifiee M365 (Entra ID, Exchange, Intune)
    Audit complet, gestion du Throttling Graph, Export Excel et protection des Handles.
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
# UTILITAIRES & PARSERS
# ==============================================================================
function Convert-EXOSizeToGB {
    param($SizeObj)
    if (-not $SizeObj) { return 0 }
    $str = $SizeObj.ToString()
    if ($str -match "\((?<bytes>\d+)\s*bytes\)") { return [math]::Round([double]$matches['bytes'] / 1GB, 2) }
    if ($str -match "(?<val>[\d\.,]+)\s*GB") { return [math]::Round([double]($matches['val'] -replace ',', '.'), 2) }
    if ($str -match "(?<val>[\d\.,]+)\s*MB") { return [math]::Round([double]($matches['val'] -replace ',', '.') / 1024, 2) }
    return 0
}

function Convert-BytesToGB {
    param([long]$Bytes)
    if ($Bytes -le 0) { return 0 }
    return [math]::Round($Bytes / 1GB, 2)
}

function Invoke-SafeOpen {
    param([string]$FilePath)
    if (Test-Path $FilePath) {
        Start-Sleep -Seconds 2
        Invoke-Item $FilePath
    }
}

# ==============================================================================
# MOTEUR DE JOURNALISATION
# ==============================================================================
function Add-SessionLog {
    param([string]$Level, [string]$Message, [string]$Details = "")
    $LogEntry = "[$(Get-Date -Format 'HH:mm:ss')] [$Level] $Message"
    if ($Details) { $LogEntry += " | DETAILS: $Details" }
    $GLOBAL:SESSION_LOGS += $LogEntry
    try { Add-Content -Path $GLOBAL:LOG_FILE -Value $LogEntry -Encoding UTF8BOM } catch {}
}

function Show-SessionLogs {
    Write-Host "`n--- JOURNAUX DE SESSION (AUDIT et ERREURS) ---" -ForegroundColor Cyan
    Write-Host " Dossier : $GLOBAL:LOG_DIR" -ForegroundColor Yellow
    if ($GLOBAL:SESSION_LOGS.Count -eq 0) { Write-Host " [INFO] Aucun evenement." -ForegroundColor DarkGray } 
    else { foreach ($Log in $GLOBAL:SESSION_LOGS) { Write-Host " $Log" -ForegroundColor $(if($Log -match "\[ERREUR\]"){"Red"}elseif($Log -match "\[ATTENTION\]"){"Yellow"}else{"DarkGray"}) } }
    $Choice = Read-Host "`n[O] Ouvrir le dossier | [0] Retour"
    if ($Choice -match "^[Oo]$") { Invoke-SafeOpen -FilePath $GLOBAL:LOG_DIR }
}

# ==============================================================================
# AUTHENTIFICATION — GRAPH & EXCHANGE
# ==============================================================================
function Connect-O365Core {
    if ($GLOBAL:GRAPH_CONNECTED) { return }
    if ($GLOBAL:EXO_CONNECTED) {
        Write-Host "`n[ERREUR] Conflit memoire MSAL (module Exchange actif). Redemarrez la console." -ForegroundColor Red
        Read-Host "Entree pour retourner au menu"; return
    }

    Write-Host "`n[INIT] Demarrage du moteur Identity (Graph)..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host "`n[INSTALLATION] Module 'Microsoft.Graph.Authentication' introuvable." -ForegroundColor Red
        if ((Read-Host "Installer le socle Microsoft Graph maintenant ? (O/N)") -match "^[Oo]$") {
            Write-Host " -> Installation en cours..." -ForegroundColor Cyan
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            try { Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop } 
            catch { Write-Host "[ERREUR] $($_.Exception.Message)" -ForegroundColor Red; return }
        } else { return }
    }
    
    try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop } catch {}
    try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction SilentlyContinue | Out-Null } catch {}
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

    do { $GLOBAL:ADMIN_UPN = Read-Host "UPN Administrateur Global" } until ($GLOBAL:ADMIN_UPN -match "@")
    $GLOBAL:TARGET_TENANT = ($GLOBAL:ADMIN_UPN -split "@")[1]

    $Scopes = @("Policy.ReadWrite.ConditionalAccess", "Policy.Read.All", "User.ReadWrite.All", "Organization.Read.All", "Domain.Read.All", "Device.Read.All", "Directory.Read.All", "Reports.Read.All", "Sites.Read.All", "Files.Read.All", "AuditLog.Read.All", "RoleManagement.Read.Directory", "Application.Read.All")

    Write-Host "`n[AUTH] Ouverture Edge en navigation privee..." -ForegroundColor Yellow
    try { Start-Process msedge.exe -ArgumentList "--inprivate --new-window --user-data-dir=`"$GLOBAL:EDGE_TEMP_DIR`" https://login.microsoft.com/device" -ErrorAction SilentlyContinue } catch {}

    try { Connect-MgGraph -Scopes $Scopes -TenantId $GLOBAL:TARGET_TENANT -UseDeviceAuthentication -ContextScope Process -NoWelcome } 
    catch { Write-Host "`n[ERREUR] Connexion Graph echouee: $($_.Exception.Message)" -ForegroundColor Red; Read-Host "Entree pour continuer"; return }

    $Ctx = Get-MgContext
    if ($Ctx.Account -ne $GLOBAL:ADMIN_UPN) {
        Write-Host "`n[ERREUR] Identite connectee ($($Ctx.Account)) != cible ($GLOBAL:ADMIN_UPN)." -ForegroundColor Red
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null; Read-Host "Entree pour continuer"; return
    }

    try { $GLOBAL:TENANT_NAME = (Invoke-MgGraphRequest -Method GET -Uri "v1.0/organization" -ErrorAction Stop).value[0].displayName } 
    catch { $GLOBAL:TENANT_NAME = $GLOBAL:TARGET_TENANT }

    $GLOBAL:GRAPH_CONNECTED = $true
    Add-SessionLog "ACTION" "Connexion Graph reussie ($GLOBAL:ADMIN_UPN)"
    Write-Host "[OK] Connecte : $($GLOBAL:TENANT_NAME)" -ForegroundColor Green
    Start-Sleep -Seconds 1
}

function Connect-O365Exchange {
    if ($GLOBAL:EXO_CONNECTED) { return }
    Write-Host "`n[INIT] Demarrage du moteur Messagerie (Exchange)..." -ForegroundColor Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if (-not $GLOBAL:ADMIN_UPN) {
        do { $GLOBAL:ADMIN_UPN = Read-Host "UPN Administrateur Global" } until ($GLOBAL:ADMIN_UPN -match "@")
        $GLOBAL:TARGET_TENANT = ($GLOBAL:ADMIN_UPN -split "@")[1]
    }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        if ((Read-Host "Module EXO introuvable. Installer ? (O/N)") -match "^[Oo]$") {
            try { Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop } 
            catch { Write-Host "[ERREUR] $($_.Exception.Message)" -ForegroundColor Red; return }
        } else { return }
    }

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -UserPrincipalName $GLOBAL:ADMIN_UPN -ShowProgress $false -InformationAction SilentlyContinue
        $GLOBAL:EXO_CONNECTED = $true
        Write-Host "[OK] Module Messagerie pret." -ForegroundColor Green
    } catch {
        $GLOBAL:EXO_CONNECTED = $false
        Write-Host "`n[ERREUR] Connexion Exchange: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Entree pour continuer"
    }
}

# ==============================================================================
# MODULE AUDIT 365 (FONCTIONS D'EXPORT MINEURES RESUMEES POUR EXCEL)
# ==============================================================================
function Get-GraphData {
    param([string]$Uri)
    $Data = @(); $Query = $Uri
    while ($Query) {
        try {
            $Resp = Invoke-MgGraphRequest -Uri $Query -Method GET -ErrorAction Stop
            if ($Resp.value) { $Data += $Resp.value }
            $Query = $Resp.'@odata.nextLink'
        } catch { Add-SessionLog "ERREUR" "Graph Query ($Uri)" "$($_.Exception.Message)"; break }
    }
    return $Data
}

function Export-OneDriveUsage {
    Write-Host "`n[VOLUMETRIE] Analyse complete OneDrive (Tous utilisateurs)..." -ForegroundColor Cyan
    try {
        $Users = Get-GraphData "v1.0/users?`$select=id,displayName,userPrincipalName&`$top=999"
        $ODStats = @(); $Count = 0; $MaxRetries = 3
        
        foreach ($U in $Users) {
            $Count++; Write-Host "   -> Traitement $Count/$($Users.Count) : $($U.displayName)..." -NoNewline -ForegroundColor DarkGray
            $RetryCount = 0; $Success = $false
            
            while (-not $Success -and $RetryCount -le $MaxRetries) {
                try {
                    $Drive = Invoke-MgGraphRequest -Uri "v1.0/users/$($U.id)/drive" -Method GET -ErrorAction Stop
                    $ODStats += [PSCustomObject]@{ Utilisateur = $U.displayName; UPN = $U.userPrincipalName; Statut = "Actif"; Espace_utilise_Go = if($Drive.quota){Convert-BytesToGB $Drive.quota.used}else{0}; Espace_total_Go = if($Drive.quota){Convert-BytesToGB $Drive.quota.total}else{0} }
                    Write-Host " [OK]" -ForegroundColor Green; $Success = $true
                } catch {
                    if ($_.Exception.Message -match "429|Too Many Requests") {
                        $RetryCount++; Start-Sleep -Seconds 5
                    } elseif ($_.Exception.Message -match "404|Not Found") {
                        $ODStats += [PSCustomObject]@{ Utilisateur = $U.displayName; UPN = $U.userPrincipalName; Statut = "Non provisionne (Jamais connecte)"; Espace_utilise_Go = 0; Espace_total_Go = 0 }
                        Write-Host " [N/A]" -ForegroundColor Yellow; $Success = $true
                    } else {
                        Write-Host " [ERREUR]" -ForegroundColor Red; $Success = $true
                    }
                }
            }
        }
        $Path = Join-Path $GLOBAL:AUDIT_DIR "Sherl0ck_Volumetrie_OneDrive_$(Get-Date -Format yyyyMMdd_HHmm).csv"
        $ODStats | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8BOM
        return $ODStats
    } catch { Write-Host " [ECHEC] $($_.Exception.Message)" -ForegroundColor Red }
}

function Export-FullAuditExcel {
    Write-Host "`n[AUDIT EXCEL] Generation du classeur global en cours..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host " Installation du module ImportExcel requise pour cette action." -ForegroundColor Yellow
        try { Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop } 
        catch { Write-Host "[ERREUR] $($_.Exception.Message)" -ForegroundColor Red; return }
        Import-Module ImportExcel
    }

    $DateStr = Get-Date -Format "yyyyMMdd_HHmm"
    $XlsxPath = Join-Path $GLOBAL:AUDIT_DIR "Sherl0ck_AuditGlobal_$($GLOBAL:TENANT_NAME)_$DateStr.xlsx"

    try {
        Write-Host " - Collecte Utilisateurs..." -ForegroundColor DarkGray
        $Users = Get-GraphData "v1.0/users?`$select=id,displayName,userPrincipalName,mail,usageLocation,assignedLicenses,userType,accountEnabled"
        $Users | Select-Object displayName, userPrincipalName, accountEnabled, userType | Export-Excel -Path $XlsxPath -WorksheetName "Utilisateurs" -AutoSize -AutoFilter

        Write-Host " - Collecte Licences..." -ForegroundColor DarkGray
        $Skus = Get-GraphData "v1.0/subscribedSkus"
        $Skus | Select-Object skuPartNumber, @{n='Total';e={$_.prepaidUnits.enabled}}, @{n='Consomme';e={$_.consumedUnits}} | Export-Excel -Path $XlsxPath -WorksheetName "Licences" -AutoSize -AutoFilter -Append

        Write-Host " - Collecte Statuts MFA..." -ForegroundColor DarkGray
        $MFA = Get-GraphData "v1.0/reports/authenticationMethods/userRegistrationDetails"
        $MFA | Select-Object userDisplayName, userPrincipalName, isMfaRegistered, @{n='Methodes';e={$_.methodsRegistered -join ', '}} | Export-Excel -Path $XlsxPath -WorksheetName "MFA" -AutoSize -AutoFilter -Append

        Write-Host " - Collecte Appareils..." -ForegroundColor DarkGray
        $Devices = Get-GraphData "v1.0/devices"
        if ($Devices) { $Devices | Select-Object displayName, operatingSystem, isCompliant, trustType, approximateLastSignInDateTime | Export-Excel -Path $XlsxPath -WorksheetName "Appareils" -AutoSize -AutoFilter -Append }

        Write-Host " - Collecte Domaines..." -ForegroundColor DarkGray
        $Domains = Get-GraphData "v1.0/domains"
        $Domains | Select-Object id, isVerified, isDefault | Export-Excel -Path $XlsxPath -WorksheetName "Domaines" -AutoSize -AutoFilter -Append

        Write-Host " - Collecte OneDrive..." -ForegroundColor DarkGray
        $OD = Export-OneDriveUsage
        if ($OD) { $OD | Export-Excel -Path $XlsxPath -WorksheetName "OneDrive" -AutoSize -AutoFilter -Append }

        Write-Host "`n[SUCCES] Audit Excel termine : $XlsxPath" -ForegroundColor Green
        Add-SessionLog "ACTION" "Export Excel reussi" $XlsxPath
        Invoke-SafeOpen -FilePath $XlsxPath
    } catch {
        Write-Host "`n[ERREUR] Generation Excel : $($_.Exception.Message)" -ForegroundColor Red
        Add-SessionLog "ERREUR" "Export Excel" $_.Exception.Message
    }
}

function Show-MenuAudit {
    Connect-O365Core
    if (-not $GLOBAL:GRAPH_CONNECTED) { return }

    $QuitAudit = $false
    do {
        Write-Host "`n--- [ AUDIT 365 ] ---" -ForegroundColor Cyan
        Write-Host " [E]  Export EXCEL multi-onglets (Recommande : Utilisateurs, MFA, Appareils, Licences, OneDrive)"
        Write-Host " [0]  Retour"

        $Choice = Read-Host "Selection"
        switch ($Choice) {
            "E" { Export-FullAuditExcel }
            "e" { Export-FullAuditExcel }
            "0" { $QuitAudit = $true }
        }
    } while (-not $QuitAudit)
}

# ==============================================================================
# POINT D'ENTREE PRINCIPAL
# ==============================================================================
$GlobalQuit = $false

do {
    Clear-Host
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host "                  SHERL0CK V4.1 - META CONSOLE                 " -ForegroundColor White
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " Tenant cible : $(if($GLOBAL:TENANT_NAME){$GLOBAL:TENANT_NAME}else{'En attente de connexion...'})" -ForegroundColor Yellow
    Write-Host " Rapports     : $GLOBAL:AUDIT_DIR" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host " [ 1 ] IDENTITY ET SECURITY   (MFA, Acces Conditionnel, Apps OAuth)" -ForegroundColor White
    Write-Host " [ 2 ] EXCHANGE ONLINE        (Boites, Quotas, Redirections)" -ForegroundColor White
    Write-Host " [ 3 ] AUDIT 365              (Collecte complete + Export Excel)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " [ 4 ] JOURNAUX DE SESSION    (Logs et erreurs)" -ForegroundColor White
    Write-Host ""
    Write-Host " [ 0 ] DECONNEXION ET QUITTER" -ForegroundColor Red

    $MenuChoice = Read-Host "`n [>] Module"
    switch ($MenuChoice) {
        "1" { Connect-O365Core }
        "2" { Connect-O365Exchange }
        "3" { Show-MenuAudit }
        "4" { Show-SessionLogs }
        "0" { $GlobalQuit = $true }
    }
} while (-not $GlobalQuit)

try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
if ($GLOBAL:EXO_CONNECTED) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {} }
try { if (Test-Path $GLOBAL:EDGE_TEMP_DIR) { Remove-Item -Path $GLOBAL:EDGE_TEMP_DIR -Recurse -Force -ErrorAction SilentlyContinue } } catch {}
Write-Host "`n[FIN] Fermeture securisee." -ForegroundColor Green
Start-Sleep -Seconds 1
Clear-Host