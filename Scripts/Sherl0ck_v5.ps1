#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('ReadOnly','ReadWrite')]
    [string]$AuditMode = 'ReadOnly',
    [switch]$SkipModuleInstall
)

Set-StrictMode -Version 3.0

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
$GLOBAL:LOG_KEY_DIR         = Join-Path $env:LOCALAPPDATA "Sherl0ck_Secrets"
if (-not (Test-Path $GLOBAL:LOG_KEY_DIR)) { New-Item -ItemType Directory -Force -Path $GLOBAL:LOG_KEY_DIR | Out-Null }
$GLOBAL:LOG_FILE            = Join-Path $GLOBAL:LOG_DIR "Session_$(Get-Date -Format 'yyyyMMdd').log"
$GLOBAL:EDGE_TEMP_DIR       = Join-Path $env:TEMP "Sherl0ck_Edge_TempProfile"
$GLOBAL:AUDIT_DIR           = Join-Path $env:USERPROFILE "Documents\Sherl0ck_Audits"
if (-not (Test-Path $GLOBAL:AUDIT_DIR)) { New-Item -ItemType Directory -Force -Path $GLOBAL:AUDIT_DIR | Out-Null }

$ModulesDir = Join-Path $PSScriptRoot "..\Modules"
foreach ($moduleName in @('Sherl0ck.Utils', 'Sherl0ck.UI', 'Sherl0ck.Auth', 'Sherl0ck.Audit', 'Sherl0ck.Exchange')) {
    $modulePath = Join-Path $ModulesDir "$moduleName.psm1"
    if (Test-Path $modulePath) {
        Import-Module $modulePath -Force -ErrorAction Stop
        Write-Host "[INIT] Loaded module: $moduleName" -ForegroundColor DarkGray
    } else {
        Write-Host "[ERROR] Module not found: $moduleName at $modulePath" -ForegroundColor Red
    }
}

# Config / Break-Glass
function Load-Configuration {
    if (Test-Path $ConfigPath) {
        try {
            $Config = Get-Content $ConfigPath | ConvertFrom-Json
            $GLOBAL:BREAKGLASS_UPN = $Config.BreakGlassUPN
        } catch { if (Get-Command Add-SessionLog -ErrorAction SilentlyContinue) { Add-SessionLog "ERREUR" "Fichier Config corrompu" $_.Exception.Message } }
    }
}
function Save-Configuration {
    $Config = @{ BreakGlassUPN = $GLOBAL:BREAKGLASS_UPN }
    try {
        $Config | ConvertTo-Json | Set-Content -Path $ConfigPath -Encoding UTF8
        Write-Host " [SUCCES] Configuration mise a jour." -ForegroundColor Green
    } catch { if (Get-Command Add-SessionLog -ErrorAction SilentlyContinue) { Add-SessionLog "ERREUR" "Echec sauvegarde config" $_.Exception.Message } }
}
function Show-MenuBreakGlass {
    Write-Host "`n=== CONCEPT BREAK-GLASS ===" -ForegroundColor Yellow
    Write-Host " Compte de secours exclu de toutes les politiques MFA/CA." -ForegroundColor White
    Write-Host " Configuration actuelle : $(if($GLOBAL:BREAKGLASS_UPN){$GLOBAL:BREAKGLASS_UPN}else{'AUCUNE'})" -ForegroundColor Cyan
    $NewUPN = Read-Host " -> UPN du compte de secours (ou 0 pour annuler)"
    if ($NewUPN -eq "0" -or $NewUPN -eq "") { return }
    if ($NewUPN -match "@") {
        $GLOBAL:BREAKGLASS_UPN = $NewUPN
        Save-Configuration
    } else {
        Write-Warning "Format UPN invalide."
    }
}

Load-Configuration

$GlobalQuit = $false
do {
    Clear-Host
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host "                  SHERL0CK V5 - META CONSOLE UNIFIEE           " -ForegroundColor White
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " Tenant cible : $(if($GLOBAL:TENANT_NAME){$GLOBAL:TENANT_NAME}else{'En attente de connexion...'})" -ForegroundColor Yellow
    Write-Host " Break-Glass  : $(if($GLOBAL:BREAKGLASS_UPN){$GLOBAL:BREAKGLASS_UPN}else{'NON DEFINI'})" -ForegroundColor $(if($GLOBAL:BREAKGLASS_UPN){"Green"}else{"Red"})
    Write-Host " Audit mode   : $AuditMode" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host " [ 1 ] IDENTITY & SECURITY   (MFA, Acces Conditionnel)" -ForegroundColor White
    Write-Host " [ 2 ] EXCHANGE ONLINE       (Boites, Quotas, Redirections)" -ForegroundColor White
    Write-Host " [ 3 ] AUDIT 365             (Collecte complete + Export Excel)" -ForegroundColor Cyan
    Write-Host " [ 4 ] BREAK-GLASS           (Compte de secours)" -ForegroundColor White
    Write-Host " [ 5 ] JOURNAUX DE SESSION   (Logs et erreurs)" -ForegroundColor White
    Write-Host ""
    Write-Host " [ 0 ] DECONNEXION ET QUITTER" -ForegroundColor Red

    $MenuChoice = Read-Host "`n [>] Module"
    switch ($MenuChoice) {
        "1" {
            Connect-O365Core -AuditMode $AuditMode -SkipModuleInstall:$SkipModuleInstall
            if ($GLOBAL:GRAPH_CONNECTED) {
                $IdentityQuit = $false
                do {
                    Write-Host "`n--- [ IDENTITY & SECURITY ] ---" -ForegroundColor Cyan
                    Write-Host " [M]  MFA Status Report"
                    Write-Host " [C]  Conditional Access Policies"
                    Write-Host " [O]  OAuth/Application Registrations"
                    Write-Host " [R]  Role-Based Access (Privileged Roles)"
                    Write-Host " [E]  Export Identity Security Workbook"
                    Write-Host " [0]  Back"
                    $IdentityChoice = Read-Host "`n [>] Selection"
                    switch ($IdentityChoice) {
                        "M" { Get-MFAStatus -SkipModuleInstall:$SkipModuleInstall }
                        "m" { Get-MFAStatus -SkipModuleInstall:$SkipModuleInstall }
                        "C" { Get-ConditionalAccessPolicies -SkipModuleInstall:$SkipModuleInstall }
                        "c" { Get-ConditionalAccessPolicies -SkipModuleInstall:$SkipModuleInstall }
                        "O" { Get-OAuthApplications -SkipModuleInstall:$SkipModuleInstall }
                        "o" { Get-OAuthApplications -SkipModuleInstall:$SkipModuleInstall }
                        "R" { Get-RoleBasedAccess -SkipModuleInstall:$SkipModuleInstall }
                        "r" { Get-RoleBasedAccess -SkipModuleInstall:$SkipModuleInstall }
                        "E" { Export-IdentitySecurityExcel -SkipModuleInstall:$SkipModuleInstall }
                        "e" { Export-IdentitySecurityExcel -SkipModuleInstall:$SkipModuleInstall }
                        "0" { $IdentityQuit = $true }
                    }
                } while (-not $IdentityQuit)
            }
        }
        "2" {
            Connect-O365Exchange -SkipModuleInstall:$SkipModuleInstall
            if ($GLOBAL:EXO_CONNECTED) { Show-MenuExchange -SkipModuleInstall:$SkipModuleInstall }
        }
        "3" { Show-MenuAudit -AuditMode $AuditMode -SkipModuleInstall:$SkipModuleInstall }
        "4" { Show-MenuBreakGlass }
        "5" { Show-SessionLogs }
        "0" { $GlobalQuit = $true }
    }
} while (-not $GlobalQuit)

try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
if ($GLOBAL:EXO_CONNECTED) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { } }
try { if (Test-Path $GLOBAL:EDGE_TEMP_DIR) { Remove-Item -Path $GLOBAL:EDGE_TEMP_DIR -Recurse -Force -ErrorAction SilentlyContinue } } catch { }
Write-Host "`n[EXIT] Secure shutdown complete." -ForegroundColor Green