#Requires -Version 5.1
<#
.SYNOPSIS
    Sherl0ck v4.1 - Unified M365 Administration (Entra ID, Exchange, Intune)
    Full audit, Graph throttling management, Excel export, and handle protection.

.DESCRIPTION
    Sherl0ck v4.1 is a modular administration and audit tool for Microsoft 365.
    It supports two OAuth audit modes: ReadOnly (read-only scopes) and
    ReadWrite (read/write scopes). The default mode is ReadOnly to minimize
    the attack surface.

.PARAMETER AuditMode
    Specifies the audit mode: 'ReadOnly' (default) or 'ReadWrite'.
    ReadOnly uses only read scopes to limit operations to auditing.
    ReadWrite is reserved for future write operations (currently warns that no write functions are implemented).

.PARAMETER SkipModuleInstall
    If set, skips automatic installation of required PowerShell modules.
    The script will warn if a required module is missing but will not prompt to install.

.EXAMPLE
    .\Sherl0ck_v4.1.ps1                          # ReadOnly mode (default)
    .\Sherl0ck_v4.1.ps1 -AuditMode ReadWrite   # ReadWrite mode
    .\Sherl0ck_v4.1.ps1 -SkipModuleInstall     # Skip auto-install of modules
#>

[CmdletBinding()]
param(
    [ValidateSet('ReadOnly','ReadWrite')]
    [string]$AuditMode = 'ReadOnly',
    [switch]$SkipModuleInstall
)

Set-StrictMode -Version 3.0  # Point 2: Enabled strict mode (was -Off)

# =============================================================================
# GLOBAL VARIABLES
# =============================================================================
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
# Point 1 fix: Log key stored in a SEPARATE directory from logs — key never co-located with encrypted files
$GLOBAL:LOG_KEY_DIR         = Join-Path $env:LOCALAPPDATA "Sherl0ck_Secrets"
if (-not (Test-Path $GLOBAL:LOG_KEY_DIR)) { New-Item -ItemType Directory -Force -Path $GLOBAL:LOG_KEY_DIR | Out-Null }
$GLOBAL:LOG_FILE            = Join-Path $GLOBAL:LOG_DIR "Session_$(Get-Date -Format 'yyyyMMdd').log"
$GLOBAL:EDGE_TEMP_DIR       = Join-Path $env:TEMP "Sherl0ck_Edge_TempProfile"
$GLOBAL:AUDIT_DIR           = Join-Path $env:USERPROFILE "Documents\Sherl0ck_Audits"
if (-not (Test-Path $GLOBAL:AUDIT_DIR)) { New-Item -ItemType Directory -Force -Path $GLOBAL:AUDIT_DIR | Out-Null }

# =============================================================================
# MODULE LOADING (L1: Modular architecture)
# =============================================================================
$ModulesDir = Join-Path $PSScriptRoot "..\Modules"

# Import modules in dependency order: Utils → UI → Auth → Audit
foreach ($moduleName in @('Sherl0ck.Utils', 'Sherl0ck.UI', 'Sherl0ck.Auth', 'Sherl0ck.Audit')) {
    $modulePath = Join-Path $ModulesDir "$moduleName.psm1"
    if (Test-Path $modulePath) {
        Import-Module $modulePath -Force -ErrorAction Stop
        Write-Host "[INIT] Loaded module: $moduleName" -ForegroundColor DarkGray
    } else {
        Write-Host "[ERROR] Module not found: $moduleName at $modulePath" -ForegroundColor Red
    }
}

# =============================================================================
# MAIN ENTRY POINT
# =============================================================================
$GlobalQuit = $false

do {
    Clear-Host
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host "                  SHERL0CK V4.1 - META CONSOLE                " -ForegroundColor White
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " Target tenant : $(if($GLOBAL:TENANT_NAME){$GLOBAL:TENANT_NAME}else{'Awaiting connection...'})" -ForegroundColor Yellow
    Write-Host " Reports dir   : $GLOBAL:AUDIT_DIR" -ForegroundColor DarkGray
    Write-Host " Audit mode    : $AuditMode" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host " [ 1 ] IDENTITY & SECURITY   (MFA, Conditional Access, OAuth Apps)" -ForegroundColor White
    Write-Host " [ 2 ] EXCHANGE ONLINE        (Mailboxes, Quotas, Redirects)" -ForegroundColor White
    Write-Host " [ 3 ] M365 AUDIT             (Complete collection + Excel Export)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host " [ 4 ] SESSION LOGS          (Logs & errors)" -ForegroundColor White
    Write-Host ""
    Write-Host " [ 0 ] DISCONNECT & QUIT" -ForegroundColor Red

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
        "2" { Connect-O365Exchange -SkipModuleInstall:$SkipModuleInstall }
        "3" { Show-MenuAudit -AuditMode $AuditMode -SkipModuleInstall:$SkipModuleInstall }
        "4" { Show-SessionLogs }
        "0" { $GlobalQuit = $true }
    }
} while (-not $GlobalQuit)

try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { Add-SessionLog "WARNING" "Disconnect MgGraph" $_.Exception.Message }
if ($GLOBAL:EXO_CONNECTED) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { Add-SessionLog "WARNING" "Disconnect EXO" $_.Exception.Message } }
try { if (Test-Path $GLOBAL:EDGE_TEMP_DIR) { Remove-Item -Path $GLOBAL:EDGE_TEMP_DIR -Recurse -Force -ErrorAction SilentlyContinue } } catch { Add-SessionLog "WARNING" "Edge temp cleanup" $_.Exception.Message }
Write-Host "`n[EXIT] Secure shutdown complete." -ForegroundColor Green
Start-Sleep -Seconds 1
Clear-Host
