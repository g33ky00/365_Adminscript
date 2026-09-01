#Requires -Version 5.1
<#
.SYNOPSIS
    Sherl0ck v4.1 - Unified M365 Administration (Entra ID, Exchange, Intune)
    Full audit, Graph throttling management, Excel export, and handle protection.

.DESCRIPTION
    Sherl0ck v4.1 is a unified administration and audit tool for Microsoft 365.
    It supports two OAuth audit modes: ReadOnly (read-only scopes) and
    ReadWrite (read/write scopes). The default mode is ReadOnly to minimize
    the attack surface.

.PARAMETER AuditMode
    Specifies the audit mode: 'ReadOnly' (default) or 'ReadWrite'.
    ReadOnly uses only read scopes to limit operations to auditing.
    ReadWrite adds privileged write scopes such as Policy.ReadWrite.* and User.ReadWrite.*.

.EXAMPLE
    .\Sherl0ck_v4.1.ps1                          # ReadOnly mode (default)
    .\Sherl0ck_v4.1.ps1 -AuditMode ReadWrite   # ReadWrite mode
#>

[CmdletBinding()]
param(
    [ValidateSet('ReadOnly','ReadWrite')]
    [string]$AuditMode = 'ReadOnly'
)

Set-StrictMode -Off

# =============================================================================
# M2: Collision-safe file path generation
# =============================================================================
function Get-UniqueFilePath {
    param(
        [string]$BasePath
    )
    # M2: If the file exists, append a numeric suffix before the extension
    if (-not (Test-Path $BasePath)) { return $BasePath }
    $dir = Split-Path $BasePath -Parent
    $name = [System.IO.Path]::GetFileNameWithoutExtension($BasePath)
    $ext = [System.IO.Path]::GetExtension($BasePath)
    $counter = 1
    do {
        $newPath = Join-Path $dir "$name`_$counter$ext"
        $counter++
    } while (Test-Path $newPath)
    return $newPath
}

# =============================================================================
# H4: Module version pinning constants
# =============================================================================
$Script:REQUIRED_MODULES = @{
    'Microsoft.Graph.Authentication' = '1.9.3'
    'ExchangeOnlineManagement'       = '3.2.0'
    'ImportExcel'                    = '7.8.0'
}

function Verify-TrustedModule {
    param(
        [string]$ModuleName,
        [string]$RequiredVersion
    )
    # H4: Verify the module source is the official PSGallery
    $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue
    if (-not $repo -or $repo.SourceLocation -ne 'https://www.powershellgallery.com/api/v2/') {
        Write-Host "[WARNING] PSGallery source not verified. Skipping module installation." -ForegroundColor Yellow
        return $false
    }
    # H4: Check that the required minimum version is available
    $available = Find-Module -Name $ModuleName -Repository PSGallery -ErrorAction SilentlyContinue |
        Where-Object { [version]$_.Version -ge [version]$RequiredVersion }
    if (-not $available) {
        Write-Host "[WARNING] Module '$ModuleName' >= $RequiredVersion not found on PSGallery." -ForegroundColor Yellow
        return $false
    }
    return $true
}

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
$GLOBAL:LOG_FILE            = Join-Path $GLOBAL:LOG_DIR "Session_$(Get-Date -Format 'yyyyMMdd').log"

$GLOBAL:EDGE_TEMP_DIR       = Join-Path $env:TEMP "Sherl0ck_Edge_TempProfile"
$GLOBAL:AUDIT_DIR           = Join-Path $env:USERPROFILE "Documents\Sherl0ck_Audits"
if (-not (Test-Path $GLOBAL:AUDIT_DIR)) { New-Item -ItemType Directory -Force -Path $GLOBAL:AUDIT_DIR | Out-Null }

# =============================================================================
# UTILITIES & PARSERS
# =============================================================================
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

# =============================================================================
# LOGGING ENGINE (M1: UPN masking + SecureString encryption)
# =============================================================================
function Mask-SensitiveData {
    param([string]$InputText)
    if (-not $InputText) { return $InputText }
    $Masked = $InputText -replace '(?i)\b[\w\.\-]+@[\w\.\-]+\.\w+\b', '***'
    return $Masked
}

function Add-SessionLog {
    param([string]$Level, [string]$Message, [string]$Details = "")
    # M1: Mask sensitive UPNs before writing
    $SafeMessage = Mask-SensitiveData -InputText $Message
    $SafeDetails = Mask-SensitiveData -InputText $Details
    $LogEntry = "[$(Get-Date -Format 'HH:mm:ss')] [$Level] $SafeMessage"
    if ($SafeDetails) { $LogEntry += " | DETAILS: $SafeDetails" }
    $GLOBAL:SESSION_LOGS += $LogEntry
    try {
        # M1: Encrypt log file via SecureString
        $SecureLog = $LogEntry | ConvertTo-SecureString -AsPlainText -Force
        $Encrypted = $SecureLog | ConvertFrom-SecureString
        Add-Content -Path $GLOBAL:LOG_FILE -Value $Encrypted -Encoding UTF8BOM
    }
    catch {
        # Fallback: log in plaintext if encryption fails
        try { Add-Content -Path $GLOBAL:LOG_FILE -Value $LogEntry -Encoding UTF8BOM } catch {}
    }
}

function Show-SessionLogs {
    Write-Host "`n--- SESSION LOGS (AUDIT & ERRORS) ---" -ForegroundColor Cyan
    Write-Host " Directory : $GLOBAL:LOG_DIR" -ForegroundColor Yellow
    if ($GLOBAL:SESSION_LOGS.Count -eq 0) { Write-Host " [INFO] No events." -ForegroundColor DarkGray }
    else { foreach ($Log in $GLOBAL:SESSION_LOGS) { Write-Host " $Log" -ForegroundColor $(if($Log -match "\[ERROR\]"){"Red"}elseif($Log -match "\[WARNING\]"){"Yellow"}else{"DarkGray"}) } }
    $Choice = Read-Host "`n[O] Open directory | [0] Back"
    if ($Choice -match "^[Oo]$") { Invoke-SafeOpen -FilePath $GLOBAL:LOG_DIR }
}

# =============================================================================
# AUTHENTICATION — GRAPH & EXCHANGE
# =============================================================================
function Connect-O365Core {
    if ($GLOBAL:GRAPH_CONNECTED) { return }
    if ($GLOBAL:EXO_CONNECTED) {
        Write-Host "`n[ERROR] MSAL memory conflict (Exchange module active). Restart the console." -ForegroundColor Red
        Read-Host "Press Enter to return to menu"; return
    }

    Write-Host "`n[INIT] Starting Identity engine (Graph)..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host "`n[INSTALLATION] Module 'Microsoft.Graph.Authentication' not found." -ForegroundColor Red
        if ((Read-Host "Install Microsoft Graph module now? (Y/N)") -match "^[Yy]$") {
            Write-Host " -> Verifying trusted source..." -ForegroundColor Cyan
            # H4: Verify PSGallery source and minimum version before install
            if (-not (Verify-TrustedModule -ModuleName 'Microsoft.Graph.Authentication' -RequiredVersion $Script:REQUIRED_MODULES['Microsoft.Graph.Authentication'])) {
                return
            }
            Write-Host " -> Installing..." -ForegroundColor Cyan
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            try { Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -MinimumVersion $Script:REQUIRED_MODULES['Microsoft.Graph.Authentication'] -ErrorAction Stop }
            catch { Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red; return }
        } else { return }
    }

    try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop } catch {}
    try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction SilentlyContinue | Out-Null } catch {}
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

    do { $GLOBAL:ADMIN_UPN = Read-Host "Global Administrator UPN" } until ($GLOBAL:ADMIN_UPN -match "@")
    $GLOBAL:TARGET_TENANT = ($GLOBAL:ADMIN_UPN -split "@")[1]

    # Base scopes (ReadOnly) for all audit operations
    $ReadOnlyScopes = @(
        "Policy.Read.All",
        "User.Read.All",
        "Organization.Read.All",
        "Domain.Read.All",
        "Device.Read.All",
        "Directory.Read.All",
        "Reports.Read.All",
        "Sites.Read.All",
        "Files.Read.All",
        "AuditLog.Read.All",
        "RoleManagement.Read.Directory",
        "Application.Read.All"
    )

    # Additional ReadWrite scopes (privileged request)
    $ReadWriteScopes = @(
        "Policy.ReadWrite.ConditionalAccess",
        "User.ReadWrite.All"
    )

    if ($AuditMode -eq 'ReadWrite') {
        $Scopes = $ReadOnlyScopes + $ReadWriteScopes
        Write-Host "[AUTH] Mode ReadWrite: privileged scopes enabled (Policy.ReadWrite.*, User.ReadWrite.All)" -ForegroundColor Yellow
    } else {
        $Scopes = $ReadOnlyScopes
        Write-Host "[AUTH] Mode ReadOnly: scopes restricted to read-only" -ForegroundColor Green
    }

    Write-Host "`n[AUTH] Opening Edge in private browsing..." -ForegroundColor Yellow
    try { Start-Process msedge.exe -ArgumentList "--inprivate --new-window --user-data-dir=`"$GLOBAL:EDGE_TEMP_DIR`" https://login.microsoft.com/device" -ErrorAction SilentlyContinue } catch {}

    try { Connect-MgGraph -Scopes $Scopes -TenantId $GLOBAL:TARGET_TENANT -UseDeviceAuthentication -ContextScope Process -NoWelcome }
    catch { Write-Host "`n[ERROR] Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red; Read-Host "Press Enter to continue"; return }

    $Ctx = Get-MgContext
    if ($Ctx.Account -ne $GLOBAL:ADMIN_UPN) {
        Write-Host "`n[ERROR] Connected identity ($($Ctx.Account)) != target ($GLOBAL:ADMIN_UPN)." -ForegroundColor Red
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null; Read-Host "Press Enter to continue"; return
    }

    try { $GLOBAL:TENANT_NAME = (Invoke-MgGraphRequest -Method GET -Uri "v1.0/organization" -ErrorAction Stop).value[0].displayName }
    catch { $GLOBAL:TENANT_NAME = $GLOBAL:TARGET_TENANT }

    $GLOBAL:GRAPH_CONNECTED = $true
    Add-SessionLog "ACTION" "Graph connection successful ($GLOBAL:ADMIN_UPN)"
    Write-Host "[OK] Connected: $($GLOBAL:TENANT_NAME)" -ForegroundColor Green
    Start-Sleep -Seconds 1
}

function Connect-O365Exchange {
    if ($GLOBAL:EXO_CONNECTED) { return }
    Write-Host "`n[INIT] Starting Messaging engine (Exchange)..." -ForegroundColor Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if (-not $GLOBAL:ADMIN_UPN) {
        do { $GLOBAL:ADMIN_UPN = Read-Host "Global Administrator UPN" } until ($GLOBAL:ADMIN_UPN -match "@")
        $GLOBAL:TARGET_TENANT = ($GLOBAL:ADMIN_UPN -split "@")[1]
    }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        if ((Read-Host "EXO module not found. Install? (Y/N)") -match "^[Yy]$") {
            # H4: Verify PSGallery source and minimum version before install
            if (-not (Verify-TrustedModule -ModuleName 'ExchangeOnlineManagement' -RequiredVersion $Script:REQUIRED_MODULES['ExchangeOnlineManagement'])) {
                return
            }
            try { Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -MinimumVersion $Script:REQUIRED_MODULES['ExchangeOnlineManagement'] -ErrorAction Stop }
            catch { Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red; return }
        } else { return }
    }

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -UserPrincipalName $GLOBAL:ADMIN_UPN -ShowProgress $false -InformationAction SilentlyContinue
        $GLOBAL:EXO_CONNECTED = $true
        Write-Host "[OK] Messaging module ready." -ForegroundColor Green
    } catch {
        $GLOBAL:EXO_CONNECTED = $false
        Write-Host "`n[ERROR] Exchange connection: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Press Enter to continue"
    }
}

# =============================================================================
# M365 AUDIT MODULE (COMPACTED EXPORT FUNCTIONS FOR EXCEL)
# =============================================================================
function Get-GraphData {
    param([string]$Uri)
    $Data = @(); $Query = $Uri
    while ($Query) {
        try {
            $Resp = Invoke-MgGraphRequest -Uri $Query -Method GET -ErrorAction Stop
            if ($Resp.value) { $Data += $Resp.value }
            $Query = $Resp.'@odata.nextLink'
        } catch { Add-SessionLog "ERROR" "Graph Query ($Uri)" "$($_.Exception.Message)"; break }
    }
    return $Data
}

function Export-OneDriveUsage {
    Write-Host "`n[STORAGE] Full OneDrive analysis (All users)..." -ForegroundColor Cyan
    try {
        $Users = Get-GraphData "v1.0/users?`$select=id,displayName,userPrincipalName&`$top=999"
        $ODStats = @(); $Count = 0; $MaxRetries = 3

        foreach ($U in $Users) {
            $Count++; Write-Host "   -> Processing $Count/$($Users.Count) : $($U.displayName)..." -NoNewline -ForegroundColor DarkGray
            $RetryCount = 0; $Success = $false

            while (-not $Success -and $RetryCount -le $MaxRetries) {
                try {
                    $Drive = Invoke-MgGraphRequest -Uri "v1.0/users/$($U.id)/drive" -Method GET -ErrorAction Stop
                    $ODStats += [PSCustomObject]@{ User = $U.displayName; UPN = $U.userPrincipalName; Status = "Active"; Used_GB = if($Drive.quota){Convert-BytesToGB $Drive.quota.used}else{0}; Total_GB = if($Drive.quota){Convert-BytesToGB $Drive.quota.total}else{0} }
                    Write-Host " [OK]" -ForegroundColor Green; $Success = $true
                } catch {
                    if ($_.Exception.Message -match "429|Too Many Requests") {
                        $RetryCount++; Start-Sleep -Seconds 5
                    } elseif ($_.Exception.Message -match "404|Not Found") {
                        $ODStats += [PSCustomObject]@{ User = $U.displayName; UPN = $U.userPrincipalName; Status = "Not provisioned (Never signed in)"; Used_GB = 0; Total_GB = 0 }
                        Write-Host " [N/A]" -ForegroundColor Yellow; $Success = $true
                    } else {
                        Write-Host " [ERROR]" -ForegroundColor Red; $Success = $true
                    }
                }
            }
        }
        $Path = Get-UniqueFilePath -BasePath (Join-Path $GLOBAL:AUDIT_DIR "Sherl0ck_OneDrive_Storage_$(Get-Date -Format yyyyMMdd_HHmm).csv")
        $ODStats | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8BOM
        return $ODStats
    } catch { Write-Host " [FAILED] $($_.Exception.Message)" -ForegroundColor Red }
}

function Export-FullAuditExcel {
    Write-Host "`n[EXCEL AUDIT] Generating global workbook..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host " ImportExcel module required for this action." -ForegroundColor Yellow
        # H4: Verify PSGallery source and minimum version before install
        if (-not (Verify-TrustedModule -ModuleName 'ImportExcel' -RequiredVersion $Script:REQUIRED_MODULES['ImportExcel'])) {
            return
        }
        try { Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -MinimumVersion $Script:REQUIRED_MODULES['ImportExcel'] -ErrorAction Stop }
        catch { Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red; return }
        Import-Module ImportExcel
    }

    $DateStr = Get-Date -Format "yyyyMMdd_HHmm"
    $XlsxPath = Get-UniqueFilePath -BasePath (Join-Path $GLOBAL:AUDIT_DIR "Sherl0ck_FullAudit_$($GLOBAL:TENANT_NAME)_$DateStr.xlsx")

    try {
        Write-Host " - Collecting Users..." -ForegroundColor DarkGray
        $Users = Get-GraphData "v1.0/users?`$select=id,displayName,userPrincipalName,mail,usageLocation,assignedLicenses,userType,accountEnabled"
        $Users | Select-Object displayName, userPrincipalName, accountEnabled, userType | Export-Excel -Path $XlsxPath -WorksheetName "Users" -AutoSize -AutoFilter

        Write-Host " - Collecting Licenses..." -ForegroundColor DarkGray
        $Skus = Get-GraphData "v1.0/subscribedSkus"
        $Skus | Select-Object skuPartNumber, @{n='Total';e={$_.prepaidUnits.enabled}}, @{n='Consumed';e={$_.consumedUnits}} | Export-Excel -Path $XlsxPath -WorksheetName "Licenses" -AutoSize -AutoFilter -Append

        Write-Host " - Collecting MFA Status..." -ForegroundColor DarkGray
        $MFA = Get-GraphData "v1.0/reports/authenticationMethods/userRegistrationDetails"
        $MFA | Select-Object userDisplayName, userPrincipalName, isMfaRegistered, @{n='Methods';e={$_.methodsRegistered -join ', '}} | Export-Excel -Path $XlsxPath -WorksheetName "MFA" -AutoSize -AutoFilter -Append

        Write-Host " - Collecting Devices..." -ForegroundColor DarkGray
        $Devices = Get-GraphData "v1.0/devices"
        if ($Devices) { $Devices | Select-Object displayName, operatingSystem, isCompliant, trustType, approximateLastSignInDateTime | Export-Excel -Path $XlsxPath -WorksheetName "Devices" -AutoSize -AutoFilter -Append }

        Write-Host " - Collecting Domains..." -ForegroundColor DarkGray
        $Domains = Get-GraphData "v1.0/domains"
        $Domains | Select-Object id, isVerified, isDefault | Export-Excel -Path $XlsxPath -WorksheetName "Domains" -AutoSize -AutoFilter -Append

        Write-Host " - Collecting OneDrive..." -ForegroundColor DarkGray
        $OD = Export-OneDriveUsage
        if ($OD) { $OD | Export-Excel -Path $XlsxPath -WorksheetName "OneDrive" -AutoSize -AutoFilter -Append }

        Write-Host "`n[SUCCESS] Excel audit complete: $XlsxPath" -ForegroundColor Green
        Add-SessionLog "ACTION" "Excel export successful" $XlsxPath
        Invoke-SafeOpen -FilePath $XlsxPath
    } catch {
        Write-Host "`n[ERROR] Excel generation: $($_.Exception.Message)" -ForegroundColor Red
        Add-SessionLog "ERROR" "Excel Export" $_.Exception.Message
    }
}

function Show-MenuAudit {
    Connect-O365Core
    if (-not $GLOBAL:GRAPH_CONNECTED) { return }

    $QuitAudit = $false
    do {
        Write-Host "`n--- [ M365 AUDIT ] ---" -ForegroundColor Cyan
        Write-Host " [E]  Export multi-tab EXCEL (Recommended: Users, MFA, Devices, Licenses, OneDrive)"
        Write-Host " [0]  Back"

        $Choice = Read-Host "Selection"
        switch ($Choice) {
            "E" { Export-FullAuditExcel }
            "e" { Export-FullAuditExcel }
            "0" { $QuitAudit = $true }
        }
    } while (-not $QuitAudit)
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
Write-Host "`n[EXIT] Secure shutdown complete." -ForegroundColor Green
Start-Sleep -Seconds 1
Clear-Host