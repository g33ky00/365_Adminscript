<#
.SYNOPSIS
    Sherl0ck UI module — user interface and logging functions.

.DESCRIPTION
    Contains session logging with M1 security: UPN masking via
    Mask-SensitiveData and log encryption via ConvertFrom-SecureString.

.NOTES
    Part of the 365_Adminscript modular architecture.
#>

# M1: Mask sensitive data (UPNs, emails) before logging
function Mask-SensitiveData {
    param([string]$InputText)
    if (-not $InputText) { return $InputText }
    $Masked = $InputText -replace '(?i)\b[\w\.\-]+@[\w\.\-]+\.\w+\b', '***'
    return $Masked
}

# M1: Encrypted session logging
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

Export-ModuleMember -Function Mask-SensitiveData, Add-SessionLog, Show-SessionLogs
