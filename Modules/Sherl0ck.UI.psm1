<#
.SYNOPSIS
    Sherl0ck UI module — user interface and logging functions.

.DESCRIPTION
    Contains session logging with M1 security: UPN masking via
    Mask-SensitiveData and log encryption via ConvertFrom-SecureString
    using a portable AES key (key file) for cross-machine compatibility.

.PARAMETER InputText
    The text to mask (UPNs/emails replaced with ***).

.PARAMETER Level
    The log level: INFO, WARN, ERROR, CRITICAL, etc.

.PARAMETER Message
    The log message text.

.PARAMETER Details
    Optional additional details to include in the log entry.

.EXAMPLE
    PS> Mask-SensitiveData -InputText "admin@contoso.com"
    Returns "***"

.EXAMPLE
    PS> Add-SessionLog -Level "INFO" -Message "Export started" -Details "user@contoso.com"
    Logs the event with UPN masked in encrypted log file.

.NOTES
    Part of the 365_Adminscript modular architecture.
    Log encryption uses a key file stored alongside the logs directory,
    enabling decryption on any machine with access to the key file.
#>
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
        # M1: Encrypt log via portable AES key stored in a SEPARATE secrets directory
        # Point 1 fix (revised): Key file stored in $GLOBAL:LOG_KEY_DIR, NOT in $GLOBAL:LOG_DIR.
        # This prevents an attacker with log directory access from also obtaining the key.
        # Plus explicit ACL with FileSystemAccessRule for the current user (not just inheritance removal).
        $KeyFilePath = Join-Path $GLOBAL:LOG_KEY_DIR "log_key.key"
        if (-not (Test-Path $KeyFilePath)) {
            # Generate a random 256-bit AES key on first run
            $Key = New-Object byte[] 32
            [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($Key)
            $Key | Out-File -FilePath $KeyFilePath -Encoding ASCII

            # Point 1 fix (revised): Explicit ACL — current user only, full control, no inheritance
            $acl = Get-Acl -Path $KeyFilePath
            $acl.SetAccessRuleProtection($true, $false)  # Disable inheritance, remove inherited rules
            $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
            $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $currentUser,
                "FullControl",
                "None",
                "None",
                "Allow"
            )
            $acl.SetOwner($currentUser)
            $acl.AddAccessRule($accessRule)
            Set-Acl -Path $KeyFilePath -AclObject $acl
        }
        $Key = Get-Content -Path $KeyFilePath -Encoding ASCII | ForEach-Object { [byte]$_ }
        $SecureLog = $LogEntry | ConvertTo-SecureString -AsPlainText -Force
        $Encrypted = $SecureLog | ConvertFrom-SecureString -Key $Key
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
