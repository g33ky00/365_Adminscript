<#
.SYNOPSIS
    Sherl0ck UI module — user interface and logging functions.

.DESCRIPTION
    Contains session logging with M1 security: UPN masking via
    Mask-SensitiveData and log encryption via ConvertFrom-SecureString
    using a portable AES key stored in SecretManagement vault 'Sherl0ckVault'.
    Exports can be protected/unprotected with AES-256-GCM + PBKDF2.

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
#>
function Mask-SensitiveData {
    param([AllowNull()][string]$InputText)
    if ($null -eq $InputText) { return $null }
    if (-not $InputText) { return $InputText }
    $Masked = $InputText -replace '(?i)\b[\w\.\-\+]+@[\w\.\-]+\.\w{2,}\b', '***'
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

    # Require SecretManagement + Sherl0ckVault — no plaintext fallback
    if (-not (Get-Module -ListAvailable -Name Microsoft.PowerShell.SecretManagement)) {
        throw "Microsoft.PowerShell.SecretManagement module is required but not available."
    }
    Import-Module Microsoft.PowerShell.SecretManagement -ErrorAction Stop
    $vault = Get-SecretVault -Name 'Sherl0ckVault' -ErrorAction SilentlyContinue
    if (-not $vault) {
        throw "SecretManagement vault 'Sherl0ckVault' is not available. Refusing to run."
    }

    $secretName = 'Sherl0ckLogKey'
    $keyB64 = $null
    try { $keyB64 = Get-Secret -Name $secretName -Vault 'Sherl0ckVault' -AsPlainText -ErrorAction Stop } catch { $keyB64 = $null }
    if (-not $keyB64) {
        $Key = New-Object byte[] 32
        [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($Key)
        $keyB64 = [Convert]::ToBase64String($Key)
        Set-Secret -Name $secretName -Secret $keyB64 -Vault 'Sherl0ckVault' -ErrorAction Stop | Out-Null
    }
    $Key = [Convert]::FromBase64String($keyB64)

    $SecureLog = New-Object System.Security.SecureString
    $LogEntry.ToCharArray() | ForEach-Object { $SecureLog.AppendChar($_) }
    $SecureLog.MakeReadOnly()
    $Encrypted = $SecureLog | ConvertFrom-SecureString -Key $Key
    if ($GLOBAL:LOG_FILE -and $Encrypted) {
        $logDir = Split-Path $GLOBAL:LOG_FILE -Parent
        if ($logDir -and -not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
        Add-Content -Path $GLOBAL:LOG_FILE -Value $Encrypted -Encoding UTF8BOM
    }
}

function Protect-Export {
    param(
        [Parameter(Mandatory)][string]$InputPath,
        [string]$OutputPath,
        [System.Security.SecureString]$Password
    )
    if (-not $OutputPath) { $OutputPath = $InputPath }
    if (-not $Password) { $Password = Read-Host -Prompt 'Export password' -AsSecureString }
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    try {
        $bytes = [IO.File]::ReadAllBytes($InputPath)
        # Apply UPN masking only when the content is text (best-effort detection).
        try {
            $text = [Text.Encoding]::UTF8.GetString($bytes)
            if (-not $text.Contains([char]0)) {
                $masked = Mask-SensitiveData -InputText $text
                $bytes = [Text.Encoding]::UTF8.GetBytes($masked)
            }
        } catch {
            # Binary content: leave bytes untouched.
        }

        $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
        $salt = New-Object byte[] 16; $rng.GetBytes($salt)
        $nonce = New-Object byte[] 12; $rng.GetBytes($nonce)
        $pbkdf = New-Object Security.Cryptography.Rfc2898DeriveBytes($plain, $salt, 100000)
        $key = $pbkdf.GetBytes(32)

        $cipher = New-Object byte[] $bytes.Length
        $tag = New-Object byte[] 16
        $aes = New-Object Security.Cryptography.AesGcm($key)
        $aes.Encrypt($nonce, $bytes, $cipher, $tag)

        $header = [Text.Encoding]::ASCII.GetBytes('SHERL0CK-ENC-v1')
        $out = New-Object byte[] ($header.Length + 16 + 12 + $cipher.Length + 16)
        [Array]::Copy($header, 0, $out, 0, $header.Length)
        [Array]::Copy($salt, 0, $out, $header.Length, 16)
        [Array]::Copy($nonce, 0, $out, $header.Length + 16, 12)
        [Array]::Copy($cipher, 0, $out, $header.Length + 28, $cipher.Length)
        [Array]::Copy($tag, 0, $out, $header.Length + 28 + $cipher.Length, 16)
        [IO.File]::WriteAllBytes($OutputPath, $out)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Unprotect-Export {
    param(
        [Parameter(Mandatory)][string]$InputPath,
        [string]$OutputPath,
        [System.Security.SecureString]$Password
    )
    if (-not $Password) { $Password = Read-Host -Prompt 'Export password' -AsSecureString }
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    try {
        $data = [IO.File]::ReadAllBytes($InputPath)
        $header = [Text.Encoding]::ASCII.GetBytes('SHERL0CK-ENC-v1')
        $headerLen = $header.Length
        if ($data.Length -lt $headerLen + 16 + 12 + 16) { throw "Invalid encrypted file." }
        for ($i = 0; $i -lt $headerLen; $i++) {
            if ($data[$i] -ne $header[$i]) { throw "Invalid file header." }
        }
        $salt = New-Object byte[] 16
        $nonce = New-Object byte[] 12
        $tag = New-Object byte[] 16
        [Array]::Copy($data, $headerLen, $salt, 0, 16)
        [Array]::Copy($data, $headerLen + 16, $nonce, 0, 12)
        $cipherLen = $data.Length - $headerLen - 16 - 12 - 16
        $cipher = New-Object byte[] $cipherLen
        [Array]::Copy($data, $headerLen + 28, $cipher, 0, $cipherLen)
        [Array]::Copy($data, $data.Length - 16, $tag, 0, 16)

        $pbkdf = New-Object Security.Cryptography.Rfc2898DeriveBytes($plain, $salt, 100000)
        $key = $pbkdf.GetBytes(32)
        $aes = New-Object Security.Cryptography.AesGcm($key)
        $plainBytes = New-Object byte[] $cipherLen
        $aes.Decrypt($nonce, $cipher, $tag, $plainBytes)
        if ($OutputPath) {
            [IO.File]::WriteAllBytes($OutputPath, $plainBytes)
        } else {
            try {
                $text = [Text.Encoding]::UTF8.GetString($plainBytes)
                if ($text.Contains([char]0)) { return $plainBytes }
                return $text
            } catch {
                return $plainBytes
            }
        }
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Show-SessionLogs {
    Write-Host "`n--- SESSION LOGS (AUDIT & ERRORS) ---" -ForegroundColor Cyan
    Write-Host " Directory : $GLOBAL:LOG_DIR" -ForegroundColor Yellow
    if ($GLOBAL:SESSION_LOGS.Count -eq 0) { Write-Host " [INFO] No events." -ForegroundColor DarkGray }
    else { foreach ($Log in $GLOBAL:SESSION_LOGS) { Write-Host " $Log" -ForegroundColor $(if($Log -match "\[ERROR\]"){"Red"}elseif($Log -match "\[WARNING\]"){"Yellow"}else{"DarkGray"}) } }
    $Choice = if ($Host.UI.RawUI) { Read-Host "`n[O] Open directory | [0] Back" } else { "0" }
    if ($Choice -match "^[Oo]$") { Invoke-SafeOpen -FilePath $GLOBAL:LOG_DIR }
}

function Get-LocalizedString {
    param(
        [Parameter(Mandatory)][string]$Key,
        [string]$Language = 'en-US'
    )
    $moduleDir = Split-Path $PSScriptRoot -Parent
    $localePath = Join-Path $moduleDir "Locales\\$Language.json"
    if (-not (Test-Path $localePath)) {
        $localePath = Join-Path $moduleDir "Locales\\en-US.json"
    }
    if (-not (Test-Path $localePath)) {
        $localePath = Join-Path (Get-Location) "Locales\\$Language.json"
    }
    try {
        $locale = Get-Content -Path $localePath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch { return $Key }
    $parts = $Key -split '\\.'
    $node = $locale
    foreach ($p in $parts) {
        if ($node -and $node.PSObject.Properties.Name -contains $p) { $node = $node.$p } else { return $Key }
    }
    if ($node -is [string]) { return $node }
    return $Key
}

Export-ModuleMember -Function Mask-SensitiveData, Add-SessionLog, Show-SessionLogs, Get-LocalizedString, Protect-Export, Unprotect-Export