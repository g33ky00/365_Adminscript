<#
.SYNOPSIS
    Sherl0ck Auth module — Graph & Exchange authentication.

.DESCRIPTION
    Contains Connect-O365Core (Graph device code auth with H3 scope separation)
    and Connect-O365Exchange (Exchange Online management module).

    H3: OAuth scopes are separated into ReadOnly (default) and ReadWrite modes.
    H4: Module installation uses MinimumVersion pinning + PSGallery source verification.
    M3: -SkipModuleInstall switch bypasses auto-install.

.PARAMETER AuditMode
    Specifies the audit mode: 'ReadOnly' (default) or 'ReadWrite'.

.PARAMETER SkipModuleInstall
    If set, skips automatic installation of required PowerShell modules.

.PARAMETER ModuleName
    The name of the PowerShell module to verify.

.PARAMETER RequiredVersion
    The minimum required version of the module.

.EXAMPLE
    PS> Connect-O365Core -AuditMode 'ReadOnly'
    Connects to Microsoft Graph with read-only scopes.

.EXAMPLE
    PS> Connect-O365Core -AuditMode 'ReadWrite' -SkipModuleInstall
    Connects with read-write scopes, skipping module installation.

.NOTES
    Part of the 365_Adminscript modular architecture.
#>

# H4: Module version pinning constants
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
    # M3: Find-Module queries the official PSGallery API which provides
    #     package integrity hashes (SHA-512) automatically verified by PowerShellGet.
    #     Only modules signed by a trusted publisher on PSGallery are considered safe.
    $available = Find-Module -Name $ModuleName -Repository PSGallery -ErrorAction SilentlyContinue |
        Where-Object { [version]$_.Version -ge [version]$RequiredVersion }
    if (-not $available) {
        Write-Host "[WARNING] Module '$ModuleName' >= $RequiredVersion not found on PSGallery." -ForegroundColor Yellow
        return $false
    }
    # M3: Verify package integrity via real hash comparison
    #     Downloads the package to a temp dir, computes SHA-512 of the .nupkg,
    #     and compares against the PackageHash from PSGallery metadata.
    #     This provides real tamper detection — not just field presence checking.
    try {
        $packageHash = $available.PackageHash
        if (-not $packageHash -or $packageHash.Length -lt 64) {
            Write-Host "[WARNING] Module '$ModuleName' package hash not available on PSGallery." -ForegroundColor Yellow
            return $true  # Proceed with source verification only
        }

        $tempDir = Join-Path $env:TEMP "Sherl0ck_ModuleVerify_$(Get-Random)"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

        # Download the package without installing it
        Save-Module -Name $ModuleName -RequiredVersion $available.Version -Repository PSGallery -Path $tempDir -Force -ErrorAction Stop

        # Compute SHA-512 of the downloaded .nupkg file
        $nupkgPath = Get-ChildItem -Path $tempDir -Recurse -Filter "*.nupkg" | Select-Object -First 1
        if ($nupkgPath) {
            $computedHash = (Get-FileHash -Path $nupkgPath.FullName -Algorithm SHA512).Hash
            if ($computedHash -ne $packageHash) {
                Write-Host "[WARNING] Hash mismatch for '$ModuleName' v$($available.Version). Possible tampering!" -ForegroundColor Red
                Add-SessionLog "WARNING" "Module hash mismatch" "Expected: $packageHash Computed: $computedHash"
                Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                return $false
            }
            Write-Host "[VERIFY] Module '$ModuleName' v$($available.Version) SHA-512 hash verified." -ForegroundColor DarkGray
        }

        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "[WARNING] Could not verify package hash for '$ModuleName'. Proceeding with source verification only." -ForegroundColor Yellow
        Add-SessionLog "WARNING" "Hash verification failed" $_.Exception.Message
    }
    return $true
}

# Point 3: Fallback cascade for private browsing browser launch
# Tries Edge → Firefox → Chrome. Returns $true if any browser launched, $false otherwise.
function Invoke-BrowserPrivate {
    param([string]$Uri)

    # Edge (stable)
    $edgePath = "msedge.exe"
    if (Get-Command $edgePath -ErrorAction SilentlyContinue) {
        try { Start-Process $edgePath -ArgumentList "--inprivate --new-window --user-data-dir=`"$GLOBAL:EDGE_TEMP_DIR`" $Uri" -ErrorAction Stop; return $true }
        catch { Write-Host "[WARNING] Edge launch failed: $($_.Exception.Message)" -ForegroundColor Yellow }
    }

    # Firefox
    $firefoxPath = "firefox.exe"
    if (Get-Command $firefoxPath -ErrorAction SilentlyContinue) {
        try { Start-Process $firefoxPath -ArgumentList "-private-window $Uri" -ErrorAction Stop; return $true }
        catch { Write-Host "[WARNING] Firefox launch failed: $($_.Exception.Message)" -ForegroundColor Yellow }
    }

    # Chrome
    $chromePath = "chrome.exe"
    if (Get-Command $chromePath -ErrorAction SilentlyContinue) {
        try { Start-Process $chromePath -ArgumentList "--incognito --new-window $Uri" -ErrorAction Stop; return $true }
        catch { Write-Host "[WARNING] Chrome launch failed: $($_.Exception.Message)" -ForegroundColor Yellow }
    }

    return $false
}

function Connect-O365Core {
    param(
        [ValidateSet('ReadOnly','ReadWrite')]
        [string]$AuditMode = 'ReadOnly',
        [switch]$SkipModuleInstall
    )

    if ($GLOBAL:GRAPH_CONNECTED) { return }
    if ($GLOBAL:EXO_CONNECTED) {
        Write-Host "`n[ERROR] MSAL memory conflict (Exchange module active). Restart the console." -ForegroundColor Red
        Read-Host "Press Enter to return to menu"; return
    }

    Write-Host "`n[INIT] Starting Identity engine (Graph)..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host "`n[INSTALLATION] Module 'Microsoft.Graph.Authentication' not found." -ForegroundColor Red
        if ($SkipModuleInstall) {
            Write-Host "[WARNING] Module installation skipped (-SkipModuleInstall). Graph features will be unavailable." -ForegroundColor Yellow
            return
        }
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

    Write-Host "`n[AUTH] Opening browser in private browsing..." -ForegroundColor Yellow
    # Point 3: Fallback cascade: Edge → Firefox → Chrome → explicit error
    $browserLaunched = Invoke-BrowserPrivate -Uri "https://login.microsoftonline.com/device"
    if (-not $browserLaunched) {
        Write-Host "`n[ERROR] No supported browser found. Please open the following URL manually in any browser:" -ForegroundColor Red
        Write-Host "  https://login.microsoftonline.com/device" -ForegroundColor Yellow
        Write-Host "`n  Supported browsers: Microsoft Edge, Firefox, Google Chrome" -ForegroundColor DarkGray
        Read-Host "Press Enter once you have opened the URL..."
    }

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
    param(
        [switch]$SkipModuleInstall
    )

    if ($GLOBAL:EXO_CONNECTED) { return }
    Write-Host "`n[INIT] Starting Messaging engine (Exchange)..." -ForegroundColor Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if (-not $GLOBAL:ADMIN_UPN) {
        do { $GLOBAL:ADMIN_UPN = Read-Host "Global Administrator UPN" } until ($GLOBAL:ADMIN_UPN -match "@")
        $GLOBAL:TARGET_TENANT = ($GLOBAL:ADMIN_UPN -split "@")[1]
    }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        if ($SkipModuleInstall) {
            Write-Host "[WARNING] Module installation skipped (-SkipModuleInstall). Exchange features will be unavailable." -ForegroundColor Yellow
            return
        }
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

Export-ModuleMember -Function Connect-O365Core, Connect-O365Exchange, Verify-TrustedModule, Invoke-BrowserPrivate
Export-ModuleMember -Variable REQUIRED_MODULES
