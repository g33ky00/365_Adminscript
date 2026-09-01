<#
.SYNOPSIS
    Sherl0ck Audit module — M365 data collection and export functions.

.DESCRIPTION
    Contains Get-GraphData (throttled Graph API pagination),
    Export-OneDriveUsage (per-user storage audit with M2 collision-safe CSV),
    Export-FullAuditExcel (consolidated multi-worksheet workbook),
    and Identity/Security audit stubs (MFA, CA policies, OAuth apps, RBAC).

    M2: File paths use Get-UniqueFilePath to prevent overwrites.
    M3: ImportExcel installation uses MinimumVersion + SkipModuleInstall support.
    Point 8: Identity & Security audit functions for targeted menu [1] operations.

.PARAMETER Uri
    The Microsoft Graph API endpoint URI to query.

.PARAMETER SkipModuleInstall
    If set, skips automatic installation of the ImportExcel module.

.PARAMETER AuditMode
    Specifies the audit mode: 'ReadOnly' (default) or 'ReadWrite'.

.EXAMPLE
    PS> Get-GraphData -Uri "v1.0/users"
    Returns all users from Microsoft Graph with automatic pagination.

.EXAMPLE
    PS> Export-FullAuditExcel -SkipModuleInstall
    Generates the full audit workbook, skipping ImportExcel installation.

.EXAMPLE
    PS> Export-IdentitySecurityExcel
    Generates the Identity & Security audit workbook (MFA, CA, OAuth, RBAC).

.NOTES
    Part of the 365_Adminscript modular architecture.
#>

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
    param(
        [switch]$SkipModuleInstall
    )

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
                        # FIX: Log Graph errors instead of silent data loss (error entry added to ODStats)
                        $errMsg = $_.Exception.Message
                        Add-SessionLog "ERROR" "OneDrive query failed for $($U.userPrincipalName)" $errMsg
                        $ODStats += [PSCustomObject]@{ User = $U.displayName; UPN = $U.userPrincipalName; Status = "Error ($errMsg)"; Used_GB = 0; Total_GB = 0 }
                        Write-Host " [ERROR]" -ForegroundColor Red; $Success = $true
                    }
                }
            }
        }
        # M2: Collision-safe file path
        $Path = Get-UniqueFilePath -BasePath (Join-Path $GLOBAL:AUDIT_DIR "Sherl0ck_OneDrive_Storage_$(Get-Date -Format yyyyMMdd_HHmm).csv")
        $ODStats | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8BOM
        return $ODStats
    } catch { Write-Host " [FAILED] $($_.Exception.Message)" -ForegroundColor Red }
}

function Export-FullAuditExcel {
    param(
        [switch]$SkipModuleInstall
    )

    Write-Host "`n[EXCEL AUDIT] Generating global workbook..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host " ImportExcel module required for this action." -ForegroundColor Yellow
        if ($SkipModuleInstall) {
            Write-Host "[WARNING] Module installation skipped (-SkipModuleInstall). Excel export unavailable." -ForegroundColor Yellow
            return
        }
        # H4: Verify PSGallery source and minimum version before install
        if (-not (Verify-TrustedModule -ModuleName 'ImportExcel' -RequiredVersion $Script:REQUIRED_MODULES['ImportExcel'])) {
            return
        }
        try { Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -MinimumVersion $Script:REQUIRED_MODULES['ImportExcel'] -ErrorAction Stop }
        catch { Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red; return }
        Import-Module ImportExcel
    }

    $DateStr = Get-Date -Format "yyyyMMdd_HHmm"
    # M2: Collision-safe file path
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
        $OD = Export-OneDriveUsage -SkipModuleInstall:$SkipModuleInstall
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
    param(
        [ValidateSet('ReadOnly','ReadWrite')]
        [string]$AuditMode = 'ReadOnly',
        [switch]$SkipModuleInstall
    )

    Connect-O365Core -AuditMode $AuditMode -SkipModuleInstall:$SkipModuleInstall
    if (-not $GLOBAL:GRAPH_CONNECTED) { return }

    $QuitAudit = $false
    do {
        Write-Host "`n--- [ M365 AUDIT ] ---" -ForegroundColor Cyan
        Write-Host " [E]  Export multi-tab EXCEL (Recommended: Users, MFA, Devices, Licenses, OneDrive)"
        Write-Host " [0]  Back"

        $Choice = Read-Host "Selection"
        switch ($Choice) {
            "E" { Export-FullAuditExcel -SkipModuleInstall:$SkipModuleInstall }
            "e" { Export-FullAuditExcel -SkipModuleInstall:$SkipModuleInstall }
            "0" { $QuitAudit = $true }
        }
    } while (-not $QuitAudit)
}

# =============================================================================
# IDENTITY & SECURITY AUDIT STUBS (Point 8 — menu [1] extension)
# These stubs follow existing patterns (Get-GraphData, Add-SessionLog error handling).
# TODO: Implement full collection logic in subsequent iteration.
# =============================================================================

function Get-MFAStatus {
    <#
    .SYNOPSIS
        Collects MFA registration status for all users.

    .DESCRIPTION
        Queries Graph reports/authenticationMethods/userRegistrationDetails
        to identify users with/without MFA registered.
        TODO: Add per-method-type breakdown (FIDO2, SMS, Authenticator, etc.)

    .EXAMPLE
        PS> Get-MFAStatus
        Returns MFA registration status for all users.
    #>
    param([switch]$SkipModuleInstall)
    Write-Host "`n[IDENTITY] Collecting MFA status..." -ForegroundColor Cyan
    try {
        $MFA = Get-GraphData "v1.0/reports/authenticationMethods/userRegistrationDetails"
        $Summary = $MFA | Group-Object -Property isMfaRegistered | ForEach-Object {
            [PSCustomObject]@{ Status = if ($_.Name -eq "True") {"MFA Registered"} else {"MFA Not Registered"}; Count = $_.Count }
        }
        Write-Host "  MFA Registered: $(($Summary | Where-Object Status -eq 'MFA Registered').Count)" -ForegroundColor Green
        Write-Host "  MFA Not Registered: $(($Summary | Where-Object Status -eq 'MFA Not Registered').Count)" -ForegroundColor Red
        Add-SessionLog "ACTION" "MFA status collected" "Total: $($MFA.Count) users"
        return $MFA
    }
    catch {
        Add-SessionLog "ERROR" "Get-MFAStatus" $_.Exception.Message
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-ConditionalAccessPolicies {
    <#
    .SYNOPSIS
        Lists all Conditional Access policies with their state.

    .DESCRIPTION
        Queries Graph identity/conditionalAccess/policies to collect CA policy names,
        states (enabled/disabled/report-only), and conditions.
        TODO: Add export to Excel with policy state comparison.

    .EXAMPLE
        PS> Get-ConditionalAccessPolicies
        Returns list of CA policies with enforced/report-only states.
    #>
    param([switch]$SkipModuleInstall)
    Write-Host "`n[IDENTITY] Collecting Conditional Access policies..." -ForegroundColor Cyan
    try {
        $CAPolicies = Get-GraphData "v1.0/identity/conditionalAccess/policies"
        $CAPolicies | ForEach-Object {
            $state = if ($_.state -eq "enabledForReportingButNotEnforced") {"Report-only"} elseif ($_.state -eq "enabled") {"Enforced"} else { $_.state }
            $color = if ($state -eq "Report-only") {"Yellow"} elseif ($state -eq "Enforced") {"Green"} else {"DarkGray"}
            Write-Host "  $($_.displayName) [$state]" -ForegroundColor $color
        }
        Add-SessionLog "ACTION" "CA policies collected" "Total: $($CAPolicies.Count) policies"
        # Note: $GLOBAL:POLICY_STATE from main script is used for comparison
        return $CAPolicies
    }
    catch {
        Add-SessionLog "ERROR" "Get-ConditionalAccessPolicies" $_.Exception.Message
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-OAuthApplications {
    <#
    .SYNOPSIS
        Lists all OAuth/application registrations with consent status.

    .DESCRIPTION
        Queries Graph applications to collect app registrations, redirect URIs,
        reply URLs, and required OAuth permissions.
        TODO: Add analysis of over-privileged app permissions (least-privilege check).

    .EXAMPLE
        PS> Get-OAuthApplications
        Returns all app registrations with redirect URIs and permissions.
    #>
    param([switch]$SkipModuleInstall)
    Write-Host "`n[IDENTITY] Collecting OAuth applications..." -ForegroundColor Cyan
    try {
        $Apps = Get-GraphData "v1.0/applications?`$select=displayName,appId,web,spa,signInAudience,requiredResourceAccess"
        $Apps | ForEach-Object {
            $redirects = ($_.web.redirectUris -join ", ") + ($_.spa.redirectUris -join ", ")
            Write-Host "  $($_.displayName) | $($_.appId)" -ForegroundColor DarkGray
            if ($redirects) { Write-Host "    Redirects: $redirects" -ForegroundColor DarkGray }
        }
        Add-SessionLog "ACTION" "OAuth applications collected" "Total: $($Apps.Count) apps"
        return $Apps
    }
    catch {
        Add-SessionLog "ERROR" "Get-OAuthApplications" $_.Exception.Message
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-RoleBasedAccess {
    <#
    .SYNOPSIS
        Collects role-based access control (RBAC) information.

    .DESCRIPTION
        Queries Graph directoryRoles and roleAssignments to identify
        users with privileged roles (Global Admin, etc.).
        TODO: Add cross-reference with MFA status for privileged accounts.

    .EXAMPLE
        PS> Get-RoleBasedAccess
        Returns all privileged role assignments.
    #>
    param([switch]$SkipModuleInstall)
    Write-Host "`n[IDENTITY] Collecting role-based access..." -ForegroundColor Cyan
    try {
        $Roles = Get-GraphData "v1.0/directoryRoles?`$expand=members"
        $PrivilegedRoles = @("62e90394-69f5-423b-847e-217ed8c7fd18", "72fafb87-16d5-45d8-91ac-1dcd94a8a94e", "88d87e72-8544-4bbd-9b9e-06d0361df1aa")
        $Roles | ForEach-Object {
            $Members = $_.members | ForEach-Object { $_.userPrincipalName } | Where-Object { $_ }
            if ($Members.Count -gt 0) {
                $color = if ($PrivilegedRoles -contains $_.id) {"Red"} else {"Yellow"}
                Write-Host "  $($_.displayName) : $($Members.Count) members" -ForegroundColor $color
            }
        }
        Add-SessionLog "ACTION" "Role-based access collected" "Total: $($Roles.Count) roles"
        return $Roles
    }
    catch {
        Add-SessionLog "ERROR" "Get-RoleBasedAccess" $_.Exception.Message
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Export-IdentitySecurityExcel {
    <#
    .SYNOPSIS
        Consolidates all Identity & Security audit data into a single Excel workbook.

    .DESCRIPTION
        Calls Get-MFAStatus, Get-ConditionalAccessPolicies, Get-OAuthApplications,
        and Get-RoleBasedAccess, then exports all results to an Excel workbook
        with separate worksheets per category.
        TODO: Integrate with Export-FullAuditExcel for unified reporting.

    .PARAMETER SkipModuleInstall
        Skips automatic ImportExcel module installation.

    .EXAMPLE
        PS> Export-IdentitySecurityExcel
        Generates Identity_Security_Audit_<tenant>_<date>.xlsx
    #>
    param([switch]$SkipModuleInstall)

    $DateStr = Get-Date -Format "yyyyMMdd_HHmm"
    # M2: Collision-safe file path
    $XlsxPath = Get-UniqueFilePath -BasePath (Join-Path $GLOBAL:AUDIT_DIR "Sherl0ck_IdentitySecurity_$($GLOBAL:TENANT_NAME)_$DateStr.xlsx")

    try {
        Write-Host "`n[EXCEL] Generating Identity & Security workbook..." -ForegroundColor Cyan

        Write-Host " - MFA Status..." -ForegroundColor DarkGray
        $MFA = Get-MFAStatus -SkipModuleInstall:$SkipModuleInstall
        if ($MFA) { $MFA | Export-Excel -Path $XlsxPath -WorksheetName "MFA_Status" -AutoSize -AutoFilter }

        Write-Host " - Conditional Access Policies..." -ForegroundColor DarkGray
        $CA = Get-ConditionalAccessPolicies -SkipModuleInstall:$SkipModuleInstall
        if ($CA) { $CA | Select-Object displayName, state, createdDateTime | Export-Excel -Path $XlsxPath -WorksheetName "Conditional_Access" -AutoSize -AutoFilter -Append }

        Write-Host " - OAuth Applications..." -ForegroundColor DarkGray
        $Apps = Get-OAuthApplications -SkipModuleInstall:$SkipModuleInstall
        if ($Apps) { $Apps | Select-Object displayName, appId, signInAudience | Export-Excel -Path $XlsxPath -WorksheetName "OAuth_Apps" -AutoSize -AutoFilter -Append }

        Write-Host " - Role-Based Access..." -ForegroundColor DarkGray
        $Roles = Get-RoleBasedAccess -SkipModuleInstall:$SkipModuleInstall
        if ($Roles) { $Roles | Select-Object displayName, id | Export-Excel -Path $XlsxPath -WorksheetName "RBAC" -AutoSize -AutoFilter -Append }

        Write-Host "`n[SUCCESS] Identity audit complete: $XlsxPath" -ForegroundColor Green
        Add-SessionLog "ACTION" "Identity security export successful" $XlsxPath
        Invoke-SafeOpen -FilePath $XlsxPath
    }
    catch {
        Write-Host "`n[ERROR] Identity export: $($_.Exception.Message)" -ForegroundColor Red
        Add-SessionLog "ERROR" "Export-IdentitySecurityExcel" $_.Exception.Message
    }
}

Export-ModuleMember -Function Get-GraphData, Export-OneDriveUsage, Export-FullAuditExcel, Show-MenuAudit,
    Get-MFAStatus, Get-ConditionalAccessPolicies, Get-OAuthApplications,
    Get-RoleBasedAccess, Export-IdentitySecurityExcel
