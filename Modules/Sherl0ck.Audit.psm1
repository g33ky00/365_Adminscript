<#
.SYNOPSIS
    Sherl0ck Audit module — M365 data collection and export functions.

.DESCRIPTION
    Contains Get-GraphData (throttled Graph API pagination),
    Export-OneDriveUsage (per-user storage audit with M2 collision-safe CSV),
    and Export-FullAuditExcel (consolidated multi-worksheet workbook).

    M2: File paths use Get-UniqueFilePath to prevent overwrites.
    M3: ImportExcel installation uses MinimumVersion + SkipModuleInstall support.

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

Export-ModuleMember -Function Get-GraphData, Export-OneDriveUsage, Export-FullAuditExcel, Show-MenuAudit
