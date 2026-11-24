<#
.SYNOPSIS
    Sherl0ck Exchange module — Exchange Online mailbox administration.

.DESCRIPTION
    Contains Show-MenuExchange (mailboxes, quotas, redirects, permissions, CSV export).
    Part of the 365_Adminscript modular architecture.
#>

function Show-MenuExchange {
    param([switch]$SkipModuleInstall)
    if (-not $GLOBAL:EXO_CONNECTED) {
        # caller should connect; keep simple
    }
    $QuitExo = $false
    do {
        Write-Host "`n--- [ EXCHANGE ONLINE ] ---" -ForegroundColor Cyan
        Write-Host " [1] Lister boites + tailles"
        Write-Host " [2] Modifier quota boite"
        Write-Host " [3] Gerer redirections"
        Write-Host " [4] Audit permissions deleguees"
        Write-Host " [5] Export CSV boites"
        Write-Host " [0] Retour"
        $Choice = Read-Host "Selection"
        if ($Choice -eq "0") { $QuitExo = $true; continue }
        if ($Choice -eq "1") {
            try {
                $Mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object DisplayName, UserPrincipalName, ProhibitSendReceiveQuota
                foreach ($mb in $Mailboxes) {
                    $stats = Get-MailboxStatistics -Identity $mb.UserPrincipalName -ErrorAction SilentlyContinue
                    $sizeGB = if ($stats) { Convert-EXOSizeToGB $stats.TotalItemSize } else { 0 }
                    Write-Host (" {0} | {1} Go | Quota: {2}" -f $mb.DisplayName, $sizeGB, $mb.ProhibitSendReceiveQuota) -ForegroundColor DarkGray
                }
            } catch { if (Get-Command Add-SessionLog -ErrorAction SilentlyContinue) { Add-SessionLog "ERREUR" "Liste boites" $_.Exception.Message } }
        }
        if ($Choice -eq "2") {
            $upn = Read-Host " UPN de la boite"
            $quota = Read-Host " Nouveau quota (ex: 50GB)"
            if ($upn -and $quota) {
                try {
                    Set-Mailbox -Identity $upn -ProhibitSendReceiveQuota $quota -ErrorAction Stop
                    Write-Host "[SUCCES] Quota mis a jour." -ForegroundColor Green
                    if (Get-Command Add-SessionLog -ErrorAction SilentlyContinue) { Add-SessionLog "ACTION" "Quota $upn -> $quota" }
                } catch { if (Get-Command Add-SessionLog -ErrorAction SilentlyContinue) { Add-SessionLog "ERREUR" "Modification quota" $_.Exception.Message } }
            }
        }
        if ($Choice -eq "3") {
            $upn = Read-Host " UPN de la boite"
            $fwd = Read-Host " Adresse de redirection (vide pour desactiver)"
            try {
                if ($fwd) { Set-Mailbox -Identity $upn -ForwardingSmtpAddress $fwd -DeliverToMailboxAndForward $true }
                else { Set-Mailbox -Identity $upn -ForwardingSmtpAddress $null }
                Write-Host "[SUCCES] Redirection mise a jour." -ForegroundColor Green
                if (Get-Command Add-SessionLog -ErrorAction SilentlyContinue) { Add-SessionLog "ACTION" "Redirection $upn -> $fwd" }
            } catch { if (Get-Command Add-SessionLog -ErrorAction SilentlyContinue) { Add-SessionLog "ERREUR" "Redirection" $_.Exception.Message } }
        }
        if ($Choice -eq "4") {
            $upn = Read-Host " UPN de la boite (vide = toutes)"
            try {
                $perms = if ($upn) { Get-MailboxPermission -Identity $upn } else { Get-Mailbox -ResultSize Unlimited | ForEach-Object { Get-MailboxPermission -Identity $_.UserPrincipalName } }
                $perms | Where-Object { $_.User -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false } | Select-Object Identity, User, AccessRights | Format-Table -AutoSize
            } catch { if (Get-Command Add-SessionLog -ErrorAction SilentlyContinue) { Add-SessionLog "ERREUR" "Permissions" $_.Exception.Message } }
        }
        if ($Choice -eq "5") {
            try {
                $Path = Join-Path $GLOBAL:AUDIT_DIR "Sherl0ck_Boites_$(Get-Date -Format yyyyMMdd_HHmm).csv"
                Get-Mailbox -ResultSize Unlimited | ForEach-Object {
                    $stats = Get-MailboxStatistics -Identity $_.UserPrincipalName -ErrorAction SilentlyContinue
                    [PSCustomObject]@{
                        DisplayName = $_.DisplayName
                        UPN = $_.UserPrincipalName
                        Quota = $_.ProhibitSendReceiveQuota
                        TailleGo = if ($stats) { Convert-EXOSizeToGB $stats.TotalItemSize } else { 0 }
                    }
                } | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8BOM
                Write-Host "[SUCCES] Export : $Path" -ForegroundColor Green
                Invoke-SafeOpen -FilePath $Path
            } catch { if (Get-Command Add-SessionLog -ErrorAction SilentlyContinue) { Add-SessionLog "ERREUR" "Export boites" $_.Exception.Message } }
        }
    } while (-not $QuitExo)
}

Export-ModuleMember -Function Show-MenuExchange