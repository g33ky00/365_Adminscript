<#
.SYNOPSIS
    Pester test stubs for Sherl0ck.UI module.

.DESCRIPTION
    Basic unit tests for logging and sensitive data masking functions.
    Run with: Invoke-Pester -Path ./Tests/Sherl0ck.Tests.UI.ps1
#>

Describe 'Mask-SensitiveData' {
    It 'Should mask UPNs (email addresses)' {
        $result = Mask-SensitiveData -InputText "admin@contoso.com"
        $result | Should -Be "***"
    }

    It 'Should mask UPNs in longer text' {
        $result = Mask-SensitiveData -InputText "user logged in as admin@contoso.com via Graph"
        $result | Should -Not -Contain "@"
    }

    It 'Should return null/empty for null input' {
        $result = Mask-SensitiveData -InputText $null
        $result | Should -Be $null
    }

    It 'Should leave non-email text unchanged' {
        $result = Mask-SensitiveData -InputText "No sensitive data here"
        $result | Should -Be "No sensitive data here"
    }
}

Describe 'Add-SessionLog' {
    BeforeAll {
        $GLOBAL:SESSION_LOGS = @()
        $GLOBAL:LOG_FILE = "TestDrive:\Session.log"
    }

    It 'Should add entry to session logs' {
        Add-SessionLog -Level "INFO" -Message "Test message"
        $GLOBAL:SESSION_LOGS.Count | Should -BeGreaterThan 0
    }

    It 'Should mask UPNs in details' {
        Add-SessionLog -Level "INFO" -Message "Connection test" -Details "admin@contoso.com"
        $lastEntry = $GLOBAL:SESSION_LOGS[-1]
        $lastEntry | Should -Not -Contain "admin@contoso.com"
    }
}

Describe 'Show-SessionLogs' {
    It 'Should not throw when no logs exist' {
        $GLOBAL:SESSION_LOGS = @()
        $GLOBAL:LOG_DIR = "TestDrive:\Logs"
        { Show-SessionLogs } | Should -Not -Throw
    }
}
