<#
.SYNOPSIS
    Pester test stubs for Sherl0ck.Audit module.

.DESCRIPTION
    Basic unit tests for audit collection and export functions.
    Run with: Invoke-Pester -Path ./Tests/Sherl0ck.Tests.Audit.ps1
#>

Describe 'Get-GraphData' {
    It 'Should return empty array for failed query (non-connected)' {
        # Requires an active Graph connection to test properly
        # This stub verifies the function exists and doesn't throw without a connection
        $result = Get-GraphData -Uri "v1.0/tenant" -ErrorAction SilentlyContinue
        $result | Should -Not -Be $null
    }
}

Describe 'Export-OneDriveUsage' {
    It 'Should handle missing Graph connection gracefully' {
        $GLOBAL:GRAPH_CONNECTED = $false
        { Export-OneDriveUsage -SkipModuleInstall -ErrorAction SilentlyContinue } | Should -Not -Throw
    }
}

Describe 'Export-FullAuditExcel' {
    It 'Should return gracefully when ImportExcel is not available' {
        $GLOBAL:GRAPH_CONNECTED = $false
        $GLOBAL:EXO_CONNECTED = $false
        $GLOBAL:AUDIT_DIR = "TestDrive:\Audits"
        $GLOBAL:TENANT_NAME = "TestTenant"
        { Export-FullAuditExcel -SkipModuleInstall -ErrorAction SilentlyContinue } | Should -Not -Throw
    }
}

Describe 'Show-MenuAudit' {
    It 'Should return when not connected to Graph' {
        $GLOBAL:GRAPH_CONNECTED = $false
        { Show-MenuAudit -AuditMode 'ReadOnly' -SkipModuleInstall -ErrorAction SilentlyContinue } | Should -Not -Throw
    }
}
