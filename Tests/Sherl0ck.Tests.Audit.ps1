<#
.SYNOPSIS
    Pester test stubs for Sherl0ck.Audit module — Identity & Security functions.

.DESCRIPTION
    Basic unit tests for MFA status, Conditional Access, OAuth applications, and RBAC collection.
    Run with: Invoke-Pester -Path ./Tests/Sherl0ck.Tests.Audit.ps1
#>

# Import the audit module before testing
# Import-Module ../Modules/Sherl0ck.Audit.psm1

Describe 'Get-MFAStatus' {
    It 'Should return $null when not connected to Graph' {
        $GLOBAL:GRAPH_CONNECTED = $false
        $result = Get-MFAStatus -SkipModuleInstall -ErrorAction SilentlyContinue
        $result | Should -BeNullOrEmpty
    }
}

Describe 'Get-ConditionalAccessPolicies' {
    It 'Should handle missing Graph connection gracefully' {
        $GLOBAL:GRAPH_CONNECTED = $false
        { Get-ConditionalAccessPolicies -SkipModuleInstall -ErrorAction SilentlyContinue } | Should -Not -Throw
    }
}

Describe 'Get-OAuthApplications' {
    It 'Should handle missing Graph connection gracefully' {
        $GLOBAL:GRAPH_CONNECTED = $false
        { Get-OAuthApplications -SkipModuleInstall -ErrorAction SilentlyContinue } | Should -Not -Throw
    }
}

Describe 'Get-RoleBasedAccess' {
    It 'Should handle missing Graph connection gracefully' {
        $GLOBAL:GRAPH_CONNECTED = $false
        { Get-RoleBasedAccess -SkipModuleInstall -ErrorAction SilentlyContinue } | Should -Not -Throw
    }
}

Describe 'Export-IdentitySecurityExcel' {
    It 'Should return gracefully when Export-OneDriveUsage fails' {
        $GLOBAL:GRAPH_CONNECTED = $false
        $GLOBAL:EXO_CONNECTED = $false
        $GLOBAL:AUDIT_DIR = "TestDrive:\Audits"
        $GLOBAL:TENANT_NAME = "TestTenant"
        { Export-IdentitySecurityExcel -SkipModuleInstall -ErrorAction SilentlyContinue } | Should -Not -Throw
    }
}

Describe 'Export-OneDriveUsage (Point 4: $batch)' {
    It 'Should handle missing Graph connection gracefully' {
        $GLOBAL:GRAPH_CONNECTED = $false
        $GLOBAL:AUDIT_DIR = "TestDrive:\Audits"
        $GLOBAL:TENANT_NAME = "TestTenant"
        { Export-OneDriveUsage -SkipModuleInstall -ErrorAction SilentlyContinue } | Should -Not -Throw
    }

    It 'Should use $batch endpoint (POST /v1.0/$batch)' {
        # This is a structural check — verify the function contains the batch endpoint
        $auditModuleContent = Get-Content (Join-Path $PSScriptRoot "..\Modules\Sherl0ck.Audit.psm1") -Raw
        $auditModuleContent | Should -Contain '$batch'
        $auditModuleContent | Should -Contain 'v1.0/$batch'
        $auditModuleContent | Should -Contain 'BatchSize = 20'
    }
}
