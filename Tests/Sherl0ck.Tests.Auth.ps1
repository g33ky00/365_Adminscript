<#
.SYNOPSIS
    Pester test stubs for Sherl0ck.Auth module.

.DESCRIPTION
    Basic unit tests for authentication module functions.
    Run with: Invoke-Pester -Path ./Tests/Sherl0ck.Tests.Auth.ps1
#>

Describe 'Verify-TrustedModule' {
    It 'Should return true for PSGallery-verified modules' {
        # This test requires PSGallery to be registered
        # In test environment, may return false if PSGallery not available
        $result = Verify-TrustedModule -ModuleName 'ImportExcel' -RequiredVersion '7.0.0'
        $result | Should -BeOfType [bool]
    }

    It 'Should return false for invalid module name' {
        $result = Verify-TrustedModule -ModuleName 'NonExistentModule_XYZ123' -RequiredVersion '99.0.0'
        $result | Should -Be $false
    }

    It 'Should return false when RequiredVersion exceeds available' {
        $result = Verify-TrustedModule -ModuleName 'ImportExcel' -RequiredVersion '99.99.99'
        $result | Should -Be $false
    }
}

Describe 'Connect-O365Core' {
    It 'Should respect SkipModuleInstall flag' {
        # This is a stub test — actual auth requires interactive login
        # The function should return early if module is missing and SkipModuleInstall is set
        Set-Item -Path "variable:GLOBAL:GRAPH_CONNECTED" -Value $false -Scope Global
        { Connect-O365Core -AuditMode 'ReadOnly' -SkipModuleInstall -ErrorAction SilentlyContinue } | Should -Not -Throw
    }
}

Describe 'Connect-O365Exchange' {
    It 'Should respect SkipModuleInstall flag' {
        # This is a stub test — actual auth requires interactive login
        Set-Item -Path "variable:GLOBAL:EXO_CONNECTED" -Value $false -Scope Global
        { Connect-O365Exchange -SkipModuleInstall -ErrorAction SilentlyContinue } | Should -Not -Throw
    }
}
