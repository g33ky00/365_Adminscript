<#
.SYNOPSIS
    Pester test stubs for Sherl0ck.Utils module.

.DESCRIPTION
    Basic unit tests for utility functions.
    Run with: Invoke-Pester -Path ./Tests/Sherl0ck.Tests.Utils.ps1
#>

# Requires Pester module
# Install-Module Pester -Scope CurrentUser -Force

Describe 'Convert-BytesToGB' {
    It 'Should return 0 for zero or negative bytes' {
        Convert-BytesToGB -Bytes 0 | Should -Be 0
        Convert-BytesToGB -Bytes -100 | Should -Be 0
    }

    It 'Should convert bytes to GB correctly' {
        $result = Convert-BytesToGB -Bytes 1073741824  # 1 GB
        $result | Should -Be 1.0
    }
}

Describe 'Convert-EXOSizeToGB' {
    It 'Should return 0 for null input' {
        Convert-EXOSizeToGB -SizeObj $null | Should -Be 0
    }

    It 'Should parse GB strings correctly' {
        $result = Convert-EXOSizeToGB -SizeObj "5.5 GB"
        $result | Should -Be 5.5
    }

    It 'Should parse MB strings correctly' {
        $result = Convert-EXOSizeToGB -SizeObj "1024 MB"
        $result | Should -Be 1.0
    }

    It 'Should parse bytes from Exchange format' {
        $result = Convert-EXOSizeToGB -SizeObj "(1073741824 bytes)"
        $result | Should -Be 1.0
    }
}

Describe 'Get-UniqueFilePath' {
    It 'Should return the original path if file does not exist' {
        $testPath = "C:\Temp\NonExistentFile_$(Get-Random).txt"
        $result = Get-UniqueFilePath -BasePath $testPath
        $result | Should -Be $testPath
    }

    It 'Should append counter if file exists' {
        $testPath = "TestDrive:\ExistingFile.txt"
        $null = New-Item -Path $testPath -ItemType File -Force
        $result = Get-UniqueFilePath -BasePath $testPath
        $result | Should -Not -Be $testPath
        $result | Should -Match '_1.txt$'
    }
}

Describe 'Invoke-SafeOpen' {
    It 'Should not throw if file does not exist' {
        { Invoke-SafeOpen -FilePath "C:\NonExistent\Path.txt" } | Should -Not -Throw
    }
}
