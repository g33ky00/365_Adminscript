<#
.SYNOPSIS
    Sherl0ck Utils module — utility and parser functions.

.DESCRIPTION
    Contains helper functions for size conversion, file operations,
    and collision-safe file path generation (M2).

.NOTES
    Part of the 365_Adminscript modular architecture.
#>

# M2: Collision-safe file path generation
function Get-UniqueFilePath {
    param(
        [string]$BasePath
    )
    # M2: If the file exists, append a numeric suffix before the extension
    if (-not (Test-Path $BasePath)) { return $BasePath }
    $dir = Split-Path $BasePath -Parent
    $name = [System.IO.Path]::GetFileNameWithoutExtension($BasePath)
    $ext = [System.IO.Path]::GetExtension($BasePath)
    $counter = 1
    do {
        $newPath = Join-Path $dir "$name`_$counter$ext"
        $counter++
    } while (Test-Path $newPath)
    return $newPath
}

function Convert-EXOSizeToGB {
    param($SizeObj)
    if (-not $SizeObj) { return 0 }
    $str = $SizeObj.ToString()
    if ($str -match "\((?<bytes>\d+)\s*bytes\)") { return [math]::Round([double]$matches['bytes'] / 1GB, 2) }
    if ($str -match "(?<val>[\d\.,]+)\s*GB") { return [math]::Round([double]($matches['val'] -replace ',', '.'), 2) }
    if ($str -match "(?<val>[\d\.,]+)\s*MB") { return [math]::Round([double]($matches['val'] -replace ',', '.') / 1024, 2) }
    return 0
}

function Convert-BytesToGB {
    param([long]$Bytes)
    if ($Bytes -le 0) { return 0 }
    return [math]::Round($Bytes / 1GB, 2)
}

function Invoke-SafeOpen {
    param([string]$FilePath)
    if (Test-Path $FilePath) {
        Start-Sleep -Seconds 2
        Invoke-Item $FilePath
    }
}

Export-ModuleMember -Function Get-UniqueFilePath, Convert-EXOSizeToGB, Convert-BytesToGB, Invoke-SafeOpen
