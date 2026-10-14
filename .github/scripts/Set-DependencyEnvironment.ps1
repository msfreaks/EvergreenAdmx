#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$dependencies = Import-PowerShellDataFile (Join-Path $PSScriptRoot '../powershell-dependencies.psd1')
$names = @{
    Pester = 'PESTER_VERSION'
    PSScriptAnalyzer = 'PSSA_VERSION'
    PowerShellGet = 'POWERSHELLGET_VERSION'
    'Microsoft.WinGet.Client' = 'WINGET_CLIENT_VERSION'
}
foreach ($name in $names.Keys) {
    $version = [version]$dependencies[$name]
    "$($names[$name])=$version" | Out-File -LiteralPath $env:GITHUB_ENV -Append -Encoding utf8
}
