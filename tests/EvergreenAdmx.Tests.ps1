#Requires -Version 5.1
# Unit tests — no downloads, no admin required.

BeforeAll {
    . (Join-Path -Path $PSScriptRoot -ChildPath 'Helpers\Import-EvergreenAdmxUnderTest.ps1')
    # Dot-source function bodies in this scope (not inside Import-*).
    foreach ($text in (Get-EvergreenAdmxFunctionTexts)) {
        . ([scriptblock]::Create($text))
    }
    $script:ScriptPath = Get-EvergreenAdmxScriptPath
}

Describe 'EvergreenAdmx script surface' {
    It 'parses without syntax errors' {
        $tokens = $null
        $errors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$tokens, [ref]$errors)
        $errors | Should -BeNullOrEmpty
    }

    It 'declares #Requires -RunAsAdministrator' {
        $firstLines = Get-Content -LiteralPath $script:ScriptPath -TotalCount 5
        ($firstLines -join "`n") | Should -Match '#Requires\s+-RunAsAdministrator'
    }

    It 'exposes expected helper functions' {
        Get-Command -Name Get-WindowsDownloadId -CommandType Function | Should -Not -BeNullOrEmpty
        Get-Command -Name New-EvergreenAdmxTaskArgumentList -CommandType Function | Should -Not -BeNullOrEmpty
    }
}

Describe 'Language tag pattern' {
    It 'accepts <Language>' -ForEach @(
        @{ Language = 'en-US' }
        @{ Language = 'es' }
        @{ Language = 'fr-FR' }
        @{ Language = 'es-419' }
        @{ Language = 'nl-NL' }
    ) {
        $Language | Should -Match $script:EvergreenAdmxLanguagePattern
    }

    It 'rejects <Language>' -ForEach @(
        @{ Language = 'english' }
        @{ Language = 'en_US' }
        @{ Language = 'e' }
        @{ Language = 'en-US-x' }
    ) {
        $Language | Should -Not -Match $script:EvergreenAdmxLanguagePattern
    }
}

Describe 'Get-WindowsDownloadId' {
    It 'returns <Expected> for Windows <WindowsVersion> / <WindowsFeatureVersion>' -ForEach @(
        @{ WindowsVersion = 10; WindowsFeatureVersion = '21H2'; Expected = '104042' }
        @{ WindowsVersion = 10; WindowsFeatureVersion = '22H2'; Expected = '104677' }
        @{ WindowsVersion = 11; WindowsFeatureVersion = '23H2'; Expected = '105667' }
        @{ WindowsVersion = 11; WindowsFeatureVersion = '24H2'; Expected = '106254' }
        @{ WindowsVersion = 11; WindowsFeatureVersion = '25H2'; Expected = '108542' }
        @{ WindowsVersion = 2022; WindowsFeatureVersion = '25H2'; Expected = '104003' }
        @{ WindowsVersion = 2025; WindowsFeatureVersion = '25H2'; Expected = '108430' }
    ) {
        $id = Get-WindowsDownloadId -WindowsVersion $WindowsVersion -WindowsFeatureVersion $WindowsFeatureVersion
        if ($id -is [array]) { $id = $id[0] }
        "$id" | Should -Be $Expected
    }

    It 'rejects Windows 10 with 25H2' {
        { Get-WindowsDownloadId -WindowsVersion 10 -WindowsFeatureVersion '25H2' } |
            Should -Throw -ExpectedMessage '*Invalid Windows Feature Version*'
    }

    It 'rejects Windows 11 with 22H2' {
        { Get-WindowsDownloadId -WindowsVersion 11 -WindowsFeatureVersion '22H2' } |
            Should -Throw -ExpectedMessage '*Invalid Windows Feature Version*'
    }
}

Describe 'New-EvergreenAdmxTaskArgumentList' {
    BeforeAll {
        $script:FakeScript = 'D:\Tools\EvergreenAdmx\EvergreenAdmx.ps1'
    }

    It 'includes powershell host switches and -File path' {
        $args = New-EvergreenAdmxTaskArgumentList -ScriptPath $script:FakeScript -BoundParameters @{}
        ($args -join ' ') | Should -Match '-NoProfile'
        ($args -join ' ') | Should -Match '-ExecutionPolicy Bypass'
        ($args -join ' ') | Should -Match ([regex]::Escape("-File `"$script:FakeScript`""))
    }

    It 'forwards bound parameters and omits CreateScheduledTask' {
        $bound = [ordered]@{
            WorkingDirectory     = 'C:\Temp\EvergreenAdmx'
            Languages            = @('en-US', 'es', 'fr-FR')
            Include              = @('Microsoft Edge', 'Windows 11')
            UseProductFolders    = [System.Management.Automation.SwitchParameter]::new($true)
            CreateScheduledTask  = [System.Management.Automation.SwitchParameter]::new($true)
        }

        $joined = (New-EvergreenAdmxTaskArgumentList -ScriptPath $script:FakeScript -BoundParameters $bound) -join ' '

        $joined | Should -Match '-WorkingDirectory "C:\\Temp\\EvergreenAdmx"'
        $joined | Should -Match "-Languages @\('en-US','es','fr-FR'\)"
        $joined | Should -Match "-Include @\('Microsoft Edge','Windows 11'\)"
        $joined | Should -Match '-UseProductFolders'
        $joined | Should -Not -Match 'CreateScheduledTask'
    }

    It 'escapes single quotes inside array values' {
        $bound = @{
            Include = @("O'Reilly")
        }
        $joined = (New-EvergreenAdmxTaskArgumentList -ScriptPath $script:FakeScript -BoundParameters $bound) -join ' '
        $joined | Should -Match "-Include @\('O''Reilly'\)"
    }
}

Describe 'Include ValidateSet' {
    BeforeAll {
        $script:Products = Get-EvergreenAdmxIncludeValidateSet
    }

    It 'includes core products' {
        $script:Products | Should -Contain 'Windows 11'
        $script:Products | Should -Contain 'Microsoft Edge'
        $script:Products | Should -Contain 'HP Anyware'
        $script:Products | Should -Contain 'Custom Policy Store'
    }

    It 'includes all Windows SKUs' {
        @('Windows 10', 'Windows 11', 'Windows 2022', 'Windows 2025') | ForEach-Object {
            $script:Products | Should -Contain $_
        }
    }

    It 'has no duplicate product names' {
        $script:Products.Count | Should -Be ($script:Products | Select-Object -Unique).Count
    }
}
