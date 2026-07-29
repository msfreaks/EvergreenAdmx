#Requires -Version 5.1
# Unit tests - no downloads, no admin required.

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

    It 'places PSScriptInfo before #Requires and comment-based help' {
        $header = Get-Content -LiteralPath $script:ScriptPath -Raw
        $psScriptInfo = $header.IndexOf('<#PSScriptInfo')
        $requires = $header.IndexOf('#Requires')
        $synopsis = $header.IndexOf('.SYNOPSIS')
        $psScriptInfo | Should -BeGreaterOrEqual 0
        $requires | Should -BeGreaterThan $psScriptInfo
        $synopsis | Should -BeGreaterThan $requires
    }

    It 'declares #Requires -Version 5.1 and #Requires -RunAsAdministrator' {
        $header = Get-Content -LiteralPath $script:ScriptPath -TotalCount 50
        ($header -join "`n") | Should -Match '#Requires\s+-Version\s+5\.1'
        ($header -join "`n") | Should -Match '#Requires\s+-RunAsAdministrator'
    }

    It 'exposes expected helper functions' {
        Get-Command -Name Get-WindowsDownloadId -CommandType Function | Should -Not -BeNullOrEmpty
        Get-Command -Name New-EvergreenAdmxTaskArgumentList -CommandType Function | Should -Not -BeNullOrEmpty
        Get-Command -Name Get-EvergreenAdmxObsoleteFilePatterns -CommandType Function | Should -Not -BeNullOrEmpty
        Get-Command -Name Clear-ObsoleteAdmx -CommandType Function | Should -Not -BeNullOrEmpty
        Get-Command -Name Initialize-PolicyStore -CommandType Function | Should -Not -BeNullOrEmpty
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
        $taskArgs = New-EvergreenAdmxTaskArgumentList -ScriptPath $script:FakeScript -BoundParameters @{}
        ($taskArgs -join ' ') | Should -Match '-NoProfile'
        ($taskArgs -join ' ') | Should -Match '-ExecutionPolicy Bypass'
        ($taskArgs -join ' ') | Should -Match ([regex]::Escape("-File `"$script:FakeScript`""))
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

    It 'forwards CleanPolicyStore switches' {
        $bound = [ordered]@{
            PolicyStore          = 'C:\PolicyDefinitions'
            CleanPolicyStore     = [System.Management.Automation.SwitchParameter]::new($true)
            CleanPolicyStoreOnly = [System.Management.Automation.SwitchParameter]::new($true)
            CreateScheduledTask  = [System.Management.Automation.SwitchParameter]::new($true)
        }

        $joined = (New-EvergreenAdmxTaskArgumentList -ScriptPath $script:FakeScript -BoundParameters $bound) -join ' '

        $joined | Should -Match '-CleanPolicyStore'
        $joined | Should -Match '-CleanPolicyStoreOnly'
        $joined | Should -Match '-PolicyStore "C:\\PolicyDefinitions"'
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
        $script:Products | Should -Contain 'Schannel'
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

Describe 'Get-EvergreenAdmxObsoleteFilePatterns' {
    It 'includes WinStoreUI, legacy Office, Adobe Classic, ctxprofile, and CitrixBase patterns' {
        $patterns = Get-EvergreenAdmxObsoleteFilePatterns
        $patterns | Should -Contain 'WinStoreUI.admx'
        $patterns | Should -Contain 'WinStoreUI.adml'
        $patterns | Should -Contain '*12*.admx'
        $patterns | Should -Contain '*15*.adml'
        $patterns | Should -Contain 'Acrobat2017.admx'
        $patterns | Should -Contain 'AcrobatReader2020.adml'
        $patterns | Should -Contain 'ctxprofile*.admx'
        $patterns | Should -Contain 'CitrixBase.admx'
        $patterns | Should -Contain 'CitrixBase.adml'
    }
}

Describe 'Initialize-PolicyStore' {
    BeforeEach {
        $script:StoreRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("EvergreenAdmx-Store-{0}" -f [guid]::NewGuid().ToString('N'))
    }

    AfterEach {
        if ($script:StoreRoot -and (Test-Path -LiteralPath $script:StoreRoot)) {
            Remove-Item -LiteralPath $script:StoreRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'creates the store and language folders when missing' {
        Initialize-PolicyStore -PolicyStore $script:StoreRoot -Languages @('en-US', 'fr-FR')
        Test-Path -LiteralPath $script:StoreRoot | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'en-US') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'fr-FR') | Should -BeTrue
    }

    It 'supports -WhatIf without creating folders' {
        Initialize-PolicyStore -PolicyStore $script:StoreRoot -Languages @('en-US') -WhatIf
        Test-Path -LiteralPath $script:StoreRoot | Should -BeFalse
    }
}

Describe 'Clear-ObsoleteAdmx' {
    BeforeEach {
        $script:StoreRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("EvergreenAdmx-Clean-{0}" -f [guid]::NewGuid().ToString('N'))
        $null = New-Item -Path $script:StoreRoot -ItemType Directory -Force
        $null = New-Item -Path (Join-Path $script:StoreRoot 'en-US') -ItemType Directory -Force
        $null = New-Item -Path (Join-Path $script:StoreRoot 'fr-FR') -ItemType Directory -Force

        # Keep
        Set-Content -Path (Join-Path $script:StoreRoot 'windows.admx') -Value 'keep'
        Set-Content -Path (Join-Path $script:StoreRoot 'en-US\windows.adml') -Value 'keep'
        Set-Content -Path (Join-Path $script:StoreRoot 'office16.admx') -Value 'keep'

        # Obsolete
        Set-Content -Path (Join-Path $script:StoreRoot 'WinStoreUI.admx') -Value 'remove'
        Set-Content -Path (Join-Path $script:StoreRoot 'en-US\WinStoreUI.adml') -Value 'remove'
        Set-Content -Path (Join-Path $script:StoreRoot 'excel15.admx') -Value 'remove'
        Set-Content -Path (Join-Path $script:StoreRoot 'en-US\excel15.adml') -Value 'remove'
        Set-Content -Path (Join-Path $script:StoreRoot 'Acrobat2017.admx') -Value 'remove'
        Set-Content -Path (Join-Path $script:StoreRoot 'AcrobatReader2020.admx') -Value 'remove'
        Set-Content -Path (Join-Path $script:StoreRoot 'ctxprofile7.admx') -Value 'remove'
        Set-Content -Path (Join-Path $script:StoreRoot 'CitrixBase.admx') -Value 'remove'
        Set-Content -Path (Join-Path $script:StoreRoot 'en-US\CitrixBase.adml') -Value 'remove'
        Set-Content -Path (Join-Path $script:StoreRoot 'readme.txt') -Value 'junk'
        $null = New-Item -Path (Join-Path $script:StoreRoot 'extract-debris') -ItemType Directory -Force
        Set-Content -Path (Join-Path $script:StoreRoot 'extract-debris\file.txt') -Value 'junk'
    }

    AfterEach {
        if ($script:StoreRoot -and (Test-Path -LiteralPath $script:StoreRoot)) {
            Remove-Item -LiteralPath $script:StoreRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'removes obsolete Admx/Adml, non-policy files, and non-language folders' {
        $removed = Clear-ObsoleteAdmx -PolicyStore $script:StoreRoot -Languages @('en-US', 'fr-FR')

        $removed.Count | Should -BeGreaterThan 0
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'WinStoreUI.admx') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'en-US\WinStoreUI.adml') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'excel15.admx') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'Acrobat2017.admx') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'AcrobatReader2020.admx') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'ctxprofile7.admx') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'CitrixBase.admx') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'en-US\CitrixBase.adml') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'readme.txt') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'extract-debris') | Should -BeFalse

        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'windows.admx') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'office16.admx') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'en-US') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'fr-FR') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'en-US\windows.adml') | Should -BeTrue
    }

    It 'supports -WhatIf without deleting files' {
        $null = Clear-ObsoleteAdmx -PolicyStore $script:StoreRoot -Languages @('en-US') -WhatIf

        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'WinStoreUI.admx') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'readme.txt') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'extract-debris') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:StoreRoot 'windows.admx') | Should -BeTrue
    }
}
