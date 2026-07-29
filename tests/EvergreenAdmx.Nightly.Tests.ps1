#Requires -Version 5.1
# Full-product health check. Tag: Nightly
# Excludes Custom Policy Store (needs path) and Windows 10 (incompatible with default Win11 25H2).
# Run elevated (EvergreenAdmx.ps1 requires Administrator).

BeforeAll {
    . (Join-Path -Path $PSScriptRoot -ChildPath 'Helpers\Import-EvergreenAdmxUnderTest.ps1')
    if (-not (Test-EvergreenAdmxIsAdministrator)) {
        throw 'Nightly tests require an elevated PowerShell session.'
    }
    $script:ScriptPath = Get-EvergreenAdmxScriptPath
    $script:Products = @(Get-EvergreenAdmxIncludeValidateSet | Where-Object {
            $_ -notin @('Custom Policy Store', 'Windows 10', 'BIS-F')
        })

    # BIS-F is excluded by default in CI because unauthenticated GitHub zipball downloads often 403.
    # Set EVERGREENADMX_INCLUDE_BISF=1 to include it.
    if ($env:EVERGREENADMX_INCLUDE_BISF -eq '1') {
        $script:Products += 'BIS-F'
    }

    $script:WorkRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("EvergreenAdmx-Nightly-{0}" -f [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $script:WorkRoot -Force | Out-Null
}

AfterAll {
    if ($script:WorkRoot -and (Test-Path -LiteralPath $script:WorkRoot)) {
        Remove-Item -LiteralPath $script:WorkRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Full product matrix' -Tag 'Nightly' {
    It 'processes all products for en-US, es, and fr-FR' {
        $script:Products.Count | Should -BeGreaterThan 20

        & $script:ScriptPath `
            -WorkingDirectory $script:WorkRoot `
            -Languages @('en-US', 'es', 'fr-FR') `
            -Include $script:Products

        $admxRoot = Join-Path -Path $script:WorkRoot -ChildPath 'admx'
        Test-Path -LiteralPath $admxRoot | Should -BeTrue
        (Get-ChildItem -LiteralPath $admxRoot -Filter '*.admx' -File).Count | Should -BeGreaterThan 50

        $enUs = Join-Path -Path $admxRoot -ChildPath 'en-US'
        Test-Path -LiteralPath $enUs | Should -BeTrue
        (Get-ChildItem -LiteralPath $enUs -Filter '*.adml' -File).Count | Should -BeGreaterThan 50

        $versionsPath = Join-Path -Path $script:WorkRoot -ChildPath 'AdmxVersions.xml'
        Test-Path -LiteralPath $versionsPath | Should -BeTrue
        $versions = Import-Clixml -LiteralPath $versionsPath
        @($versions.Keys).Count | Should -BeGreaterThan 20
    }
}
