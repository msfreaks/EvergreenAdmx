#Requires -Version 5.1
# Integration smoke — small real download. Tag: Integration
# Run elevated (EvergreenAdmx.ps1 requires Administrator).

BeforeAll {
    . (Join-Path -Path $PSScriptRoot -ChildPath 'Helpers\Import-EvergreenAdmxUnderTest.ps1')
    if (-not (Test-EvergreenAdmxIsAdministrator)) {
        throw 'Integration tests require an elevated PowerShell session.'
    }
    $script:ScriptPath = Get-EvergreenAdmxScriptPath
    $script:WorkRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("EvergreenAdmx-Smoke-{0}" -f [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $script:WorkRoot -Force | Out-Null
}

AfterAll {
    if ($script:WorkRoot -and (Test-Path -LiteralPath $script:WorkRoot)) {
        Remove-Item -LiteralPath $script:WorkRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Release smoke download' -Tag 'Integration' {
    It 'downloads Microsoft Edge for en-US, es, and fr-FR' {
        $null = & $script:ScriptPath `
            -WorkingDirectory $script:WorkRoot `
            -Languages @('en-US', 'es', 'fr-FR') `
            -Include @('Microsoft Edge')

        $admxRoot = Join-Path -Path $script:WorkRoot -ChildPath 'admx'
        Test-Path -LiteralPath $admxRoot | Should -BeTrue

        $edgeAdmx = Get-ChildItem -LiteralPath $admxRoot -Filter 'msedge.admx' -File -ErrorAction SilentlyContinue
        $edgeAdmx | Should -Not -BeNullOrEmpty

        $enUs = Join-Path -Path $admxRoot -ChildPath 'en-US'
        Test-Path -LiteralPath $enUs | Should -BeTrue
        (Get-ChildItem -LiteralPath $enUs -Filter '*.adml' -File).Count | Should -BeGreaterThan 0

        # es may fall back to en-US; fr-FR is commonly present for Edge
        $fr = Join-Path -Path $admxRoot -ChildPath 'fr-FR'
        if (Test-Path -LiteralPath $fr) {
            (Get-ChildItem -LiteralPath $fr -Filter '*.adml' -File).Count | Should -BeGreaterThan 0
        }

        $versions = Join-Path -Path $script:WorkRoot -ChildPath 'AdmxVersions.xml'
        Test-Path -LiteralPath $versions | Should -BeTrue
    }
}

Describe 'CreateScheduledTask registration' -Tag 'Integration' {
    BeforeAll {
        $script:TaskName = 'EvergreenAdmx'
        $script:TaskWork = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("EvergreenAdmx-Task-{0}" -f [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $script:TaskWork -Force | Out-Null
    }

    AfterAll {
        Unregister-ScheduledTask -TaskName $script:TaskName -Confirm:$false -ErrorAction SilentlyContinue
        if ($script:TaskWork -and (Test-Path -LiteralPath $script:TaskWork)) {
            Remove-Item -LiteralPath $script:TaskWork -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'registers a weekly SYSTEM task and forwards parameters' {
        $null = & $script:ScriptPath `
            -WorkingDirectory $script:TaskWork `
            -Languages @('en-US', 'es', 'fr-FR') `
            -Include @('Microsoft Edge') `
            -CreateScheduledTask

        $task = Get-ScheduledTask -TaskName $script:TaskName
        $task | Should -Not -BeNullOrEmpty
        $task.Principal.UserId | Should -Match 'SYSTEM'
        $task.Principal.RunLevel | Should -Be 'Highest'

        $taskArgs = $task.Actions.Arguments
        $taskArgs | Should -Match 'en-US'
        $taskArgs | Should -Match 'fr-FR'
        $taskArgs | Should -Match 'Microsoft Edge'
        $taskArgs | Should -Not -Match 'CreateScheduledTask'

        $trigger = $task.Triggers | Select-Object -First 1
        $trigger.CimClass.CimClassName | Should -Be 'MSFT_TaskWeeklyTrigger'
    }
}
