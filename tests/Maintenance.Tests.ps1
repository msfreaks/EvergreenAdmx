#Requires -Version 5.1
BeforeAll {
    . (Join-Path $PSScriptRoot 'Helpers/Import-EvergreenAdmxUnderTest.ps1')
    foreach ($text in (Get-EvergreenAdmxFunctionText)) { . ([scriptblock]::Create($text)) }
    . (Join-Path $PSScriptRoot '../.github/scripts/Maintenance.ps1')
    $script:Source = [IO.File]::ReadAllText((Get-EvergreenAdmxScriptPath))
}

Describe 'Dependency update classification' {
    It 'classifies <Current> -> <Latest> as <Expected>' -ForEach @(
        @{ Current = '1.2.3'; Latest = '1.2.4'; Expected = 'routine' }
        @{ Current = '1.2.3'; Latest = '1.3.0'; Expected = 'routine' }
        @{ Current = '1.29.380'; Latest = '1.29.381.1'; Expected = 'routine' }
        @{ Current = '1.2.3'; Latest = '2.0.0'; Expected = 'major' }
        @{ Current = '1.2.3'; Latest = '1.2.3'; Expected = 'none' }
        @{ Current = '2.0.0'; Latest = '1.9.0'; Expected = 'none' }
    ) {
        Get-DependencyUpdateKind $Current $Latest | Should -Be $Expected
    }

    It 'rejects malformed dependency versions' {
        { Get-DependencyUpdateKind '1.0.0' 'latest' } | Should -Throw
    }

    It 'loads all dependency pins into the workflow environment' {
        $previousEnvironmentFile = $env:GITHUB_ENV
        $env:GITHUB_ENV = Join-Path $TestDrive 'environment.txt'
        try {
            & (Join-Path $PSScriptRoot '../.github/scripts/Set-DependencyEnvironment.ps1')
            $lines = Get-Content $env:GITHUB_ENV
            $dependencies = Import-PowerShellDataFile (Join-Path $PSScriptRoot '../.github/powershell-dependencies.psd1')
            $lines | Should -Contain "PESTER_VERSION=$($dependencies.Pester)"
            $lines | Should -Contain "PSSA_VERSION=$($dependencies.PSScriptAnalyzer)"
            $lines | Should -Contain "POWERSHELLGET_VERSION=$($dependencies.PowerShellGet)"
            $lines | Should -Contain "WINGET_CLIENT_VERSION=$($dependencies.'Microsoft.WinGet.Client')"
        } finally { $env:GITHUB_ENV = $previousEnvironmentFile }
    }
}

Describe 'Maintenance publication boundaries' {
    It 'permits only the dependency manifest for routine updates' {
        @(Get-MaintenanceFileList routine @('.github/powershell-dependencies.psd1')) | Should -Be @('.github/powershell-dependencies.psd1')
    }
    It 'rejects script changes disguised as routine dependency updates' {
        { Get-MaintenanceFileList routine @('EvergreenAdmx.ps1') } | Should -Throw '*unexpected file paths*'
    }
    It 'rejects traversal paths and incomplete family updates' {
        { Get-MaintenanceFileList families @('../README.md', 'EvergreenAdmx.ps1', 'CHANGELOG.md') } | Should -Throw
        { Get-MaintenanceFileList families @('EvergreenAdmx.ps1', 'EvergreenAdmx.ps1', 'CHANGELOG.md') } | Should -Throw
    }
    It 'rejects unknown update kinds' {
        { Get-MaintenanceFileList custom @('README.md') } | Should -Throw '*Unknown maintenance*'
    }
}

Describe 'Unchanged maintenance proposals' {
    BeforeEach {
        $folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = New-Item $folder -ItemType Directory
        Push-Location $folder
        git init -q -b main
        git config user.name 'Maintenance test'
        git config user.email 'maintenance@example.invalid'
        'baseline' | Set-Content dependency.txt
        git add .
        git -c commit.gpgsign=false commit -q -m baseline
        git checkout -q -b automation/major
        'update' | Set-Content dependency.txt
        git add .
        git -c commit.gpgsign=false commit -q -m proposal
        $script:ExistingHead = git rev-parse HEAD
        git checkout -q main
        'update' | Set-Content dependency.txt
        git add .
    }
    AfterEach { Pop-Location }

    It 'keeps an identical proposal on the same baseline' {
        Test-MaintenanceProposalUnchanged $script:ExistingHead | Should -BeTrue
    }
    It 'refreshes a proposal when its files change' {
        'newer update' | Set-Content dependency.txt
        git add .
        Test-MaintenanceProposalUnchanged $script:ExistingHead | Should -BeFalse
    }
    It 'refreshes a proposal when main advances' {
        git reset -q --hard HEAD
        'unrelated change' | Set-Content readme.txt
        git add .
        git -c commit.gpgsign=false commit -q -m advance
        'update' | Set-Content dependency.txt
        git add .
        Test-MaintenanceProposalUnchanged $script:ExistingHead | Should -BeFalse
    }
}

Describe 'Maintenance discovery orchestration' -Skip:($PSVersionTable.PSVersion -lt [version]'7.4') {
    BeforeEach {
        $script:PreviousRunnerTemp = $env:RUNNER_TEMP
        $script:PreviousOutput = $env:GITHUB_OUTPUT
        $env:RUNNER_TEMP = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = New-Item $env:RUNNER_TEMP -ItemType Directory
        $env:GITHUB_OUTPUT = Join-Path $env:RUNNER_TEMP 'outputs.txt'
        $script:DiscoveryRepo = Join-Path $env:RUNNER_TEMP 'repo'
        $scripts = Join-Path $script:DiscoveryRepo '.github/scripts'
        $null = New-Item $scripts -ItemType Directory -Force
        foreach ($file in @('Discover-Maintenance.ps1', 'Maintenance.ps1')) {
            Copy-Item (Join-Path $PSScriptRoot "../.github/scripts/$file") $scripts
        }
        @'
param([string]$Kind, [switch]$DryRun, [hashtable]$DependencyCache)
$manifest = Join-Path $PSScriptRoot '../powershell-dependencies.psd1'
if ((Get-Content $manifest -Raw).Trim() -ne 'baseline') { throw 'Discovery baseline was contaminated.' }
$case = Get-Content (Join-Path $env:RUNNER_TEMP 'case.txt') -Raw
if ($case.Trim() -eq 'failure' -and $Kind -eq 'families') { throw 'Vendor unavailable.' }
$changes = @()
$files = @()
if ($case.Trim() -ne 'empty' -and $Kind -ne 'families') {
    "updated $Kind" | Set-Content $manifest
    $changes = @("$Kind update")
    $files = @('.github/powershell-dependencies.psd1')
}
if (-not $DryRun) {
    @{ kind = $Kind; changes = $changes; files = $files } | ConvertTo-Json | Set-Content (Join-Path $PSScriptRoot '../../maintenance-result.json')
}
'@ | Set-Content (Join-Path $scripts 'Update-Maintenance.ps1')
        'baseline' | Set-Content (Join-Path $script:DiscoveryRepo '.github/powershell-dependencies.psd1')
        Push-Location $script:DiscoveryRepo
        git init -q -b main
        git config user.name 'Maintenance test'
        git config user.email 'maintenance@example.invalid'
        git add .
        git -c commit.gpgsign=false commit -q -m baseline
    }
    AfterEach {
        Pop-Location
        $env:RUNNER_TEMP = $script:PreviousRunnerTemp
        $env:GITHUB_OUTPUT = $script:PreviousOutput
    }

    It 'skips every publisher and artifact when nothing changes' {
        'empty' | Set-Content (Join-Path $env:RUNNER_TEMP 'case.txt')
        & ./.github/scripts/Discover-Maintenance.ps1
        Get-Content $env:GITHUB_OUTPUT | Should -Contain 'changed=false'
        Test-Path (Join-Path $env:RUNNER_TEMP 'maintenance-proposals') | Should -BeFalse
    }
    It 'packages only changed categories from separate clean baselines' {
        'changes' | Set-Content (Join-Path $env:RUNNER_TEMP 'case.txt')
        & ./.github/scripts/Discover-Maintenance.ps1
        $matrix = (Get-Content $env:GITHUB_OUTPUT | Where-Object { $_ -like 'matrix=*' }).Substring(7) | ConvertFrom-Json
        $matrix.kind | Should -Be @('routine', 'major')
        foreach ($kind in $matrix.kind) {
            (Get-Content (Join-Path $env:RUNNER_TEMP "maintenance-proposals/$kind/.github/powershell-dependencies.psd1") -Raw).Trim() | Should -Be "updated $kind"
        }
        (Get-Content .github/powershell-dependencies.psd1 -Raw).Trim() | Should -Be 'baseline'
    }
    It 'preserves dependency proposals when vendor discovery fails' {
        'failure' | Set-Content (Join-Path $env:RUNNER_TEMP 'case.txt')
        { & ./.github/scripts/Discover-Maintenance.ps1 -WarningAction SilentlyContinue } | Should -Throw '*Vendor unavailable*'
        Get-Content $env:GITHUB_OUTPUT | Should -Contain 'changed=true'
        $matrix = (Get-Content $env:GITHUB_OUTPUT | Where-Object { $_ -like 'matrix=*' }).Substring(7) | ConvertFrom-Json
        $matrix.kind | Should -Be @('routine', 'major')
    }
    It 'keeps dry runs free of proposal artifacts' {
        'changes' | Set-Content (Join-Path $env:RUNNER_TEMP 'case.txt')
        & ./.github/scripts/Discover-Maintenance.ps1 -DryRun
        Get-Content $env:GITHUB_OUTPUT | Should -Contain 'changed=false'
        Test-Path (Join-Path $env:RUNNER_TEMP 'maintenance-proposals') | Should -BeFalse
    }
    It 'queries each dependency once across routine and major discovery' {
        Mock Find-PSResource { [pscustomobject]@{ Version = [version]'999.0.0'; Prerelease = $false } }
        $cache = @{}
        $updater = Join-Path $PSScriptRoot '../.github/scripts/Update-Maintenance.ps1'
        & $updater -Kind routine -DryRun -DependencyCache $cache
        & $updater -Kind major -DryRun -DependencyCache $cache
        $dependencies = Import-PowerShellDataFile (Join-Path $PSScriptRoot '../.github/powershell-dependencies.psd1')
        Should -Invoke Find-PSResource -Times $dependencies.Count -Exactly
    }
}

Describe 'New Windows package validation' {
    BeforeEach {
        $script:PreviousTemp = $env:TEMP
        $env:TEMP = $TestDrive
    }
    AfterEach { $env:TEMP = $script:PreviousTemp }
    It 'refuses a package that extracts no templates and cleans up its working directory' {
        Mock Invoke-EvergreenAdmxWindows {}
        { Test-WindowsPolicyPackage @{ Version = '11'; Feature = '99H2' } } | Should -Throw '*no ADMX templates*'
        @(Get-ChildItem $TestDrive -Directory).Count | Should -Be 0
    }
    It 'validates paired XML from the extraction flow' {
        Mock Invoke-EvergreenAdmxWindows {
            $folder = Get-ChildItem $env:TEMP -Directory | Select-Object -First 1
            $language = Join-Path $folder.FullName 'admx/en-US'
            $null = New-Item $language -ItemType Directory -Force
            '<policyDefinitions revision="1.0" />' | Set-Content (Join-Path $folder.FullName 'admx/example.admx')
            '<policyDefinitionResources revision="1.0" />' | Set-Content (Join-Path $language 'example.adml')
        }
        { Test-WindowsPolicyPackage @{ Version = '11'; Feature = '99H2' } } | Should -Not -Throw
        @(Get-ChildItem $TestDrive -Directory).Count | Should -Be 0
    }
}

Describe 'Reviewed release catalog updates' {
    BeforeEach {
        # Synthetic baseline stays fixed while the runtime's reviewed releases advance.
        $script:Catalog = @{
            ABBYYMajor = 16
            FoxitMajor = 2026
            Windows = @(
                @{ Version = '10'; Feature = '22H2'; DownloadId = '104677'; Default = $true }
                @{ Version = '11'; Feature = '25H2'; DownloadId = '108542'; Default = $false }
                @{ Version = '11'; Feature = '26H2'; DownloadId = '108847'; Default = $true }
                @{ Version = '2025'; Feature = ''; DownloadId = '108430'; Default = $true }
            )
        }
    }

    It 'keeps exactly one default per Windows family' {
        foreach ($version in ($script:Catalog.Windows.Version | Select-Object -Unique)) {
            @($script:Catalog.Windows | Where-Object { $_.Version -eq $version -and $_.Default }).Count | Should -Be 1
        }
    }

    It 'does not downgrade an existing download or restore historical editions' {
        $before = $script:Catalog | ConvertTo-Json -Depth 5
        Add-NewWindowsRelease $script:Catalog @(
            @{ Version = '11'; Feature = '25H2'; DownloadId = '108394'; Default = $true }
            @{ Version = '11'; Feature = '22H2'; DownloadId = '100001'; Default = $true }
            @{ Version = '10'; Feature = '20H2'; DownloadId = '100002'; Default = $true }
        ) | Should -BeFalse
        ($script:Catalog | ConvertTo-Json -Depth 5) | Should -Be $before
    }

    It 'adds a future edition and retains explicit older selectors' {
        Add-NewWindowsRelease $script:Catalog @(@{ Version = '11'; Feature = '99H2'; DownloadId = '199999'; Default = $true }) | Should -BeTrue
        @($script:Catalog.Windows | Where-Object { $_.Version -eq '11' -and $_.Default }).Feature | Should -Be '99H2'
        @($script:Catalog.Windows | Where-Object { $_.Version -eq '11' -and $_.Feature -eq '26H2' }).DownloadId | Should -Be '108847'
        $updated = ConvertTo-UpdatedReleaseCatalogSource $script:Source $script:Catalog
        $updated | Should -Match "Feature = '99H2'"
        $updated | Should -Match "Feature = '26H2'"
    }

    It 'updates the quick-start default while keeping older release examples explicit' {
        Add-NewWindowsRelease $script:Catalog @(@{ Version = '12'; Feature = '99H2'; DownloadId = '199998'; Default = $true }) | Should -BeTrue
        $readme = "Defaults (Windows 11 26H2 plus Edge)`n`nWindows 11 25H2 instead of the default 26H2:"
        $updated = ConvertTo-UpdatedWindowsQuickStart $readme $script:Catalog
        $updated | Should -Match '^Defaults \(Windows 12 99H2 plus Edge\)'
        $updated | Should -Match 'Windows 11 25H2 instead of the default 26H2:'
    }

    It 'updates a newer package ID for an existing edition without changing its default' {
        Mock Get-WindowsReleaseCandidate { @{ PackageVersion = '3.0'; Published = [datetime]'2026-09-01' } }
        Add-NewWindowsRelease $script:Catalog @(@{ Version = '11'; Feature = '25H2'; DownloadId = '199999'; Default = $true; PackageVersion = '4.0'; Published = [datetime]'2026-10-01' }) | Should -BeTrue
        $edition = @($script:Catalog.Windows | Where-Object { $_.Version -eq '11' -and $_.Feature -eq '25H2' })[0]
        $edition.DownloadId | Should -Be '199999'
        $edition.Default | Should -BeFalse
        @($script:Catalog.Windows | Where-Object { $_.Version -eq '11' -and $_.Default }).Feature | Should -Be '26H2'
    }

    It 'rejects an older package version even if its page has a newer publication date' {
        Mock Get-WindowsReleaseCandidate { @{ PackageVersion = '3.0'; Published = [datetime]'2026-09-01' } }
        Add-NewWindowsRelease $script:Catalog @(@{ Version = '11'; Feature = '25H2'; DownloadId = '108394'; PackageVersion = '1.0'; Published = [datetime]'2026-10-01' }) | Should -BeFalse
        @($script:Catalog.Windows | Where-Object { $_.Version -eq '11' -and $_.Feature -eq '25H2' }).DownloadId | Should -Be '108542'
    }

    It 'updates only the matching Windows fixture when an approved download ID changes' {
        $baseline = @($script:Catalog.Windows | ForEach-Object { $_.Clone() })
        $edition = $script:Catalog.Windows | Where-Object { $_.Version -eq '11' -and $_.Feature -eq '25H2' }
        $edition.DownloadId = '199999'
        $fixture = "@{ WindowsVersion = 11; WindowsFeatureVersion = '25H2'; Expected = '108542' }`n@{ WindowsVersion = 11; WindowsFeatureVersion = '26H2'; Expected = '108847' }"
        $updated = ConvertTo-UpdatedWindowsTestSource $fixture $baseline $script:Catalog
        $updated | Should -Match "25H2'; Expected = '199999'"
        $updated | Should -Match "26H2'; Expected = '108847'"
    }

    It 'updates Server fixtures when they use an ignored feature version' {
        $baseline = @($script:Catalog.Windows | ForEach-Object { $_.Clone() })
        ($script:Catalog.Windows | Where-Object { $_.Version -eq '2025' }).DownloadId = '199999'
        $fixture = "@{ WindowsVersion = 2025; WindowsFeatureVersion = '25H2'; Expected = '108430' }"
        $updated = ConvertTo-UpdatedWindowsTestSource $fixture $baseline $script:Catalog
        $updated | Should -Match "WindowsVersion = 2025; WindowsFeatureVersion = '25H2'; Expected = '199999'"
    }

    It 'supports a newly reviewed client major and server year without new dispatch branches' {
        Add-NewWindowsRelease $script:Catalog @(
            @{ Version = '12'; Feature = '99H2'; DownloadId = '199998'; Default = $true }
            @{ Version = '2099'; Feature = ''; DownloadId = '199997'; Default = $true }
        ) | Should -BeTrue
        $updated = ConvertTo-UpdatedReleaseCatalogSource $script:Source $script:Catalog
        $updated | Should -Match '\$WindowsVersion = ''12'''
        $updated | Should -Match '\$WindowsVersion -lt 2000'
        $catalog = $script:Catalog
        Mock Get-EvergreenAdmxReleaseCatalog { $catalog }
        (Get-EvergreenAdmxProductCatalog).Name | Should -Contain 'Windows 12'
        (Get-EvergreenAdmxProductCatalog).Name | Should -Contain 'Windows 2099'
        Get-WindowsDownloadId -WindowsVersion 12 -WindowsFeatureVersion '99H2' | Should -Be '199998'
        Get-WindowsDownloadId -WindowsVersion 2099 | Should -Be '199997'
    }

    It 'rejects unknown Windows releases before selecting a package' {
        { Get-EvergreenAdmxWindowsRelease -WindowsVersion 99 } | Should -Throw '*Unsupported Windows version*'
    }

    It 'rejects source injection through downloaded catalog data' {
        $script:Catalog.Windows[0].DownloadId = "1'; Write-Host injected"
        { ConvertTo-UpdatedReleaseCatalogSource $script:Source $script:Catalog } | Should -Throw '*Invalid Windows release data*'
    }

    It 'generates documentation for every reviewed family' {
        $markdown = Get-ReleaseCatalogMarkdown $script:Catalog
        foreach ($release in $script:Catalog.Windows) { $markdown | Should -Match $release.DownloadId }
        $markdown | Should -Match "approved major: \*\*$($script:Catalog.ABBYYMajor)\*\*"
    }
}

Describe 'Microsoft release discovery' {
    It 'reads a single-component Windows package version from structured metadata' {
        Mock Invoke-WebRequest {
            $json = @{ dlcDetailsView = @{ downloadFile = @(@{ url = 'https://download.microsoft.com/example/admx.msi'; version = '1' }) } } | ConvertTo-Json -Depth 5 -Compress
            @{ Content = "<script>window.__DLCDetails__=$json</script>" }
        }
        $package = Get-EvergreenAdmxWindows -DownloadId 199999
        $package.Version | Should -Be '199999.1.0'
        { [version]$package.Version } | Should -Not -Throw
    }

    It 'does not propose preview Windows releases' {
        Mock Invoke-WebRequest {
            $json = @{ dlcDetailsView = @{ downloadTitle = 'Administrative Templates for Windows 11 (99H2) Insider Preview' } } | ConvertTo-Json -Depth 5 -Compress
            @{ Content = "<script>window.__DLCDetails__=$json</script>" }
        }
        @(Get-WindowsReleaseCandidate '<a href="https://www.microsoft.com/download/details.aspx?id=199999">ADMX</a>').Count | Should -Be 0
    }

    It 'accepts an official ADMX MSI and ignores duplicate links' {
        Mock Invoke-WebRequest {
            $json = @{ dlcDetailsView = @{
                downloadTitle = 'Administrative Templates (.admx) for Windows 11 2099 Update (99H2)'
                downloadFile = @(@{ url = 'https://download.microsoft.com/example/admx.msi'; version = '1'; datePublished = '9/29/2099' })
            } } | ConvertTo-Json -Depth 5 -Compress
            @{ Content = "<script>window.__DLCDetails__=$json</script>" }
        }
        $index = '<a href="https://www.microsoft.com/en-us/download/details.aspx?id=199999">ADMX</a>'
        $releases = @(Get-WindowsReleaseCandidate "$index$index")
        $releases.Count | Should -Be 1
        $releases[0].Feature | Should -Be '99H2'
        $releases[0].DownloadId | Should -Be '199999'
    }

    It 'rejects packages hosted outside Microsoft' {
        Mock Invoke-WebRequest {
            $json = @{ dlcDetailsView = @{
                downloadTitle = 'Administrative Templates for Windows 11 (99H2)'
                downloadFile = @(@{ url = 'https://example.com/admx.msi'; version = '1.0'; datePublished = '9/29/2099' })
            } } | ConvertTo-Json -Depth 5 -Compress
            @{ Content = "<script>window.__DLCDetails__=$json</script>" }
        }
        { Get-WindowsReleaseCandidate '<a href="https://www.microsoft.com/en-us/download/details.aspx?id=199999">ADMX</a>' } | Should -Throw '*official MSI*'
    }

    It 'fails visibly when the index no longer exposes packages' {
        { Get-WindowsReleaseCandidate '<html>changed</html>' } | Should -Throw '*no Download Center links*'
    }
}

Describe 'ABBYY discovery' {
    It 'discovers numeric Windows families and excludes Mac guides' {
        Get-ABBYYReleaseCandidate '/en-us/finereader/16/admin_guide/ /en-us/finereader/17/admin_guide/ /en-us/finereader/99mac/admin_guide/' | Should -Be 17
    }

    It 'fails visibly when the help index changes' {
        { Get-ABBYYReleaseCandidate '<html>changed</html>' } | Should -Throw '*no Windows FineReader*'
    }

    It 'uses the reviewed major and fresh attachment URLs' {
        Mock Get-EvergreenAdmxReleaseCatalog { @{ ABBYYMajor = 17 } }
        Mock Invoke-WebRequest { @{ Content = 'https://support.abbyy.com/hc/en-us/article_attachments/123/FineReader17.admx https://support.abbyy.com/hc/en-us/article_attachments/456/FineReader17.adml' } }
        Mock Resolve-Uri { @{ LastModified = [datetime]'2026-09-24' } }
        $package = Get-EvergreenAdmxABBYYFineReader
        $package.Version | Should -Be '17.2026.9.24'
        $package.URI | Should -Match '/123/FineReader17.admx$'
        Should -Invoke Invoke-WebRequest -ParameterFilter { $Uri -match '/finereader/17/' } -Times 1 -Exactly
    }

    It 'rejects a mismatched ADMX and ADML family instead of using stale URLs' {
        Mock Invoke-WebRequest { @{ Content = 'https://support.abbyy.com/hc/en-us/article_attachments/123/FineReader16.admx https://support.abbyy.com/hc/en-us/article_attachments/456/FineReader15.adml' } }
        { Get-EvergreenAdmxABBYYFineReader } | Should -Throw '*matching ADMX and ADML*'
    }

    It 'rejects ADMX references absent from the ADML' {
        Mock Invoke-WebRequest {
            if ($Uri -like '*.admx') { @{ Content = '<policyDefinitions><policies><policy displayName="$(string.missing)" /></policies></policyDefinitions>' } }
            else { @{ Content = '<policyDefinitionResources><resources><stringTable><string id="present">Present</string></stringTable></resources></policyDefinitionResources>' } }
        }
        { Test-ABBYYPolicyPair @{ URI = 'https://example.com/file.admx'; AdmlURI = 'https://example.com/file.adml' } } | Should -Throw '*Missing ABBYY resource*'
    }
}

Describe 'ABBYY version tracking' {
    BeforeEach {
        $script:PreviousTemp = $env:TEMP
        $env:TEMP = $TestDrive
        $script:WorkingDirectory = Join-Path $TestDrive 'work'
        $null = New-Item (Join-Path $script:WorkingDirectory 'downloads') -ItemType Directory -Force
        Mock Copy-Admx {}
        Mock Invoke-FileDownload { '<xml />' | Set-Content -LiteralPath $OutFile }
    }
    AfterEach { $env:TEMP = $script:PreviousTemp }

    It 'migrates an old date-only record even when the attachment date is older' {
        Mock Get-EvergreenAdmxABBYYFineReader { @{ Version = '16.2024.7.28'; URI = 'https://example.com/FineReader16.admx'; AdmlURI = 'https://example.com/FineReader16.adml' } }
        (Invoke-EvergreenAdmxABBYYFineReader -Version '2026.9.24' -Languages @('en-US')).Version | Should -Be '16.2024.7.28'
        Should -Invoke Invoke-FileDownload -Times 2 -Exactly
    }

    It 'downloads a newer major despite older attachment timestamps' {
        Mock Get-EvergreenAdmxABBYYFineReader { @{ Version = '17.2023.1.1'; URI = 'https://example.com/FineReader17.admx'; AdmlURI = 'https://example.com/FineReader17.adml' } }
        (Invoke-EvergreenAdmxABBYYFineReader -Version '16.2024.7.28' -Languages @('en-US')).Version | Should -Be '17.2023.1.1'
        Should -Invoke Invoke-FileDownload -ParameterFilter { $OutFile -like '*FineReader17.*' } -Times 2 -Exactly
    }

    It 'skips an unchanged major and date' {
        Mock Get-EvergreenAdmxABBYYFineReader { @{ Version = '16.2024.7.28' } }
        Invoke-EvergreenAdmxABBYYFineReader -Version '16.2024.7.28' | Should -BeNullOrEmpty
        Should -Invoke Invoke-FileDownload -Times 0 -Exactly
    }
}

Describe 'Foxit package discovery' {
    BeforeEach {
        Mock Invoke-WebRequest {
            if ($Uri -like 'https://www.foxit.com/*') { return @{ Content = 'Version 2026.2.1.39815 Version 2026.2.1.39815 Version 2026.1.0.36452' } }
            if ($Uri -like '*/2026.2.1/*') { throw '404 unpublished templates' }
            return @{ StatusCode = 206; Content = [byte[]]@(80, 75, 3, 4) }
        }
    }

    It 'uses the latest published paired templates with ranged GET rather than HEAD' {
        $package = Get-EvergreenAdmxFoxit
        $package.Version | Should -Be '2026.1.0'
        $package.ReaderURI | Should -Match 'Reader'
        Should -Invoke Invoke-WebRequest -ParameterFilter { $Method -eq 'Head' } -Times 0 -Exactly
        Should -Invoke Invoke-WebRequest -ParameterFilter { $Method -eq 'Get' -and $Headers.Range -eq 'bytes=0-3' } -Times 3 -Exactly
    }

    It 'rejects a release whose Reader templates are missing' {
        Mock Invoke-WebRequest { throw '404 missing Reader' } -ParameterFilter { $Uri -like '*Reader*' }
        { Get-EvergreenAdmxFoxit } | Should -Throw '*Unable to locate*'
    }

    It 'fails visibly when version metadata disappears' {
        Mock Invoke-WebRequest { @{ Content = '<html>changed</html>' } } -ParameterFilter { $Uri -like 'https://www.foxit.com/*' }
        { Get-EvergreenAdmxFoxit } | Should -Throw '*no recognizable versions*'
    }

    It 'rejects a CDN error page even when its status is successful' {
        Mock Invoke-WebRequest { @{ StatusCode = 200; Content = '<html>access denied</html>' } } -ParameterFilter { $Method -eq 'Get' }
        { Get-EvergreenAdmxFoxit } | Should -Throw '*Unable to locate*'
    }

    It 'keeps unreviewed annual families out of runtime downloads while discovery can inspect them' {
        Mock Invoke-WebRequest { @{ Content = 'Version 2099.1.0.12345 Version 2026.1.0.36452' } } -ParameterFilter { $Uri -like 'https://www.foxit.com/*' }
        (Get-EvergreenAdmxFoxit).Version | Should -Be '2026.1.0'
        (Get-EvergreenAdmxFoxit -MaximumMajor ([int]::MaxValue)).Version | Should -Be '2099.1.0'
    }
}
