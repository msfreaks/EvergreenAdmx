#Requires -Version 7.4
[CmdletBinding()]
param(
    [ValidateSet('routine', 'major', 'families')][string]$Kind,
    [switch]$DryRun,
    [hashtable]$DependencyCache = @{}
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Maintenance.ps1')
$root = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$changes = [System.Collections.Generic.List[string]]::new()
$files = @()
if ($Kind -ne 'families') {
    $manifest = Join-Path $root '.github/powershell-dependencies.psd1'
    $dependencies = Import-PowerShellDataFile $manifest
    foreach ($name in ($dependencies.Keys | Sort-Object)) {
        if (-not $DependencyCache.ContainsKey($name)) {
            $DependencyCache[$name] = Find-PSResource -Name $name -Repository PSGallery -ErrorAction Stop
        }
        $latest = $DependencyCache[$name]
        if (-not $latest -or $latest.Prerelease) { throw "No stable Gallery release found for $name." }
        $version = $latest.Version.ToString()
        if ((Get-DependencyUpdateKind $dependencies[$name] $version) -eq $Kind) {
            $changes.Add("$name`: $($dependencies[$name]) -> $version")
            $dependencies[$name] = $version
        }
    }
    if ($changes.Count) {
        $rows = $dependencies.Keys | Sort-Object | ForEach-Object { "    '$_' = '$($dependencies[$_])'" }
        if (-not $DryRun) { [IO.File]::WriteAllText($manifest, "@{`n$($rows -join "`n")`n}`n") }
        $files = @('.github/powershell-dependencies.psd1')
    }
} else {
    $scriptPath = Join-Path $root 'EvergreenAdmx.ps1'
    $source = [IO.File]::ReadAllText($scriptPath)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$tokens, [ref]$errors)
    if ($errors) { throw 'Product script does not parse.' }
    foreach ($function in $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        . ([scriptblock]::Create($function.Extent.Text))
    }
    $catalog = Get-EvergreenAdmxReleaseCatalog
    $baseline = @($catalog.Windows | ForEach-Object { $_.Clone() })
    $index = (Invoke-WebRequest -Uri 'https://learn.microsoft.com/en-us/troubleshoot/windows-client/group-policy/create-and-manage-central-store' -UseBasicParsing -TimeoutSec 30).Content
    $before = @($catalog.Windows | ForEach-Object { "$($_.Version)/$($_.Feature)/$($_.DownloadId)" })
    if (Add-NewWindowsRelease $catalog @(Get-WindowsReleaseCandidate $index)) {
        . ([scriptblock]::Create((ConvertTo-ReleaseCatalogFunction $catalog)))
        foreach ($release in $catalog.Windows) {
            if ("$($release.Version)/$($release.Feature)/$($release.DownloadId)" -notin $before) {
                Test-WindowsPolicyPackage $release
                $changes.Add("Windows $($release.Version) $($release.Feature): download $($release.DownloadId)")
            }
        }
    }
    $abbyyIndex = (Invoke-WebRequest -Uri 'https://help.abbyy.com/en-us/' -UseBasicParsing -TimeoutSec 30).Content
    $major = Get-ABBYYReleaseCandidate $abbyyIndex
    if ($major -gt $catalog.ABBYYMajor) {
        $catalog.ABBYYMajor = $major
        # Override only the catalog function in this trusted updater process to validate the candidate.
        . ([scriptblock]::Create((ConvertTo-ReleaseCatalogFunction $catalog)))
        Test-ABBYYPolicyPair (Get-EvergreenAdmxABBYYFineReader)
        $changes.Add("ABBYY FineReader approved major -> $major")
    }
    $foxit = Get-EvergreenAdmxFoxit -MaximumMajor ([int]::MaxValue)
    $foxitMajor = ([version]$foxit.Version).Major
    if ($foxitMajor -gt $catalog.FoxitMajor) {
        $catalog.FoxitMajor = $foxitMajor
        $changes.Add("Foxit approved annual family -> $foxitMajor (Reader/Editor templates $($foxit.Version))")
    }
    if ($changes.Count) {
        $newSource = ConvertTo-UpdatedReleaseCatalogSource $source $catalog
        $testPath = Join-Path $root 'tests/EvergreenAdmx.Tests.ps1'
        $testSource = ConvertTo-UpdatedWindowsTestSource ([IO.File]::ReadAllText($testPath)) $baseline $catalog
        $readmePath = Join-Path $root 'README.md'
        $readme = [IO.File]::ReadAllText($readmePath)
        if ($readme -notmatch '(?s)<!-- release-catalog:start -->.*?<!-- release-catalog:end -->') { throw 'README release catalog markers not found.' }
        $readme = ConvertTo-UpdatedWindowsQuickStart $readme $catalog
        $readme = [regex]::Replace($readme, '(?s)<!-- release-catalog:start -->.*?<!-- release-catalog:end -->', [System.Text.RegularExpressions.MatchEvaluator]{ Get-ReleaseCatalogMarkdown $catalog })
        $changelogPath = Join-Path $root 'CHANGELOG.md'
        $changelog = [IO.File]::ReadAllText($changelogPath)
        $heading = [regex]::Match($changelog, '(?m)^### Added\r?\n')
        if (-not $heading.Success) { throw 'Changelog Added section not found.' }
        $entry = "`n- Automatically discovered product families (requires review): $($changes -join '; ')`n"
        $changelog = $changelog.Insert($heading.Index + $heading.Length, $entry)
        if (-not $DryRun) {
            [IO.File]::WriteAllText($scriptPath, $newSource)
            [IO.File]::WriteAllText($readmePath, $readme)
            [IO.File]::WriteAllText($changelogPath, $changelog)
            [IO.File]::WriteAllText($testPath, $testSource)
        }
        $files = @('EvergreenAdmx.ps1', 'README.md', 'CHANGELOG.md', 'tests/EvergreenAdmx.Tests.ps1')
    }
}
$changes | Write-Output
if ($DryRun) { return }
$result = @{ kind = $Kind; changes = @($changes); files = $files }
[IO.File]::WriteAllText((Join-Path $root 'maintenance-result.json'), ($result | ConvertTo-Json -Depth 5))
