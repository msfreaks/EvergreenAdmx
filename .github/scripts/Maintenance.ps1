#Requires -Version 5.1

function Get-DependencyUpdateKind {
    param([string]$Current, [string]$Latest)
    $old = [version]$Current
    $new = [version]$Latest
    if ($new -le $old) { return 'none' }
    if ($new.Major -gt $old.Major) { return 'major' }
    return 'routine'
}

function Get-MaintenanceFileList {
    param([string]$Kind, [string[]]$Files)
    $allowed = switch ($Kind) {
        'routine' { @('.github/powershell-dependencies.psd1') }
        'major' { @('.github/powershell-dependencies.psd1') }
        'families' { @('EvergreenAdmx.ps1', 'README.md', 'CHANGELOG.md', 'tests/EvergreenAdmx.Tests.ps1') }
        default { throw 'Unknown maintenance update kind.' }
    }
    if ($Files.Count -ne $allowed.Count -or @($allowed | Where-Object { $_ -notin $Files }).Count) {
        throw 'Maintenance artifact contains unexpected file paths.'
    }
    return $allowed
}

function Test-MaintenanceProposalUnchanged {
    param([string]$ExistingHead)
    $existingParent = git show -s --format=%P $ExistingHead
    if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect the maintenance baseline.' }
    $existingTree = git rev-parse "$ExistingHead^{tree}"
    if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect the existing proposal.' }
    $proposedTree = git write-tree
    if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect the proposed changes.' }
    $baseline = git rev-parse HEAD
    if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect the current baseline.' }
    return $existingParent -eq $baseline -and $existingTree -eq $proposedTree
}

function Test-WindowsPolicyPackage {
    param([hashtable]$Release)
    $WorkingDirectory = Join-Path $env:TEMP "EvergreenAdmx-Discovery-$([guid]::NewGuid().ToString('N'))"
    $UseProductFolders = $false
    $PolicyStoreSpecified = $false
    $StampAdmxRevision = $false
    try {
        $null = New-Item -Path (Join-Path $WorkingDirectory 'downloads') -ItemType Directory -Force
        $null = New-Item -Path (Join-Path $WorkingDirectory 'admx') -ItemType Directory -Force
        $null = Invoke-EvergreenAdmxWindows -WindowsVersion $Release.Version -WindowsFeatureVersion $Release.Feature -Languages @('en-US')
        $templates = @(Get-ChildItem (Join-Path $WorkingDirectory 'admx') -Filter '*.admx' -File)
        if (-not $templates) { throw 'New Windows package produced no ADMX templates.' }
        foreach ($template in $templates) {
            $admx = [xml](Get-Content $template.FullName -Raw)
            $languageFile = Join-Path (Join-Path $WorkingDirectory 'admx/en-US') "$($template.BaseName).adml"
            if (-not (Test-Path -LiteralPath $languageFile)) { throw "Missing Windows ADML for $($template.Name)." }
            $adml = [xml](Get-Content $languageFile -Raw)
            if (-not $admx.policyDefinitions -or -not $adml.policyDefinitionResources) { throw 'Invalid Windows policy XML.' }
        }
    } finally {
        if (Test-Path -LiteralPath $WorkingDirectory) { Remove-Item -LiteralPath $WorkingDirectory -Recurse -Force }
    }
}

function Get-WindowsReleaseCandidate {
    param([string]$Content)
    $ids = @([regex]::Matches($Content, 'https://www\.microsoft\.com/(?:[^"''<>\s]+/)?download/details\.aspx\?id=(?<id>\d+)') |
        ForEach-Object { $_.Groups['id'].Value } | Select-Object -Unique)
    if (-not $ids) { throw 'Microsoft ADMX index contains no Download Center links.' }
    foreach ($id in $ids) {
        $page = (Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/details.aspx?id=$id" -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop).Content
        $match = [regex]::Match($page, '(?s)__DLCDetails__=(?<json>.*?)</script>')
        if (-not $match.Success) { throw "Microsoft download $id contains no package metadata." }
        $details = ($match.Groups['json'].Value | ConvertFrom-Json).dlcDetailsView
        if ($details.downloadTitle -match '(?i)\b(preview|insider|beta)\b') { continue }
        $title = [regex]::Match($details.downloadTitle, '(?i)Administrative Templates.*Windows\s+(?:Server\s+)?(?<version>\d{2,4})\b')
        if (-not $title.Success) { continue } # The index also links policy spreadsheets.
        $version = $title.Groups['version'].Value
        $feature = [regex]::Match($details.downloadTitle, '\b\d{2}H[12]\b').Value
        if ([int]$version -lt 2000 -and -not $feature) { continue }
        $file = @($details.downloadFile | Where-Object { $_.url -match '^https://download\.microsoft\.com/.+\.msi$' })
        if ($file.Count -ne 1) { throw "Microsoft ADMX download $id does not contain exactly one official MSI." }
        if ($file[0].version -notmatch '^\d+(?:\.\d+){0,3}$') { throw "Microsoft download $id has an invalid package version." }
        $null = [datetime]::Parse($file[0].datePublished, [cultureinfo]'en-US')
        $packageVersion = [string]$file[0].version
        if ($packageVersion -notmatch '\.') { $packageVersion += '.0' }
        @{ Version = $version; Feature = $feature; DownloadId = $id; Default = $true; PackageVersion = $packageVersion; Published = [datetime]::Parse($file[0].datePublished, [cultureinfo]'en-US') }
    }
}

function Add-NewWindowsRelease {
    param([hashtable]$Catalog, [object[]]$Candidates)
    $changed = $false
    foreach ($candidate in ($Candidates | Sort-Object { [int]$_.Version }, Feature)) {
        $family = @($Catalog.Windows | Where-Object Version -eq $candidate.Version)
        if ($family) {
            $existing = @($family | Where-Object Feature -eq $candidate.Feature)
            if ($existing.Count) {
                if ($candidate.DownloadId -eq $existing[0].DownloadId -or -not $candidate.Published) { continue }
                $approved = @(Get-WindowsReleaseCandidate "https://www.microsoft.com/download/details.aspx?id=$($existing[0].DownloadId)")
                if ($approved.Count -ne 1) { throw 'Cannot compare the reviewed Windows package metadata.' }
                if ([version]$candidate.PackageVersion -lt [version]$approved[0].PackageVersion -or $candidate.Published -le $approved[0].Published) { continue }
                $existing[0].DownloadId = $candidate.DownloadId
                $changed = $true
                continue
            }
            $latest = ($family.Feature | Sort-Object -Descending | Select-Object -First 1)
            if ($candidate.Feature -le $latest) { continue } # Never reintroduce historical releases.
        } else {
            $server = [int]$candidate.Version -ge 2000
            $known = @($Catalog.Windows | Where-Object { ([int]$_.Version -ge 2000) -eq $server })
            if ([int]$candidate.Version -le ($known.Version | ForEach-Object { [int]$_ } | Measure-Object -Maximum).Maximum) { continue }
        }
        foreach ($release in $family) { $release.Default = $false }
        $Catalog.Windows += $candidate
        $changed = $true
    }
    return $changed
}

function Get-ABBYYReleaseCandidate {
    param([string]$Content)
    $majors = @([regex]::Matches($Content, '/en-us/finereader/(?<major>\d+)/admin_guide/') |
        ForEach-Object { [int]$_.Groups['major'].Value } | Sort-Object -Unique -Descending)
    if (-not $majors) { throw 'ABBYY help index contains no Windows FineReader administrator guides.' }
    return $majors[0]
}

function Test-ABBYYPolicyPair {
    param([hashtable]$Package)
    $documents = foreach ($uri in @($Package.URI, $Package.AdmlURI)) {
        [xml](Invoke-WebRequest -Uri $uri -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop).Content
    }
    $admx = $documents[0]
    $adml = $documents[1]
    if (-not $admx.policyDefinitions -or -not $adml.policyDefinitionResources) { throw 'ABBYY downloads are not ADMX/ADML XML.' }
    $strings = @($adml.policyDefinitionResources.resources.stringTable.string.id)
    $presentations = @($adml.policyDefinitionResources.resources.presentationTable.presentation.id)
    foreach ($reference in [regex]::Matches($admx.OuterXml, '\$\((?<kind>string|presentation)\.(?<id>[^)]+)\)')) {
        $resources = if ($reference.Groups['kind'].Value -eq 'string') { $strings } else { $presentations }
        if ($reference.Groups['id'].Value -notin $resources) { throw "Missing ABBYY resource '$($reference.Value)'." }
    }
}

function ConvertTo-ReleaseCatalogFunction {
    param([hashtable]$Catalog)
    $rows = foreach ($release in $Catalog.Windows) {
        if ($release.Version -notmatch '^\d{2,4}$' -or $release.Feature -notmatch '^(\d{2}H[12])?$' -or $release.DownloadId -notmatch '^\d+$') {
            throw 'Invalid Windows release data.'
        }
        $default = if ($release.Default) { '$true' } else { '$false' }
        "            @{ Version = '$($release.Version)'; Feature = '$($release.Feature)'; DownloadId = '$($release.DownloadId)'; Default = $default }"
    }
    $major = [int]$Catalog.ABBYYMajor
    $foxitMajor = [int]$Catalog.FoxitMajor
    return @"
function Get-EvergreenAdmxReleaseCatalog {
    # Embedded because Install-Script distributes EvergreenAdmx.ps1 alone.
    return @{
        ABBYYMajor = $major
        FoxitMajor = $foxitMajor
        Windows = @(
$($rows -join "`n")
        )
    }
}
"@
}

function ConvertTo-UpdatedReleaseCatalogSource {
    param([string]$Source, [hashtable]$Catalog)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($Source, [ref]$tokens, [ref]$errors)
    if ($errors) { throw 'Cannot update a script with syntax errors.' }
    $function = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-EvergreenAdmxReleaseCatalog' }, $true)
    if (-not $function) { throw 'Embedded release catalog not found.' }
    $Source = $Source.Remove($function.Extent.StartOffset, $function.Extent.EndOffset - $function.Extent.StartOffset).Insert($function.Extent.StartOffset, (ConvertTo-ReleaseCatalogFunction $Catalog))
    # A newly reviewed client major becomes the default; server selectors stay explicit.
    $latestClient = ($Catalog.Windows.Version | Where-Object { [int]$_ -lt 2000 } | ForEach-Object { [int]$_ } | Measure-Object -Maximum).Maximum.ToString()
    $Source = [regex]::Replace($Source, '(\[System.String\]\s+\$WindowsVersion\s*=\s*)''\d+''', "`${1}'$latestClient'")
    $Source = [regex]::Replace($Source, '(\[(?:string|int)\]\s*\$WindowsVersion\s*=\s*)''\d+''', "`${1}'$latestClient'")
    $null = [System.Management.Automation.Language.Parser]::ParseInput($Source, [ref]$tokens, [ref]$errors)
    if ($errors) { throw 'Generated release catalog does not parse.' }
    return $Source
}

function Get-ReleaseCatalogMarkdown {
    param([hashtable]$Catalog)
    $lines = @('<!-- release-catalog:start -->', '', '| Windows family | Feature version | Download ID | Default for family |', '| --- | --- | --- | --- |')
    foreach ($release in $Catalog.Windows) {
        $default = if ($release.Default) { 'Yes' } else { 'No' }
        $feature = if ($release.Feature) { $release.Feature } else { 'None' }
        $lines += "| $($release.Version) | $feature | [$($release.DownloadId)](https://www.microsoft.com/en-us/download/details.aspx?id=$($release.DownloadId)) | $default |"
    }
    $lines += @('', "ABBYY FineReader approved major: **$($Catalog.ABBYYMajor)**. Foxit approved annual family: **$($Catalog.FoxitMajor)**.", '', '<!-- release-catalog:end -->')
    return $lines -join "`n"
}

function ConvertTo-UpdatedWindowsQuickStart {
    param([string]$Source, [hashtable]$Catalog)
    $clientVersion = ($Catalog.Windows.Version | Where-Object { [int]$_ -lt 2000 } | ForEach-Object { [int]$_ } | Measure-Object -Maximum).Maximum
    $release = @($Catalog.Windows | Where-Object { [int]$_.Version -eq $clientVersion -and $_.Default })
    if ($release.Count -ne 1) { throw 'Latest Windows client default not found.' }
    $label = "Windows $($release[0].Version) $($release[0].Feature)"
    $pattern = '(?m)^Defaults \(Windows \d+ \d{2}H[12] plus'
    if ($Source -notmatch $pattern) { throw 'README quick-start default text not found.' }
    return [regex]::Replace($Source, $pattern, [System.Text.RegularExpressions.MatchEvaluator]{ "Defaults ($label plus" }, 1)
}

function ConvertTo-UpdatedWindowsTestSource {
    param([string]$Source, [object[]]$Baseline, [hashtable]$Catalog)
    foreach ($old in $Baseline) {
        $release = $Catalog.Windows | Where-Object { $_.Version -eq $old.Version -and $_.Feature -eq $old.Feature } | Select-Object -First 1
        if ($release.DownloadId -eq $old.DownloadId) { continue }
        $featurePattern = if ([int]$old.Version -ge 2000) { "[^']*" } else { [regex]::Escape($old.Feature) }
        $pattern = '(WindowsVersion\s*=\s*' + [regex]::Escape($old.Version) + ';\s*WindowsFeatureVersion\s*=\s*''' + $featurePattern + ''';\s*Expected\s*=\s*'')' + [regex]::Escape($old.DownloadId) + '('')'
        $replacement = '${1}' + $release.DownloadId + '${2}'
        $Source = [regex]::Replace($Source, $pattern, $replacement)
    }
    return $Source
}
