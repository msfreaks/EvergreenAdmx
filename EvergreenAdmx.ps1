<#PSScriptInfo

.VERSION 2607.1

.GUID 999952b7-1337-4018-a1b9-499fad48e734

.AUTHOR Arjan Mensch & Jonathan Pitre

.COMPANYNAME

.COPYRIGHT

.TAGS GroupPolicy GPO Admx Evergreen Automation

.LICENSEURI https://github.com/msfreaks/EvergreenAdmx/blob/main/LICENSE

.PROJECTURI https://github.com/msfreaks/EvergreenAdmx

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

.PRIVATEDATA

#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Downloads the latest Admx files for several products.

.DESCRIPTION
    Automatically downloads the latest Admx files for several products.
    Optionally copies the latest Admx files to a folder of your choosing, for example a Policy Store.

.PARAMETER WindowsVersion
    Specifies Windows major version. Supports 10, 11, 2022 or 2025.
    Default is 11.

.PARAMETER WindowsFeatureVersion
    Specifies Windows 10 or 11 feature version to get the Admx files for.
    Valid values are: 21H2, 22H2 for Windows 10.
    Valid values are: 23H2, 24H2, 25H2 for Windows 11.
    Defaults to 25H2.

    Note: Current Windows 11 ADMX templates (23H2 / 24H2 / 25H2) can also manage Windows 10 clients; some settings apply only to newer OS versions.

.PARAMETER WorkingDirectory
    Specifies a Working Directory for the script.
    Admx files will be stored in a subdirectory called "admx".
    Downloaded files will be stored in a subdirectory called "downloads".
    Defaults to current script location.

.PARAMETER PolicyStore
    Specifies a Policy Store location to copy the Admx files to after processing.

.PARAMETER Languages
    Specifies an array of languages to process. Entries must be in 'xy-XY' format.
    Defaults to 'en-US'.

.PARAMETER UseProductFolders
    Admx files are copied to their respective product folders in a subfolder of 'Admx' in the WorkingDirectory.

.PARAMETER CustomPolicyStore
    Specifies a location for custom policy files. Can be UNC format or local folder.
    Find .admx files in this location, and at least one language folder holding the .adml file(s).
    Versioning will be done based on the newest file found recursively in this location (any .admx or .adml).
    Note that if any file has changed the script will process all files found in location.

.PARAMETER Include
    Array containing Admx products to include when checking for updates.
    Accepts canonical product names and short aliases (ProductKey, compact forms, and former names), e.g. "BISF" for "BIS-F", "Edge" for "Microsoft Edge".
    Valid products are: "Custom Policy Store", "Windows 10", "Windows 11", "Windows 2022", "Windows 2025", "Microsoft Edge", "Microsoft OneDrive", "Microsoft 365 Apps", "Microsoft FSLogix", "Adobe Acrobat", "Adobe Reader", "Adobe DC", "BIS-F", "Citrix Workspace App", "Google Chrome", "Mozilla Firefox", "Mozilla Thunderbird", "Zoom", "Zoom VDI", "Microsoft AVD", "Microsoft Winget", "Microsoft PowerToys", "Windows Terminal", "Brave Browser", "Microsoft Notepad", "Microsoft Clipchamp", "Microsoft Visual Studio", "Microsoft VS Code", "Slack", "1Password", "TeamViewer", "Security ADMX", "Schannel", "Dell Command Update", "Winget-AutoUpdate", "Winget-AutoUpdate-Intune", "PSAppDeployToolkit", "Devolutions Remote Desktop Manager", "Dropbox", "Foxit PDF", "LibreOffice", "HP Anyware", "Specops Authentication Client", "WSL", "Lenovo Dock Manager", "PDF-XChange", "RealVNC Connect", "ABBYY FineReader PDF", "Admin By Request", "GoTo".
    Defaults to "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft 365 Apps", "Microsoft Clipchamp", "Microsoft Notepad", "Microsoft Winget", "Windows Terminal".

.PARAMETER PreferLocalOneDrive
    Microsoft OneDrive Admx files are only available after installing OneDrive.
    If this script is running on a machine that has OneDrive installed locally, use this switch to prevent automatically uninstalling OneDrive.

.PARAMETER CleanPolicyStore
    After processing, remove known obsolete or conflicting Admx/Adml files from -PolicyStore
    (WinStoreUI, Microsoft-Windows-Geolocation-WLPAdm, legacy Office *12-*15, Adobe Classic 2017/2020,
    ctxprofile*, CitrixBase, non-policy junk) and copy missing language ADMLs from en-US when available.
    Requires -PolicyStore. Supports -WhatIf.

.PARAMETER CleanPolicyStoreOnly
    Skip downloads and only clean -PolicyStore using the same obsolete-file and ADML-repair rules as -CleanPolicyStore.
    Requires -PolicyStore. Supports -WhatIf.

.PARAMETER CreateScheduledTask
    Creates or updates a weekly Windows Scheduled Task named EvergreenAdmx that runs this script
    as SYSTEM with highest privileges (Sunday at 01:00). Bound parameters are forwarded to the
    task action except CreateScheduledTask. The script exits after registering the task and does not download Admx files in that run.

.PARAMETER StampAdmxRevision
    When set, stamps ADMX/ADML revision attributes from the product release Version (Major.Minor
    ADMX versionString form) so Intune Imported Administrative Templates show a meaningful Version.
    Only attributes currently set to 1.0 are updated; higher vendor revisions are left unchanged.
    Also updates ADMX resources minRequiredRevision when it is 1.0.

.EXAMPLE
    .\EvergreenAdmx.ps1

    Downloads the latest admx files for Windows 11, Microsoft Edge, Microsoft OneDrive, Microsoft 365 Apps, Microsoft Clipchamp, Microsoft Notepad, Microsoft Winget, and Windows Terminal to the current folder.

.EXAMPLE
    .\EvergreenAdmx.ps1 -WindowsVersion 2025

    Downloads the latest admx files for Windows 2025, Microsoft Edge, Microsoft OneDrive, Microsoft 365 Apps, Microsoft Clipchamp, Microsoft Notepad, Microsoft Winget, and Windows Terminal to the current folder.

.EXAMPLE
    .\EvergreenAdmx.ps1 -WorkingDirectory "C:\Temp\EvergreenAdmx" -Include @('Windows 11', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps', 'Microsoft FSLogix')

    Downloads the latest admx files for the specified products to C:\Temp\EvergreenAdmx folder.

.EXAMPLE
    .\EvergreenAdmx.ps1 -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions" -Languages @("en-US", "nl-NL") -UseProductFolders

    Downloads the default set of products policy definitions files, stores them in product folders for both English and Dutch languages, and copies them to the specified Policy store.

.EXAMPLE
    .\EvergreenAdmx.ps1 -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions" -Languages @("en-US", "nl-NL") -UseProductFolders -CreateScheduledTask

    Creates or updates a weekly SYSTEM scheduled task named EvergreenAdmx that runs this script with the bound parameters (Sunday at 01:00), then exits.

.EXAMPLE
    .\EvergreenAdmx.ps1 -PolicyStore "\\contoso.com\SYSVOL\contoso.com\Policies\PolicyDefinitions" -Include @('Windows 11', 'Microsoft Edge') -CleanPolicyStore

    Downloads the specified products, copies them to the Central Policy Store (creating it if missing), then removes known obsolete Admx files and repairs missing ADMLs.

.EXAMPLE
    .\EvergreenAdmx.ps1 -PolicyStore "\\contoso.com\SYSVOL\contoso.com\Policies\PolicyDefinitions" -CleanPolicyStoreOnly -WhatIf

    Shows which obsolete Admx/Adml files would be removed (and which missing ADMLs would be copied from en-US) without changing the store.

.LINK
    https://github.com/msfreaks/EvergreenAdmx

.LINK
    https://msfreaks.wordpress.com

#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [ValidateSet('10', '11', '2022', '2025')]
    [System.String] $WindowsVersion = '11',
    [Alias('WindowsFeatureEdition')]
    [ValidateSet('21H2', '22H2', '23H2', '24H2', '25H2')]
    [System.String] $WindowsFeatureVersion = $(
        switch ($WindowsVersion) {
            '10' { '22H2' }
            '11' { '25H2' }
            default { '25H2' }
        }
    ),
    [Parameter(Mandatory = $false)]
    [System.String] $WorkingDirectory,
    [Parameter(Mandatory = $false)]
    [System.String] $PolicyStore = $null,
    [Parameter(Mandatory = $false)]
    [System.String[]] $Languages = @('en-US'),
    [Parameter(Mandatory = $false)]
    [switch] $UseProductFolders,
    [Parameter(Mandatory = $false)]
    [System.String] $CustomPolicyStore = $null,
    [Parameter(Mandatory = $false)]
    [switch] $PreferLocalOneDrive,
    [Parameter(Mandatory = $false)]
    [switch] $CleanPolicyStore,
    [Parameter(Mandatory = $false)]
    [switch] $CleanPolicyStoreOnly,
    [Parameter(Mandatory = $false)]
    [switch] $CreateScheduledTask,
    [Parameter(Mandatory = $false)]
    [switch] $StampAdmxRevision,
    [ArgumentCompleter({
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
            $null = $commandName, $parameterName, $fakeBoundParameters
            $candidates = @()
            if (Get-Command -Name Get-EvergreenAdmxProductCatalog -ErrorAction SilentlyContinue) {
                $catalog = Get-EvergreenAdmxProductCatalog
                $candidates = foreach ($product in $catalog) {
                    @($product.Name) + @($product.Aliases)
                }
            }
            elseif ($commandAst -and $commandAst.Extent.File -and (Test-Path -LiteralPath $commandAst.Extent.File)) {
                $tokens = $null
                $errors = $null
                $ast = [System.Management.Automation.Language.Parser]::ParseFile($commandAst.Extent.File, [ref]$tokens, [ref]$errors)
                $fn = $ast.FindAll({
                        param($node)
                        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                        $node.Name -eq 'Get-EvergreenAdmxProductCatalog'
                    }, $true) | Select-Object -First 1
                if ($fn) {
                    $catalog = & ([scriptblock]::Create($fn.Extent.Text + "; Get-EvergreenAdmxProductCatalog"))
                    $candidates = foreach ($product in $catalog) {
                        @($product.Name) + @($product.Aliases)
                    }
                }
            }
            $candidates |
                Sort-Object -Unique |
                Where-Object { $_ -like "$wordToComplete*" } |
                ForEach-Object {
                    [System.Management.Automation.CompletionResult]::new(
                        "'$($_ -replace "'", "''")'",
                        $_,
                        'ParameterValue',
                        $_
                    )
                }
        })]
    [System.String[]] $Include = $(
        switch ($WindowsVersion) {
            '10' { @('Windows 10', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps', 'Microsoft Clipchamp', 'Microsoft Notepad', 'Microsoft Winget', 'Windows Terminal') }
            '11' { @('Windows 11', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps', 'Microsoft Clipchamp', 'Microsoft Notepad', 'Microsoft Winget', 'Windows Terminal') }
            '2022' { @('Windows 2022', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps', 'Microsoft Clipchamp', 'Microsoft Notepad', 'Microsoft Winget', 'Windows Terminal') }
            '2025' { @('Windows 2025', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps', 'Microsoft Clipchamp', 'Microsoft Notepad', 'Microsoft Winget', 'Windows Terminal') }
            default { @('Windows 11', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps', 'Microsoft Clipchamp', 'Microsoft Notepad', 'Microsoft Winget', 'Windows Terminal') }
        }
    )
)

#region init


function Get-EvergreenAdmxProductCatalog {
    <#
    .SYNOPSIS
        Returns canonical -Include product names and accepted aliases.
    #>
    [CmdletBinding()]
    [OutputType([System.Array])]
    param()

    @(
        [PSCustomObject]@{ Name = 'Custom Policy Store'; Aliases = @('CustomPolicyStore', 'CustomPolicyLocation') }
        [PSCustomObject]@{ Name = 'Windows 10'; Aliases = @('Windows10', 'Win10') }
        [PSCustomObject]@{ Name = 'Windows 11'; Aliases = @('Windows11', 'Win11') }
        [PSCustomObject]@{ Name = 'Windows 2022'; Aliases = @('Windows2022', 'WindowsServer2022', 'Win2022') }
        [PSCustomObject]@{ Name = 'Windows 2025'; Aliases = @('Windows2025', 'WindowsServer2025', 'Win2025') }
        [PSCustomObject]@{ Name = 'Microsoft Edge'; Aliases = @('Edge', 'MSEdge', 'MicrosoftEdge') }
        [PSCustomObject]@{ Name = 'Microsoft OneDrive'; Aliases = @('OneDrive', 'MicrosoftOneDrive') }
        [PSCustomObject]@{ Name = 'Microsoft 365 Apps'; Aliases = @('365Apps', 'Microsoft365Apps', 'M365', 'Office', 'Microsoft Office', 'MicrosoftOffice') }
        [PSCustomObject]@{ Name = 'Microsoft FSLogix'; Aliases = @('FSLogix', 'MicrosoftFSLogix') }
        [PSCustomObject]@{ Name = 'Adobe Acrobat'; Aliases = @('AdobeAcrobat', 'Acrobat') }
        [PSCustomObject]@{ Name = 'Adobe Reader'; Aliases = @('AdobeReader', 'Reader') }
        [PSCustomObject]@{ Name = 'Adobe DC'; Aliases = @('AdobeDC') }
        [PSCustomObject]@{ Name = 'BIS-F'; Aliases = @('BISF', 'Base Image Script Framework', 'BaseImageScriptFramework') }
        [PSCustomObject]@{ Name = 'Citrix Workspace App'; Aliases = @('CitrixWorkspaceApp', 'Citrix', 'CWA') }
        [PSCustomObject]@{ Name = 'Google Chrome'; Aliases = @('GoogleChrome', 'Chrome') }
        [PSCustomObject]@{ Name = 'Mozilla Firefox'; Aliases = @('MozillaFirefox', 'Firefox') }
        [PSCustomObject]@{ Name = 'Mozilla Thunderbird'; Aliases = @('MozillaThunderbird', 'Thunderbird') }
        [PSCustomObject]@{ Name = 'Zoom'; Aliases = @('Zoom Desktop Client', 'ZoomDesktopClient') }
        [PSCustomObject]@{ Name = 'Zoom VDI'; Aliases = @('ZoomVDI') }
        [PSCustomObject]@{ Name = 'Microsoft AVD'; Aliases = @('AVD', 'AzureVirtualDesktop', 'Azure Virtual Desktop', 'MicrosoftAVD', 'Microsoft Azure Virtual Desktop') }
        [PSCustomObject]@{ Name = 'Microsoft Winget'; Aliases = @('Winget', 'MicrosoftWinget') }
        [PSCustomObject]@{ Name = 'Microsoft PowerToys'; Aliases = @('PowerToys', 'MicrosoftPowerToys') }
        [PSCustomObject]@{ Name = 'Windows Terminal'; Aliases = @('WindowsTerminal', 'Terminal', 'WT') }
        [PSCustomObject]@{ Name = 'Brave Browser'; Aliases = @('Brave', 'BraveBrowser') }
        [PSCustomObject]@{ Name = 'Microsoft Notepad'; Aliases = @('Notepad', 'MicrosoftNotepad') }
        [PSCustomObject]@{ Name = 'Microsoft Clipchamp'; Aliases = @('Clipchamp', 'MicrosoftClipchamp') }
        [PSCustomObject]@{ Name = 'Microsoft Visual Studio'; Aliases = @('VisualStudio', 'MicrosoftVisualStudio', 'VS') }
        [PSCustomObject]@{ Name = 'Microsoft VS Code'; Aliases = @('VSCode', 'MicrosoftVSCode', 'VisualStudioCode') }
        [PSCustomObject]@{ Name = 'Slack'; Aliases = @() }
        [PSCustomObject]@{ Name = '1Password'; Aliases = @('OnePassword') }
        [PSCustomObject]@{ Name = 'TeamViewer'; Aliases = @() }
        [PSCustomObject]@{ Name = 'Security ADMX'; Aliases = @('SecurityAdmx', 'SecurityADMX', 'Security-ADMX') }
        [PSCustomObject]@{ Name = 'Schannel'; Aliases = @() }
        [PSCustomObject]@{ Name = 'Dell Command Update'; Aliases = @('DellCommandUpdate', 'DCU') }
        [PSCustomObject]@{ Name = 'Winget-AutoUpdate'; Aliases = @('WingetAutoUpdate', 'WAU') }
        [PSCustomObject]@{ Name = 'Winget-AutoUpdate-Intune'; Aliases = @('WingetAutoUpdateIntune', 'WAUIntune', 'WAU-Intune') }
        [PSCustomObject]@{ Name = 'PSAppDeployToolkit'; Aliases = @('PSADT') }
        [PSCustomObject]@{ Name = 'Devolutions Remote Desktop Manager'; Aliases = @('DevolutionsRDM', 'Devolutions', 'RDM') }
        [PSCustomObject]@{ Name = 'Dropbox'; Aliases = @() }
        [PSCustomObject]@{ Name = 'Foxit PDF'; Aliases = @('FoxitPDF', 'Foxit') }
        [PSCustomObject]@{ Name = 'LibreOffice'; Aliases = @('Collabora', 'CollaboraOffice') }
        [PSCustomObject]@{ Name = 'HP Anyware'; Aliases = @('HPAnyware', 'Anyware', 'PCoIP') }
        [PSCustomObject]@{ Name = 'Specops Authentication Client'; Aliases = @('Specops', 'SpecopsClient', 'SpecopsAuthenticationClient', 'SPR', 'uReset', 'SpecopsPasswordReset') }
        [PSCustomObject]@{ Name = 'WSL'; Aliases = @('Windows Subsystem for Linux', 'WindowsSubsystemForLinux') }
        [PSCustomObject]@{ Name = 'Lenovo Dock Manager'; Aliases = @('DockManager', 'LenovoDockManager') }
        [PSCustomObject]@{ Name = 'PDF-XChange'; Aliases = @('PDFXChange', 'PDF-X-Change', 'Tracker') }
        [PSCustomObject]@{ Name = 'RealVNC Connect'; Aliases = @('RealVNC', 'VNC', 'RealVNCServer', 'RealVNCViewer') }
        [PSCustomObject]@{ Name = 'ABBYY FineReader PDF'; Aliases = @('ABBYY', 'FineReader', 'FineReaderPDF', 'ABBYYFineReader') }
        [PSCustomObject]@{ Name = 'Admin By Request'; Aliases = @('AdminByRequest', 'ABR') }
        [PSCustomObject]@{ Name = 'GoTo'; Aliases = @('GoToConnect', 'GoTo App', 'GoToApp', 'LogMeIn GoTo') }
    )
}

function Get-EvergreenAdmxRemovedIncludeMap {
    <#
    .SYNOPSIS
        Maps removed -Include product names to a short removal reason.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    return @{
        'Microsoft Desktop Optimization Pack' = 'Microsoft Desktop Optimization Pack (MDOP) support was removed (extended support ended April 14, 2026).'
        'MDOP'                                = 'Microsoft Desktop Optimization Pack (MDOP) support was removed (extended support ended April 14, 2026).'
        'MicrosoftDesktopOptimizationPack'    = 'Microsoft Desktop Optimization Pack (MDOP) support was removed (extended support ended April 14, 2026).'
        'Adobe Acrobat Classic 2017'          = 'Adobe Acrobat Classic tracks were removed; use Continuous track via ''Adobe Acrobat''.'
        'Adobe Acrobat Classic 2020'          = 'Adobe Acrobat Classic tracks were removed; use Continuous track via ''Adobe Acrobat''.'
        'Acrobat2017'                         = 'Adobe Acrobat Classic tracks were removed; use Continuous track via ''Adobe Acrobat''.'
        'Acrobat2020'                         = 'Adobe Acrobat Classic tracks were removed; use Continuous track via ''Adobe Acrobat''.'
        'Adobe Reader Classic 2017'           = 'Adobe Reader Classic tracks were removed; use Continuous track via ''Adobe Reader''.'
        'Adobe Reader Classic 2020'           = 'Adobe Reader Classic tracks were removed; use Continuous track via ''Adobe Reader''.'
        'AcrobatReader2017'                   = 'Adobe Reader Classic tracks were removed; use Continuous track via ''Adobe Reader''.'
        'AcrobatReader2020'                   = 'Adobe Reader Classic tracks were removed; use Continuous track via ''Adobe Reader''.'
    }
}

function Resolve-EvergreenAdmxInclude {
    <#
    .SYNOPSIS
        Resolves -Include values (canonical names or aliases) to canonical product names.
    #>
    [CmdletBinding()]
    [OutputType([System.Object[]])]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.String[]] $Include
    )

    $catalog = @(Get-EvergreenAdmxProductCatalog)
    $lookup = @{}
    foreach ($product in $catalog) {
        $lookup[$product.Name.ToLowerInvariant()] = $product.Name
        foreach ($alias in @($product.Aliases)) {
            if ([string]::IsNullOrWhiteSpace($alias)) { continue }
            $key = $alias.ToLowerInvariant()
            if ($lookup.ContainsKey($key) -and $lookup[$key] -ne $product.Name) {
                throw "Duplicate -Include alias '$alias' maps to both '$($lookup[$key])' and '$($product.Name)'."
            }
            $lookup[$key] = $product.Name
        }
    }

    $removed = Get-EvergreenAdmxRemovedIncludeMap
    $resolved = [System.Collections.Generic.List[System.String]]::new()
    $seen = @{}

    foreach ($value in $Include) {
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $trimmed = $value.Trim()
        $key = $trimmed.ToLowerInvariant()

        $removedEntry = $removed.GetEnumerator() | Where-Object { $_.Key.ToLowerInvariant() -eq $key } | Select-Object -First 1
        if ($removedEntry) {
            throw "-Include value '$trimmed' is no longer supported. $($removedEntry.Value)"
        }

        if (-not $lookup.ContainsKey($key)) {
            $validList = ($catalog.Name | Sort-Object | ForEach-Object { "  $_" }) -join [Environment]::NewLine
            throw @"
Cannot resolve -Include value '$trimmed'. Valid products:
$validList
Aliases such as ProductKey short names are accepted (e.g. 'BISF' for 'BIS-F').
"@
        }

        $canonical = $lookup[$key]
        if (-not $seen.ContainsKey($canonical)) {
            $seen[$canonical] = $true
            $resolved.Add($canonical)
        }
    }

    return , @($resolved.ToArray())
}

# Validate feature version based on Windows version
if ($WindowsVersion -eq '2022' -and ($PSBoundParameters.ContainsKey('WindowsFeatureVersion'))) {
    Write-Warning 'Windows feature version parameters are ignored when WindowsVersion is set to 2022'
} elseif ($WindowsVersion -eq '2025' -and ($PSBoundParameters.ContainsKey('WindowsFeatureVersion'))) {
    Write-Warning 'Windows feature version parameters are ignored when WindowsVersion is set to 2025'
}

$ProgressPreference = 'SilentlyContinue'
#$ErrorActionPreference = 'SilentlyContinue'

$AdmxVersions = $null
if (-not $WorkingDirectory) { $WorkingDirectory = $PWD }
if (Test-Path -Path "$($WorkingDirectory)\AdmxVersions.xml") { $AdmxVersions = Import-Clixml -Path "$($WorkingDirectory)\AdmxVersions.xml" }
if ($null -eq $AdmxVersions) { $AdmxVersions = @{} }
if (-not (Test-Path -Path "$($WorkingDirectory)\admx")) { $null = New-Item -Path "$($WorkingDirectory)\admx" -ItemType Directory -Force }
if (-not (Test-Path -Path "$($WorkingDirectory)\downloads")) { $null = New-Item -Path "$($WorkingDirectory)\downloads" -ItemType Directory -Force }
$PolicyStoreSpecified = $PSBoundParameters.ContainsKey('PolicyStore')
if (($CleanPolicyStore -or $CleanPolicyStoreOnly) -and -not $PolicyStoreSpecified) {
    throw '-CleanPolicyStore and -CleanPolicyStoreOnly require -PolicyStore.'
}
if ($PolicyStore -and -not $PolicyStore.EndsWith('\')) { $PolicyStore += '\' }
elseif ($null -eq $PolicyStore) { $PolicyStore = $PWD }
# Allow 'xx', 'xx-XX', and numeric region tags like 'es-419'
if ($Languages -notmatch '^([A-Za-z]{2})(-([A-Za-z]{2}|\d{3}))?$') { Write-Warning "Language not in expected format: $($Languages -notmatch '^([A-Za-z]{2})(-([A-Za-z]{2}|\d{3}))?$')" }
if ($CustomPolicyStore -and -not (Test-Path -Path "$($CustomPolicyStore)")) { throw "'$($CustomPolicyStore)' is not a valid path." }
if ($CustomPolicyStore -and -not $CustomPolicyStore.EndsWith('\')) { $CustomPolicyStore += '\' }
if ($CustomPolicyStore -and (Get-ChildItem -Path $CustomPolicyStore -Directory) -notmatch '^([A-Za-z]{2})(-([A-Za-z]{2}|\d{3}))?$') { throw "'$($CustomPolicyStore)' does not contain at least one subfolder matching the language format (e.g 'en-US', 'es', 'es-419')." }
$Include = @(Resolve-EvergreenAdmxInclude -Include $Include)
if ($PSBoundParameters.ContainsKey('Include')) {
    $PSBoundParameters['Include'] = $Include
}

If ($PreferLocalOneDrive -and $Include -notcontains 'Microsoft OneDrive') {
    $Include += 'Microsoft OneDrive'
}

# Parameter debugging
Write-Verbose "Windows Version:`t'$($WindowsVersion)'"
If ($WindowsVersion -eq '10' -or $WindowsVersion -eq '11') {
    Write-Verbose "Windows Feature Version:`t'$($WindowsFeatureVersion)'"
}
Write-Verbose "WorkingDirectory:`t'$($WorkingDirectory)'"
If ($PolicyStore) {
    Write-Verbose "PolicyStore:`t'$($PolicyStore)'"
}
If ($CustomPolicyStore) {
    Write-Verbose "Add admx Path:`t`'$($CustomPolicyStore)'"
}
Write-Verbose "Languages:`t`'$($Languages)'"
Write-Verbose "Use product folders:`t'$($UseProductFolders)'"
Write-Verbose "Admx path:`t`'$($WorkingDirectory)\admx'"
Write-Verbose "Download path:`t`'$($WorkingDirectory)\downloads'"
Write-Verbose "Included:`t`'$($Include -join ', ')'"
Write-Verbose "PreferLocalOneDrive:`t'$($PreferLocalOneDrive)'"
Write-Verbose "CleanPolicyStore:`t'$($CleanPolicyStore)'"
Write-Verbose "CleanPolicyStoreOnly:`t'$($CleanPolicyStoreOnly)'"
Write-Verbose "CreateScheduledTask:`t'$($CreateScheduledTask)'"
Write-Verbose "StampAdmxRevision:`t'$($StampAdmxRevision)'"

function New-EvergreenAdmxTaskArgumentList {
    <#
    .SYNOPSIS
        Builds powershell.exe arguments for the EvergreenAdmx scheduled task action.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([System.Collections.Generic.List[System.String]])]
    param (
        [Parameter(Mandatory = $true)]
        [System.String] $ScriptPath,
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary] $BoundParameters
    )

    if (-not $PSCmdlet.ShouldProcess($ScriptPath, 'Build scheduled task argument list')) {
        return
    }

    $ArgumentList = [System.Collections.Generic.List[System.String]]::new()
    $ArgumentList.Add('-NoProfile')
    $ArgumentList.Add('-ExecutionPolicy')
    $ArgumentList.Add('Bypass')
    $ArgumentList.Add('-File')
    $ArgumentList.Add("`"$ScriptPath`"")

    foreach ($Key in $BoundParameters.Keys) {
        if ($Key -eq 'CreateScheduledTask') { continue }
        $Value = $BoundParameters[$Key]
        if ($Value -is [System.Management.Automation.SwitchParameter]) {
            if ($Value.IsPresent) { $ArgumentList.Add("-$Key") }
            continue
        }
        if ($Value -is [System.Array]) {
            $QuotedItems = foreach ($Item in $Value) {
                "'$($Item -replace "'", "''")'"
            }
            $ArgumentList.Add("-$Key")
            $ArgumentList.Add("@($($QuotedItems -join ','))")
            continue
        }
        $ArgumentList.Add("-$Key")
        $ArgumentList.Add("`"$($Value -replace '"', '`"')`"")
    }

    return $ArgumentList
}

if ($CreateScheduledTask) {
    $ScriptPath = $PSCommandPath
    if (-not $ScriptPath) { $ScriptPath = $MyInvocation.MyCommand.Path }
    if (-not $ScriptPath -or -not (Test-Path -LiteralPath $ScriptPath)) {
        throw 'Unable to resolve the path to EvergreenAdmx.ps1 for scheduled task registration.'
    }

    $ArgumentList = New-EvergreenAdmxTaskArgumentList -ScriptPath $ScriptPath -BoundParameters $PSBoundParameters
    $TaskAction = New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument ($ArgumentList -join ' ')
    $TaskTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 1am
    $TaskPrincipal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $TaskSettings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 1) -MultipleInstances IgnoreNew
    $null = Register-ScheduledTask -TaskName 'EvergreenAdmx' -Action $TaskAction -Trigger $TaskTrigger -Principal $TaskPrincipal -Settings $TaskSettings -Force
    Write-Output "Scheduled task 'EvergreenAdmx' created or updated (weekly on Sunday at 01:00, runs as SYSTEM)."
    return
}
#endregion

#region functions
function Get-Link {
    <#
    .SYNOPSIS
        Returns a specific link from a web page.

    .DESCRIPTION
        Returns a specific link from a web page.

    .NOTES
        Site: https://packageology.com
        Author: Dan Gough
        Twitter: @packageologist

    .LINK
        https://github.com/DanGough/Nevergreen

    .PARAMETER Uri
        The URI to query.

    .PARAMETER MatchProperty
        Which property the RegEx pattern should be applied to, e.g. href, outerHTML, class, title.

    .PARAMETER Pattern
        The RegEx pattern to apply to the selected property. Supply an array of patterns to receive multiple links.

    .PARAMETER ReturnProperty
        Optional. Specifies which property to return from the link. Defaults to href, but 'data-filename' can also be useful to retrieve.

    .PARAMETER UserAgent
        Optional parameter to provide a user agent for Invoke-WebRequest to use. Examples are:

        Googlebot: 'Googlebot/2.1 (+http://www.google.com/bot.html)'
        Microsoft Edge: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.135 Safari/537.36 Edge/12.246'

    .EXAMPLE
        Get-Link -Uri 'http://somewhere.com' -MatchProperty href -Pattern '\.exe$'

        Description:
        Returns first download link matching *.exe from http://somewhere.com.
    #>
    [CmdletBinding(SupportsShouldProcess = $False)]
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline)]
        [ValidatePattern('^(http|https)://')]
        [Alias('Url')]
        [String] $Uri,
        [Parameter(
            Mandatory = $true,
            Position = 1)]
        [ValidateNotNullOrEmpty()]
        #[ValidateSet('href', 'outerHTML', 'innerHTML', 'outerText', 'innerText', 'class', 'title', 'tagName', 'data-filename')]
        [String] $MatchProperty,
        [Parameter(
            Mandatory = $true,
            Position = 2)]
        [ValidateNotNullOrEmpty()]
        [String[]] $Pattern,
        [Parameter(
            Mandatory = $false,
            Position = 3)]
        [ValidateNotNullOrEmpty()]
        [String] $ReturnProperty = 'href',
        [Parameter(
            Mandatory = $false)]
        [String] $UserAgent,
        [System.Collections.Hashtable] $Headers,
        [Switch] $PrefixDomain,
        [Switch] $PrefixParent
    )

    process {
        $ProgressPreference = 'SilentlyContinue'

        $ParamHash = @{
            Uri = $Uri
            Method = 'GET'
            UseBasicParsing = $True
            DisableKeepAlive = $True
            ErrorAction = 'Stop'
        }

        if ($UserAgent) {
            $ParamHash.UserAgent = $UserAgent
        }

        if ($Headers) {
            $ParamHash.Headers = $Headers
        }

        try {
            $Response = Invoke-WebRequest @ParamHash

            foreach ($CurrentPattern in $Pattern) {
                $Link = $Response.Links | Where-Object $MatchProperty -Match $CurrentPattern | Select-Object -First 1 -ExpandProperty $ReturnProperty

                if ($PrefixDomain) {
                    $BaseURL = ($Uri -split '/' | Select-Object -First 3) -join '/'
                    $Link = Set-UriPrefix -Uri $Link -Prefix $BaseURL
                } elseif ($PrefixParent) {
                    $BaseURL = ($Uri -split '/' | Select-Object -SkipLast 1) -join '/'
                    $Link = Set-UriPrefix -Uri $Link -Prefix $BaseURL
                }

                $Link

            }
        } catch {
            Write-Error "$($MyInvocation.MyCommand): $($_.Exception.Message)"
        }
    }

}

function Get-Version {
    <#
    .SYNOPSIS
        Extracts a version number from either a string or the content of a web page using a chosen or pre-defined match pattern.

    .DESCRIPTION
        Extracts a version number from either a string or the content of a web page using a chosen or pre-defined match pattern.

    .NOTES
        Site: https://packageology.com
        Author: Dan Gough
        Twitter: @packageologist

    .LINK
        https://github.com/DanGough/Nevergreen

    .PARAMETER String
        The string to process.

    .PARAMETER Uri
        The Uri to load web content from to process.

    .PARAMETER UserAgent
        Optional parameter to provide a user agent for Invoke-WebRequest to use. Examples are:

        Googlebot: 'Googlebot/2.1 (+http://www.google.com/bot.html)'
        Microsoft Edge: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.135 Safari/537.36 Edge/12.246'

    .PARAMETER Pattern
        Optional RegEx pattern to use for version matching. Pattern to return must be included in parentheses.

    .PARAMETER ReplaceWithDot
        Switch to automatically replace characters - or _ with . in detected version.

    .EXAMPLE
        Get-Version -String 'http://somewhere.com/somefile_1.2.3.exe'

        Description:
        Returns '1.2.3'
    #>
    [CmdletBinding(SupportsShouldProcess = $False)]
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ParameterSetName = 'String')]
        [ValidateNotNullOrEmpty()]
        [String[]] $String,
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'Uri')]
        [ValidatePattern('^(http|https)://')]
        [String] $Uri,
        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'Uri')]
        [String] $UserAgent,
        [Parameter(
            Mandatory = $false,
            Position = 1)]
        [ValidateNotNullOrEmpty()]
        [String] $Pattern = '((?:\d+\.)+\d+)',
        [Switch] $ReplaceWithDot
    )

    begin {

    }

    process {

        if ($PsCmdlet.ParameterSetName -eq 'Uri') {

            $ProgressPreference = 'SilentlyContinue'

            try {
                $ParamHash = @{
                    Uri = $Uri
                    Method = 'GET'
                    UseBasicParsing = $True
                    DisableKeepAlive = $True
                    ErrorAction = 'Stop'
                }

                if ($UserAgent) {
                    $ParamHash.UserAgent = $UserAgent
                }

                $String = (Invoke-WebRequest @ParamHash).Content
            } catch {
                Write-Error "Unable to query URL '$Uri': $($_.Exception.Message)"
            }

        }

        foreach ($CurrentString in $String) {
            if ($ReplaceWithDot) {
                $CurrentString = $CurrentString.Replace('-', '.').Replace('+', '.').Replace('_', '.')
            }
            if ($CurrentString -match $Pattern) {
                $matches[1]
            } else {
                Write-Warning "No version found within $CurrentString using pattern $Pattern"
            }

        }

    }

    end {
    }

}

function Resolve-Uri {
    <#
    .SYNOPSIS
        Resolves a URI and also returns the filename and last modified date if found.

    .DESCRIPTION
        Resolves a URI and also returns the filename and last modified date if found.

    .NOTES
        Site: https://packageology.com
        Author: Dan Gough
        Twitter: @packageologist

    .LINK
        https://github.com/DanGough/Nevergreen

    .PARAMETER Uri
        The URI resolve. Accepts an array of strings or pipeline input.

    .PARAMETER UserAgent
        Optional parameter to provide a user agent for Invoke-WebRequest to use. Examples are:

        Googlebot: 'Googlebot/2.1 (+http://www.google.com/bot.html)'
        Microsoft Edge: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.135 Safari/537.36 Edge/12.246'

    .EXAMPLE
        Resolve-Uri -Uri 'http://somewhere.com/somefile.exe'

        Description:
        Returns the absolute redirected URI, filename and last modified date.
    #>
    [CmdletBinding(SupportsShouldProcess = $False)]
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName)]
        [ValidatePattern('^(http|https)://')]
        [Alias('Url')]
        [String[]] $Uri,
        [Parameter(
            Mandatory = $false,
            Position = 1)]
        [String] $UserAgent,
        [System.Collections.Hashtable] $Headers
    )

    begin {
        $ProgressPreference = 'SilentlyContinue'
    }

    process {

        foreach ($UriToResolve in $Uri) {

            try {

                $ParamHash = @{
                    Uri = $UriToResolve
                    Method = 'Head'
                    UseBasicParsing = $True
                    DisableKeepAlive = $True
                    ErrorAction = 'Stop'
                }

                if ($UserAgent) {
                    $ParamHash.UserAgent = $UserAgent
                }

                if ($Headers) {
                    $ParamHash.Headers = $Headers
                }

                $Response = Invoke-WebRequest @ParamHash

                if ($IsCoreCLR) {
                    $ResolvedUri = $Response.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
                } else {
                    $ResolvedUri = $Response.BaseResponse.ResponseUri.AbsoluteUri
                }

                Write-Verbose "$($MyInvocation.MyCommand): URI resolved to: $ResolvedUri"

                #PowerShell 7 returns each header value as single unit arrays instead of strings which messes with the -match operator coming up, so use Select-Object:
                $ContentDisposition = $Response.Headers.'Content-Disposition' | Select-Object -First 1

                if ($ContentDisposition -match 'filename="?([^\\/:\*\?"<>\|]+)') {
                    $FileName = $matches[1]
                    Write-Verbose "$($MyInvocation.MyCommand): Content-Disposition header found: $ContentDisposition"
                    Write-Verbose "$($MyInvocation.MyCommand): File name determined from Content-Disposition header: $FileName"
                } else {
                    $Slug = [uri]::UnescapeDataString($ResolvedUri.Split('?')[0].Split('/')[-1])
                    if ($Slug -match '^[^\\/:\*\?"<>\|]+\.[^\\/:\*\?"<>\|]+$') {
                        Write-Verbose "$($MyInvocation.MyCommand): URI slug is a valid file name: $FileName"
                        $FileName = $Slug
                    } else {
                        $FileName = $null
                    }
                }

                try {
                    $LastModified = [DateTime]($Response.Headers.'Last-Modified' | Select-Object -First 1)
                    Write-Verbose "$($MyInvocation.MyCommand): Last modified date: $LastModified"
                } catch {
                    Write-Verbose "$($MyInvocation.MyCommand): Unable to parse date from last modified header: $($Response.Headers.'Last-Modified')"
                    $LastModified = $null
                }

            } catch {
                Throw "$($MyInvocation.MyCommand): Unable to resolve URI: $($_.Exception.Message)"
            }

            if ($ResolvedUri) {
                [PSCustomObject]@{
                    Uri = $ResolvedUri
                    FileName = $FileName
                    LastModified = $LastModified
                }
            }

        }
    }

    end {
    }

}

function Invoke-FileDownload {
    <#
    .SYNOPSIS
        Downloads a file to disk using System.Net.Http.HttpClient.

    .DESCRIPTION
        Streams a remote file to a temporary path, then moves it to -OutFile.
        Faster and more memory-efficient than Invoke-WebRequest -OutFile on
        Windows PowerShell 5.1. Compatible with Windows PowerShell 5.1 and PowerShell 7+.

    .PARAMETER Uri
        The URI of the file to download.

    .PARAMETER OutFile
        Full path of the destination file.

    .PARAMETER UseDefaultCredentials
        Use the current user's default network credentials for the request.

    .PARAMETER UserAgent
        Optional User-Agent header value. Defaults to 'EvergreenAdmx' when omitted
        (required by hosts such as api.github.com).

    .PARAMETER Headers
        Optional additional request headers.

    .PARAMETER TimeoutSeconds
        Maximum seconds allowed for the download (default 3600).
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [Alias('Url')]
        [string]$Uri,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$OutFile,

        [switch]$UseDefaultCredentials,

        [string]$UserAgent = 'EvergreenAdmx',

        [hashtable]$Headers,

        [ValidateRange(1, 86400)]
        [int]$TimeoutSeconds = 3600
    )

    # System.Net.Http is not always loaded in Windows PowerShell 5.1
    Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue

    # Prefer TLS 1.3; fall back to TLS 1.2 when 1.3 is unavailable (older OS/.NET)
    $tlsFlags = [Net.SecurityProtocolType]::Tls12
    if ([enum]::GetNames([Net.SecurityProtocolType]) -contains 'Tls13') {
        $tlsFlags = $tlsFlags -bor [Net.SecurityProtocolType]::Tls13
    }
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $tlsFlags

    $destinationDirectory = [System.IO.Path]::GetDirectoryName($OutFile)
    if ($destinationDirectory -and -not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
        $null = New-Item -Path $destinationDirectory -ItemType Directory -Force
    }

    $tempFilePath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ((New-Guid).ToString('N') + '.tmp')
    $timeout = [System.TimeSpan]::FromSeconds($TimeoutSeconds)
    # .NET Framework CancellationTokenSource has int-ms ctor, not TimeSpan
    $cancellationTokenSource = [System.Threading.CancellationTokenSource]::new([int]$timeout.TotalMilliseconds)

    $handler = [System.Net.Http.HttpClientHandler]::new()
    $handler.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
    try {
        $sslProtocols = [System.Security.Authentication.SslProtocols]::Tls12
        if ([enum]::GetNames([System.Security.Authentication.SslProtocols]) -contains 'Tls13') {
            $sslProtocols = $sslProtocols -bor [System.Security.Authentication.SslProtocols]::Tls13
        }
        $handler.SslProtocols = $sslProtocols
    } catch {
        Write-Verbose "$($MyInvocation.MyCommand): Unable to set HttpClientHandler.SslProtocols; relying on SecurityProtocol defaults"
    }
    if ($UseDefaultCredentials) {
        $handler.UseDefaultCredentials = $true
    }

    $httpClient = [System.Net.Http.HttpClient]::new($handler)
    $httpClient.Timeout = [System.Threading.Timeout]::InfiniteTimeSpan

    if ($UserAgent) {
        $httpClient.DefaultRequestHeaders.TryAddWithoutValidation('User-Agent', $UserAgent) | Out-Null
    }
    if ($Headers) {
        foreach ($header in $Headers.GetEnumerator()) {
            $httpClient.DefaultRequestHeaders.TryAddWithoutValidation($header.Key, [string]$header.Value) | Out-Null
        }
    }

    $response = $null
    $responseStream = $null
    $fileStream = $null

    try {
        Write-Verbose "$($MyInvocation.MyCommand): Downloading '$Uri' to '$OutFile'"

        $response = $httpClient.GetAsync($Uri, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead, $cancellationTokenSource.Token).GetAwaiter().GetResult()
        if (-not $response.IsSuccessStatusCode) {
            throw "Download failed with status $([int]$response.StatusCode) ($($response.ReasonPhrase)) for '$Uri'"
        }

        # .NET Framework ReadAsStreamAsync has no CancellationToken overload
        if ($IsCoreCLR) {
            $responseStream = $response.Content.ReadAsStreamAsync($cancellationTokenSource.Token).GetAwaiter().GetResult()
        } else {
            $readTask = $response.Content.ReadAsStreamAsync()
            $waitMs = [Math]::Max(1, [int]$timeout.TotalMilliseconds)
            if (-not $readTask.Wait($waitMs)) {
                throw [System.TimeoutException]::new("Timeout exceeded while opening response stream for '$Uri'")
            }
            $responseStream = $readTask.GetAwaiter().GetResult()
        }

        $fileStream = [System.IO.File]::Create($tempFilePath)
        $buffer = New-Object byte[] 65536

        while ($true) {
            $readBytes = $responseStream.ReadAsync($buffer, 0, $buffer.Length, $cancellationTokenSource.Token).GetAwaiter().GetResult()
            if ($readBytes -eq 0) {
                break
            }
            $fileStream.Write($buffer, 0, $readBytes)
        }

        $fileStream.Close()
        $fileStream.Dispose()
        $fileStream = $null

        Move-Item -LiteralPath $tempFilePath -Destination $OutFile -Force
        Write-Verbose "$($MyInvocation.MyCommand): Download complete: '$OutFile'"
    } catch [System.OperationCanceledException] {
        throw [System.TimeoutException]::new("Timeout exceeded while downloading '$Uri' (limit ${TimeoutSeconds}s)")
    } finally {
        if ($fileStream) { $fileStream.Dispose() }
        if ($responseStream) { $responseStream.Dispose() }
        if ($response) { $response.Dispose() }
        if ($httpClient) { $httpClient.Dispose() }
        if ($handler) { $handler.Dispose() }
        if ($cancellationTokenSource) { $cancellationTokenSource.Dispose() }
        if ((Test-Path -LiteralPath $tempFilePath -PathType Leaf)) {
            Remove-Item -LiteralPath $tempFilePath -Force -ErrorAction SilentlyContinue
        }
    }
}

function ConvertTo-AdmxRevisionString {
    <#
    .SYNOPSIS
        Converts a product release Version to an ADMX versionString (Major.Minor).
    #>
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Version
    )

    $normalized = $Version.Trim()
    if ($normalized -match '^v(?<ver>.+)$') {
        $normalized = $Matches['ver']
    }

    try {
        $parsed = [version]$normalized
        return ('{0}.{1}' -f $parsed.Major, $parsed.Minor)
    } catch {
        Write-Warning "Unable to convert '$Version' to ADMX revision (Major.Minor)."
        return $null
    }
}

function Test-AdmxRevisionIsOnePointZero {
    <#
    .SYNOPSIS
        Returns $true when an ADMX/ADML revision attribute is exactly 1.0.
    #>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    try {
        return ([version]$Value) -eq ([version]'1.0')
    } catch {
        return ($Value -eq '1.0')
    }
}

function Save-AdmxXmlDocument {
    <#
    .SYNOPSIS
        Saves an XmlDocument as UTF-8 without BOM.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Xml,

        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $encoding = New-Object System.Text.UTF8Encoding $false
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = $encoding
    # PreserveWhitespace on load; avoid re-indenting the document on save
    $settings.Indent = $false
    $settings.OmitXmlDeclaration = $false
    $writer = [System.Xml.XmlWriter]::Create($FilePath, $settings)
    try {
        $Xml.Save($writer)
    } finally {
        $writer.Dispose()
    }
}

function Set-AdmxRevision {
    <#
    .SYNOPSIS
        Stamps ADMX/ADML revision attributes when the current value is exactly 1.0.

    .DESCRIPTION
        Updates policyDefinitions/@revision, resources/@minRequiredRevision (ADMX), and
        policyDefinitionResources/@revision (ADML) only when each attribute is currently 1.0.
        Higher vendor revisions are left unchanged.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Revision
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Verbose "Set-AdmxRevision: path '$Path' does not exist"
        return
    }

    $admxFiles = @(Get-ChildItem -LiteralPath $Path -Filter '*.admx' -File -ErrorAction SilentlyContinue)
    $admlFiles = @(Get-ChildItem -LiteralPath $Path -Filter '*.adml' -File -Recurse -ErrorAction SilentlyContinue)

    foreach ($file in $admxFiles) {
        if (-not $PSCmdlet.ShouldProcess($file.FullName, "Stamp ADMX revision to $Revision when 1.0")) { continue }
        try {
            $xml = New-Object System.Xml.XmlDocument
            $xml.PreserveWhitespace = $true
            $xml.Load($file.FullName)
            $root = $xml.DocumentElement
            if ($null -eq $root -or $root.LocalName -ne 'policyDefinitions') {
                Write-Verbose "Skipping '$($file.Name)': root is not policyDefinitions"
                continue
            }

            $changed = $false
            $currentRevision = $root.GetAttribute('revision')
            if (Test-AdmxRevisionIsOnePointZero -Value $currentRevision) {
                $root.SetAttribute('revision', $Revision)
                $changed = $true
                Write-Verbose "Stamped '$($file.Name)' revision $currentRevision -> $Revision"
            } else {
                Write-Verbose "Skipped '$($file.Name)' revision '$currentRevision' (not 1.0)"
            }

            $resources = $root.SelectSingleNode('*[local-name()="resources"]')
            if ($resources) {
                $minRequired = $resources.GetAttribute('minRequiredRevision')
                if (Test-AdmxRevisionIsOnePointZero -Value $minRequired) {
                    $resources.SetAttribute('minRequiredRevision', $Revision)
                    $changed = $true
                    Write-Verbose "Stamped '$($file.Name)' minRequiredRevision $minRequired -> $Revision"
                } else {
                    Write-Verbose "Skipped '$($file.Name)' minRequiredRevision '$minRequired' (not 1.0)"
                }
            }

            if ($changed) {
                Save-AdmxXmlDocument -Xml $xml -FilePath $file.FullName
            }
        } catch {
            Write-Warning "Failed to stamp ADMX revision on '$($file.FullName)': $_"
        }
    }

    foreach ($file in $admlFiles) {
        if (-not $PSCmdlet.ShouldProcess($file.FullName, "Stamp ADML revision to $Revision when 1.0")) { continue }
        try {
            $xml = New-Object System.Xml.XmlDocument
            $xml.PreserveWhitespace = $true
            $xml.Load($file.FullName)
            $root = $xml.DocumentElement
            if ($null -eq $root -or $root.LocalName -ne 'policyDefinitionResources') {
                Write-Verbose "Skipping '$($file.Name)': root is not policyDefinitionResources"
                continue
            }

            $currentRevision = $root.GetAttribute('revision')
            if (Test-AdmxRevisionIsOnePointZero -Value $currentRevision) {
                $root.SetAttribute('revision', $Revision)
                Write-Verbose "Stamped '$($file.Name)' revision $currentRevision -> $Revision"
                Save-AdmxXmlDocument -Xml $xml -FilePath $file.FullName
            } else {
                Write-Verbose "Skipped '$($file.Name)' revision '$currentRevision' (not 1.0)"
            }
        } catch {
            Write-Warning "Failed to stamp ADML revision on '$($file.FullName)': $_"
        }
    }
}

function Copy-Admx {
    param (
        [string]$SourceFolder,
        [string]$TargetFolder,
        [string]$PolicyStore = $null,
        [string]$ProductName,
        [switch]$Quiet,
        [string[]]$Languages = $null,
        [string]$Revision = $null
    )
    if (-not (Test-Path -Path "$($TargetFolder)")) { $null = (New-Item -Path "$($TargetFolder)" -ItemType Directory -Force) }
    if (-not $Languages -or $Languages -eq '') { $Languages = @('en-US') }

    Write-Verbose "Copying Admx files from '$($SourceFolder)' to '$($TargetFolder)'"
    Copy-Item -Path "$($SourceFolder)\*.admx" -Destination "$($TargetFolder)" -Force
    foreach ($language in $Languages) {
        if (-not (Test-Path -Path "$($SourceFolder)\$($language)")) {
            Write-Verbose "$($language) not found"
            if (-not $Quiet) { Write-Warning "Language '$($language)' not found for '$($ProductName)'. Processing 'en-US' instead." }
            $language = 'en-US'
        }
        if (-not (Test-Path -Path "$($TargetFolder)\$($language)")) {
            Write-Verbose "'$($TargetFolder)\$($language)' does not exist, creating folder"
            $null = (New-Item -Path "$($TargetFolder)\$($language)" -ItemType Directory -Force)
        }
        Write-Verbose "Copying '$($SourceFolder)\$($language)\*.adml' to '$($TargetFolder)\$($language)'"
        Copy-Item -Path "$($SourceFolder)\$($language)\*.adml" -Destination "$($TargetFolder)\$($language)" -Force
    }
    if ($PolicyStore) {
        Write-Verbose "Copying Admx files from '$($SourceFolder)' to '$($PolicyStore)'"
        Copy-Item -Path "$($SourceFolder)\*.admx" -Destination "$($PolicyStore)" -Force
        foreach ($language in $Languages) {
            if (-not (Test-Path -Path "$($SourceFolder)\$($language)")) { $language = 'en-US' }
            if (-not (Test-Path -Path "$($PolicyStore)$($language)")) { $null = (New-Item -Path "$($PolicyStore)$($language)" -ItemType Directory -Force) }
            Copy-Item -Path "$($SourceFolder)\$($language)\*.adml" -Destination "$($PolicyStore)$($language)" -Force
        }
    }

    if ($Revision) {
        $normalizedRevision = ConvertTo-AdmxRevisionString -Version $Revision
        if ($normalizedRevision) {
            Write-Verbose "Stamping ADMX/ADML revision '$normalizedRevision' for '$ProductName'"
            Set-AdmxRevision -Path $TargetFolder -Revision $normalizedRevision
            if ($PolicyStore) {
                Set-AdmxRevision -Path $PolicyStore -Revision $normalizedRevision
            }
        }
    }
}

function Get-EvergreenAdmxObsoleteFilePattern {
    <#
    .SYNOPSIS
        Returns file name patterns for obsolete or conflicting Policy Store Admx/Adml files.
    #>
    [CmdletBinding()]
    [OutputType([System.Object[]])]
    param()

    return @(
        'WinStoreUI.admx'
        'WinStoreUI.adml'
        'Microsoft-Windows-Geolocation-WLPAdm.admx'
        'Microsoft-Windows-Geolocation-WLPAdm.adml'
        '*12*.admx'
        '*12*.adml'
        '*13*.admx'
        '*13*.adml'
        '*14*.admx'
        '*14*.adml'
        '*15*.admx'
        '*15*.adml'
        'Acrobat2017.admx'
        'Acrobat2017.adml'
        'AcrobatReader2017.admx'
        'AcrobatReader2017.adml'
        'Acrobat2020.admx'
        'Acrobat2020.adml'
        'AcrobatReader2020.admx'
        'AcrobatReader2020.adml'
        'ctxprofile*.admx'
        'ctxprofile*.adml'
        'CitrixBase.admx'
        'CitrixBase.adml'
    )
}

function Initialize-PolicyStore {
    <#
    .SYNOPSIS
        Creates a Central Policy Store folder and language subfolders when missing.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)]
        [string]$PolicyStore,
        [Parameter(Mandatory = $false)]
        [string[]]$Languages = @('en-US')
    )

    if (-not $Languages -or $Languages -eq '') { $Languages = @('en-US') }

    $store = $PolicyStore.TrimEnd('\')
    if (-not (Test-Path -LiteralPath $store)) {
        if ($PSCmdlet.ShouldProcess($store, 'Create Policy Store directory')) {
            Write-Verbose "Creating Policy Store '$store'"
            $null = New-Item -Path $store -ItemType Directory -Force
        }
    }

    foreach ($language in $Languages) {
        $languagePath = Join-Path -Path $store -ChildPath $language
        if (-not (Test-Path -LiteralPath $languagePath)) {
            if ($PSCmdlet.ShouldProcess($languagePath, 'Create Policy Store language folder')) {
                Write-Verbose "Creating Policy Store language folder '$languagePath'"
                $null = New-Item -Path $languagePath -ItemType Directory -Force
            }
        }
    }
}

function Clear-ObsoleteAdmx {
    <#
    .SYNOPSIS
        Removes known obsolete or conflicting Admx/Adml files from a Policy Store,
        and copies missing language ADMLs from en-US when available.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([System.Object[]])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$PolicyStore,
        [Parameter(Mandatory = $false)]
        [string[]]$Languages = @('en-US')
    )

    if (-not $Languages -or $Languages -eq '') { $Languages = @('en-US') }

    $store = $PolicyStore.TrimEnd('\')
    if (-not (Test-Path -LiteralPath $store)) {
        Write-Verbose "Policy Store '$store' does not exist; nothing to clean."
        return @()
    }

    $removed = [System.Collections.Generic.List[string]]::new()
    $patterns = Get-EvergreenAdmxObsoleteFilePattern
    $languagePattern = '^([A-Za-z]{2})(-([A-Za-z]{2}|\d{3}))?$'

    $obsoleteFiles = @(
        Get-ChildItem -Path $store -Recurse -File -Include $patterns -ErrorAction SilentlyContinue
    )
    foreach ($file in $obsoleteFiles) {
        if ($PSCmdlet.ShouldProcess($file.FullName, 'Remove obsolete policy definition file')) {
            Write-Verbose "Removing obsolete file '$($file.FullName)'"
            Remove-Item -LiteralPath $file.FullName -Force
            $removed.Add($file.FullName)
        }
    }

    $junkFiles = @(
        Get-ChildItem -LiteralPath $store -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -notin @('.admx', '.adml') }
    )
    foreach ($file in $junkFiles) {
        if ($PSCmdlet.ShouldProcess($file.FullName, 'Remove non-policy file from Policy Store')) {
            Write-Verbose "Removing non-policy file '$($file.FullName)'"
            Remove-Item -LiteralPath $file.FullName -Force
            $removed.Add($file.FullName)
        }
    }

    $junkFolders = @(
        Get-ChildItem -LiteralPath $store -Directory -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -notmatch $languagePattern -and
                ($Languages -notcontains $_.Name)
            }
    )
    foreach ($folder in $junkFolders) {
        if ($PSCmdlet.ShouldProcess($folder.FullName, 'Remove non-language folder from Policy Store')) {
            Write-Verbose "Removing non-language folder '$($folder.FullName)'"
            Remove-Item -LiteralPath $folder.FullName -Recurse -Force
            $removed.Add($folder.FullName)
        }
    }

    # Repair missing ADMLs: copy en-US into requested language folders when the ADMX has no matching ADML
    $enUsFolder = Join-Path -Path $store -ChildPath 'en-US'
    $admxFiles = @(
        Get-ChildItem -LiteralPath $store -Filter '*.admx' -File -ErrorAction SilentlyContinue |
            Where-Object {
                $name = $_.Name
                -not ($patterns | Where-Object { $name -like $_ })
            }
    )
    $languageFolders = @(
        Get-ChildItem -LiteralPath $store -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match $languagePattern -or ($Languages -contains $_.Name) }
    )

    foreach ($admx in $admxFiles) {
        $admlName = "$($admx.BaseName).adml"
        $enUsAdml = Join-Path -Path $enUsFolder -ChildPath $admlName

        foreach ($language in $Languages) {
            if ($language -eq 'en-US') { continue }
            $languageFolder = Join-Path -Path $store -ChildPath $language
            $targetAdml = Join-Path -Path $languageFolder -ChildPath $admlName
            if (Test-Path -LiteralPath $targetAdml) { continue }
            if (-not (Test-Path -LiteralPath $enUsAdml)) { continue }

            if ($PSCmdlet.ShouldProcess($targetAdml, "Copy missing ADML from en-US for language '$language'")) {
                if (-not (Test-Path -LiteralPath $languageFolder)) {
                    $null = New-Item -Path $languageFolder -ItemType Directory -Force
                }
                Write-Warning "Missing ADML for '$($admx.Name)' in '$language'; copying from en-US."
                Copy-Item -LiteralPath $enUsAdml -Destination $targetAdml -Force
            }
        }

        $hasAnyAdml = $false
        foreach ($folder in $languageFolders) {
            if (Test-Path -LiteralPath (Join-Path -Path $folder.FullName -ChildPath $admlName)) {
                $hasAnyAdml = $true
                break
            }
        }
        if (-not $hasAnyAdml -and (Test-Path -LiteralPath $enUsAdml)) {
            $hasAnyAdml = $true
        }
        if (-not $hasAnyAdml) {
            Write-Warning "No ADML found for '$($admx.Name)' in any language folder under '$store'."
        }
    }

    Write-Verbose "Policy Store cleanup removed $($removed.Count) item(s)"
    return $removed.ToArray()
}

# Get-EvergreenAdmx functions
function Get-EvergreenAdmxFSLogix {
    <#
    .SYNOPSIS
        Returns latest download url and version for Microsoft FSLogix policy definitions files.
    #>

    try {
        # Grab URI (redirected url)
        $URL = 'https://aka.ms/fslogix/download'
        $URI = (Resolve-Uri -Uri $URL).URI
        # Grab version
        $Version = Get-Version -String $URI -Pattern '(\d+(\.\d+){1,4})'

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmx365App {
    <#
    .SYNOPSIS
        Returns latest both x86 and x64 download url and version for Microsoft 365 Apps policy definitions files.
    #>

    $id = '49030'
    $urlVersion = "https://www.microsoft.com/en-us/download/details.aspx?id=$($id)"
    $JSONBlobPattern = '(?<scriptStart><script>[\w.]+__DLCDetails__=).*?(?<JSObject-scriptStart></script>)'

    try {

        # Load web page for scrapping url version
        $web = Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $urlVersion -MaximumRedirection 0
        # Grab version
        $regEx = '(version\":")((?:\d+\.)+(?:\d+))"'
        $version = ($web.RawContent | Select-String -Pattern $regEx).Matches.Groups[2].Value

        # Carve JSON from script tag
        $web = $web.Content | Select-String -Pattern $JSONBlobPattern | Select-Object -ExpandProperty Matches | ForEach-Object { $_.Groups['JSObject'].Value } | Select-Object -First 1 | ConvertFrom-Json
        # Grab x64 version
        $hrefx64 = $web.dlcDetailsView.downloadFile | Where-Object { $_.url -like '*x64*' } | Select-Object -First 1
        # Grab x86 version
        $hrefx86 = $web.dlcDetailsView.downloadFile | Where-Object { $_.url -like '*x86*' } | Select-Object -First 1

        # Return evergreen object
        return @( @{ Version = $version; URI = $hrefx64.url; Architecture = 'x64' }, @{ Version = $version; URI = $hrefx86.url; Architecture = 'x86' })
    } catch {
        Throw $_
    }
}

function Get-WindowsDownloadId {
    <#
    .SYNOPSIS
        Returns Windows admx download Id

    .PARAMETER WindowsVersion
        Specifies Windows major version. Supports 10, 11, 2022 or 2025. Default is 11.

    .PARAMETER WindowsFeatureVersion
        Specifies Windows client feature edition. Default is 25H2.

    .EXAMPLE
        Get-WindowsDownloadId -WindowsVersion 11 -WindowsFeatureVersion 25H2
    #>

    param (
        [Parameter(Position = 0)]
        [ValidateSet('10', '11', '2022', '2025')]
        [ValidateNotNullOrEmpty()]
        [int]$WindowsVersion = '11',
        [Parameter(Position = 1)]
        [ValidateScript({
                if ($WindowsVersion -eq '10' -and $_ -in @('21H2', '22H2')) {
                    return $true
                } elseif ($WindowsVersion -eq '11' -and $_ -in @('23H2', '24H2', '25H2')) {
                    return $true
                } elseif ($WindowsVersion -eq '2022' -or $WindowsVersion -eq '2025') {
                    return $true
                } else {
                    throw "Invalid Windows Feature Version '$_' for Windows $WindowsVersion. Windows 10 supports: 21H2, 22H2. Windows 11 supports: 23H2, 24H2, 25H2. Windows 2022 and 2025 has no Windows Feature Versions."
                }
            })]
        [ValidateNotNullOrEmpty()]
        [string]$WindowsFeatureVersion = '25H2'
    )

    switch ($WindowsVersion) {
        10 {
            return (@( @{ '21H2' = '104042' }, @{ '22H2' = '104677' } ).$WindowsFeatureVersion)
            break
        }
        11 {
            return (@( @{ '23H2' = '105667' }, @{ '24H2' = '106254' }, @{ '25H2' = '108542' } ).$WindowsFeatureVersion)
            break
        }
        2022 {
            return @( '104003' )
            break
        }
        2025 {
            return @( '108430' )
            break
        }
    }
}

function Get-EvergreenAdmxWindows {
    <#
    .SYNOPSIS
        Returns url and latest version for Windows 10, Windows 11 or Windows Server 2022 or 2025 policy definitions files.

    .PARAMETER DownloadId
        Id returned from Get-WindowsDownloadId. Default is the latest version for the current Windows major version.
    #>

    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param(
        [int]$DownloadId = (Get-WindowsDownloadId)
    )

    $urlVersion = "https://www.microsoft.com/en-us/download/details.aspx?id=$($DownloadId)"
    $JSONBlobPattern = '(?<scriptStart><script>[\w.]+__DLCDetails__=).*?(?<JSObject-scriptStart></script>)'

    try {

        # Load web page for scrapping url version
        $web = Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $urlVersion

        # Grab version
        $regEx = '(version\":")((?:\d+\.)+(?:\d+))"'
        $version = ('{0}.{1}' -f $DownloadId, ($web | Select-String -Pattern $regEx).Matches.Groups[2].Value)

        # Carve JSON from script tag
        $web = $web.Content | Select-String -Pattern $JSONBlobPattern | Select-Object -ExpandProperty Matches | ForEach-Object { $_.Groups['JSObject'].Value } | Select-Object -First 1 | ConvertFrom-Json
        $href = $web.dlcDetailsView.downloadFile | Where-Object { $_.url -like '*.msi' } | Select-Object -First 1

        # Return Evergreen object
        return @{ Version = $version; URI = $href.url }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmxOneDrive {
    <#
    .SYNOPSIS
        Returns url and latest version for Microsoft OneDrive policy definitions files.
    #>

    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param (
        [switch]$PreferLocalOneDrive
    )

    try {
        # Detect if OneDrive is installed
        if (Get-Variable -Name isOneDriveInstalled -ErrorAction SilentlyContinue) {
            Clear-Variable -Name isOneDriveInstalled -Force
        }
        $UserInstall = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where-Object { $_.DisplayName -like '*OneDrive*' }) | Sort-Object -Property DisplayVersion -Descending | Select-Object -First 1
        $Systemx64Install = (Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where-Object { $_.DisplayName -like '*OneDrive*' }) | Sort-Object -Property DisplayVersion -Descending | Select-Object -First 1
        $Systemx86Install = (Get-ItemProperty -Path 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where-Object { $_.DisplayName -like '*OneDrive*' }) | Sort-Object -Property DisplayVersion -Descending | Select-Object -First 1
        $url = 'https://evergreen-api.stealthpuppy.com/app/MicrosoftOneDrive'
        $architecture = 'x64'
        $ring = 'Insider'
        $type = 'exe'
        $Evergreen = Invoke-RestMethod -Uri $url -UserAgent 'Googlebot/2.1 (+http://www.google.com/bot.html)'
        $Evergreen = $Evergreen | Where-Object { $_.Architecture -eq $architecture -and $_.Ring -eq $ring -and $_.Type -eq $type } | `
                Sort-Object -Property @{ Expression = { [System.Version]$_.Version }; Descending = $true } | Select-Object -First 1

        $oneDriveADMXFolder = $null
        If (-not [string]::IsNullOrWhiteSpace($UserInstall)) {
            Write-Verbose "User OneDrive install found: $($UserInstall.DisplayVersion)"
            $isOneDriveInstalled = $true
            $OneDriveInstalledVersion = $UserInstall.DisplayVersion
            $oneDriveADMXFolder = (Get-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\OneDrive').CurrentVersionPath
        } elseif (-not [string]::IsNullOrWhiteSpace($Systemx64Install)) {
            Write-Verbose "System x64 OneDrive install found: $($Systemx64Install.DisplayVersion)"
            $isOneDriveInstalled = $true
            $OneDriveInstalledVersion = $Systemx64Install.DisplayVersion
            $oneDriveADMXFolder = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\OneDrive').CurrentVersionPath
        } elseif (-not [string]::IsNullOrWhiteSpace($Systemx86Install)) {
            Write-Verbose "System x86 OneDrive install found: $($Systemx86Install.DisplayVersion)"
            $isOneDriveInstalled = $true
            $OneDriveInstalledVersion = $Systemx86Install.DisplayVersion
            $oneDriveADMXFolder = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\OneDrive').CurrentVersionPath
        } else {
            $isOneDriveInstalled = $false
        }

        if ($PreferLocalOneDrive) {
            If ($isOneDriveInstalled) {
                return @{ Version = $OneDriveInstalledVersion; ADMXFolder = $oneDriveADMXFolder }
            } else {
                Write-Warning 'No local installation of Microsoft OneDrive install found.'
                # Grab download uri
                $URI = $Evergreen.URI

                # Grab version
                $Version = $Evergreen.Version

                # Return evergreen object
                return @{ Version = $Version; URI = $URI }
            }
        } else {
            # Grab download uri
            $URI = $Evergreen.URI

            # Grab version
            $Version = $Evergreen.Version

            # Return evergreen object
            return @{ Version = $Version; URI = $URI }
        }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmxEdge {
    <#
    .SYNOPSIS
        Returns url and latest version for Microsoft Edge policy definitions files.
    #>

    try {

        $url = 'https://edgeupdates.microsoft.com/api/products?view=enterprise'
        # Grab json containing product info
        $json = Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing -MaximumRedirection 0 | ConvertFrom-Json
        # Filter out the newest release
        $release = ($json | Where-Object { $_.Product -like 'Policy' }).Releases | Sort-Object ProductVersion -Descending | Select-Object -First 1
        # Grab version
        $Version = $release.ProductVersion
        # Grab uri
        $URI = $release.Artifacts[0].Location

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }

}

function Get-EvergreenAdmxChrome {
    <#
    .SYNOPSIS
        Returns url and latest version for Google Chrome policy definitions files.
    #>

    try {

        $DownloadUrl = 'https://dl.google.com/dl/edgedl/chrome/policy/policy_templates.zip'

        $url = 'https://evergreen-api.stealthpuppy.com/app/GoogleChrome'
        $channel = 'Stable'
        $architecture = 'x64'
        $type = 'msi'
        $Evergreen = Invoke-RestMethod -Uri $url -UserAgent 'Googlebot/2.1 (+http://www.google.com/bot.html)'
        $Evergreen = $Evergreen | Where-Object { $_.Channel -eq $channel -and $_.Architecture -eq $architecture -and $_.Type -eq $type } | `
                Sort-Object -Property @{ Expression = { [System.Version]$_.Version }; Descending = $true } | Select-Object -First 1

        $Version = $Evergreen.Version

        # Return evergreen object
        return @{ Version = $Version; URI = $DownloadUrl }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmxAdobeAcrobat {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for the Adobe Acrobat Continuous track Admx files. Use this for Acrobat, 64-bit Reader, and the unified installer.
    #>

    $URL = 'https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/misc/AcrobatADMTemplate.zip'

    try {
        # grab uri
        $URI = (Resolve-Uri -Uri $URL).URI

        # grab version
        $LastModifiedDate = (Resolve-Uri -Uri $URL).LastModified
        [version]$Version = $LastModifiedDate.ToString('yyyy.MM.dd')

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmxAdobeReader {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for the Adobe Reader Continuous track Admx files.
    #>

    $URL = 'https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/misc/ReaderADMTemplate.zip'

    try {
        # grab uri
        $URI = (Resolve-Uri -Uri $URL).URI

        # grab version
        $LastModifiedDate = (Resolve-Uri -Uri $URL).LastModified
        [version]$Version = $LastModifiedDate.ToString('yyyy.MM.dd')

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmxWorkplaceApp {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Citrix Workspace App ADMX files
    #>

    try {

        $url = 'https://www.citrix.com/downloads/workspace-app/windows/workspace-app-for-windows-latest.html'
        # grab content
        $web = (Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing -DisableKeepAlive).RawContent
        # find line with ADMX download
        $str = ($web -split "`r`n" | Select-String -Pattern '_ADMX_')[0].ToString().Trim()
        # extract url from ADMX download string
        $URI = "https:$(((Select-String '(\/\/)([^\s,]+)(?=")' -Input $str).Matches.Value))"
        # grab version
        $VersionRegEx = 'Version\: ((?:\d+\.)+(?:\d+)) \((.+)\)'
        $Version = ($web | Select-String -Pattern $VersionRegEx).Matches.Groups[1].Value

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmxFirefox {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Mozilla Firefox ADMX files
    #>

    try {

        # define github repo
        $repo = 'mozilla/policy-templates'
        # grab latest release properties
        $latest = (Invoke-WebRequest -UseDefaultCredentials -Uri "https://api.github.com/repos/$($repo)/releases" -UseBasicParsing | ConvertFrom-Json)[0]

        # grab version
        $Version = ($latest.tag_name | Select-String -Pattern '(\d+(\.\d+){1,4})' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        # grab uri
        $URI = $latest.assets.browser_download_url

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmxBISF {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for BIS-F ADMX files
    #>

    try {

        # define github repo
        $repo = 'EUCweb/BIS-F'
        # grab latest release properties
        $latest = (Invoke-WebRequest -UseDefaultCredentials -Uri "https://api.github.com/repos/$($repo)/releases" -UseBasicParsing | ConvertFrom-Json)[0]

        # grab version
        $Version = ($latest.tag_name | Select-String -Pattern '(\d+(\.\d+){1,4})' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        # grab uri
        $URI = $latest.zipball_url

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmxZoom {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Zoom ADMX files
    #>

    try {
        $url = 'https://support.zoom.com/hc/en/article?id=zm_kb&sysparm_article=KB0065466'

        # grab content
        $web = Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing -UserAgent 'Googlebot/2.1 (+http://www.google.com/bot.html)'
        # find ADMX download
        $URI = (($web.Links | Where-Object { $_.href -like '*msi-templates*.zip' })[-1]).href
        # grab version
        $Version = ($URI.Split('/')[-1] | Select-String -Pattern '(\d+(\.\d+){1,4})' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmxZoomVDI {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Zoom VDI ADMX files
    #>

    try {
        $url = 'https://support.zoom.com/hc/en/article?id=zm_kb&sysparm_article=KB0064784'

        # grab content
        $web = Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing -UserAgent 'Googlebot/2.1 (+http://www.google.com/bot.html)'
        # find ADMX download
        $URI = (($web.Links | Where-Object { $_.href -like '*msi-templates*.zip' })[-1]).href
        # grab version
        $Version = ($URI.Split('/')[-1] | Select-String -Pattern '(\d+(\.\d+){1,4})' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Get-CustomPolicyOnline {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Custom Policies

    .PARAMETER CustomPolicyStore
        Folder where Custom Policies can be found
    #>

    param(
        [string] $CustomPolicyStore
    )

    $newestFileDate = Get-Date -Date ((Get-ChildItem -Path $CustomPolicyStore -Include '*.admx', '*.adml' -Recurse | Sort-Object LastWriteTime -Descending) | Select-Object -First 1).LastWriteTime

    $version = Get-Date -Date $newestFileDate -Format 'yyMM.dd.HHmmss'

    return @{ Version = $version; URI = $CustomPolicyStore }
}

function Get-EvergreenAdmxAVD {
    <#
    .SYNOPSIS
        Returns latest url and version for MicrosoftMicrosoft AVD policy definition files.
    #>

    try {
        $URL = 'https://aka.ms/avdgpo'

        # Grab uri
        $URI = (Resolve-Uri -Uri $URL).URI

        # Grab version
        $LastModifiedDate = (Resolve-Uri -Uri $URL).LastModified
        [version]$Version = $LastModifiedDate.ToString('yyyy.MM.dd')

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmxWinget {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Winget-cli ADMX files
    #>

    try {

        # Define github repo
        $repo = 'microsoft/winget-cli'
        # Grab latest release properties
        $latest = ((Invoke-WebRequest -UseDefaultCredentials -Uri "https://api.github.com/repos/$($repo)/releases" -UseBasicParsing | ConvertFrom-Json) | Where-Object { $_.name -notlike '*-preview' -and $_.draft -eq $false -and $_.assets.browser_download_url -match 'DesktopAppInstallerPolicies.zip' })[0]

        # Grab version
        $Version = ($latest.tag_name | Select-String -Pattern '(\d+(\.\d+){1,4})' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        # Grab uri
        $URI = $latest.assets.browser_download_url | Where-Object { $_ -like '*/DesktopAppInstallerPolicies.zip' }

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmxBrave {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Brave ADMX files
    #>

    try {

        # define github repo
        $repo = 'brave/brave-browser'
        # grab latest release properties
        $latest = ((Invoke-WebRequest -UseDefaultCredentials -Uri "https://api.github.com/repos/$($repo)/releases/latest" -UseBasicParsing | ConvertFrom-Json) | Where-Object { $_.assets.browser_download_url -match 'policy_templates.zip' })[0]

        # grab version
        $Version = ($latest.tag_name | Select-String -Pattern '(\d+(\.\d+){1,4})' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        # grab uri
        $URI = $latest.assets.browser_download_url | Where-Object { $_ -like '*/policy_templates.zip' }

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Get-EvergreenAdmxNotepad {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Windows Notepad ADMX files
    #>

    try {
        $URL = 'https://download.microsoft.com/download/72ea16a9-4cc9-4032-945d-3a56a483d034/WindowsNotepadAdminTemplates.cab'

        # Grab uri
        $URI = (Resolve-Uri -Uri $URL).URI

        # Grab version
        $LastModifiedDate = (Resolve-Uri -Uri $URL).LastModified
        [version]$Version = $LastModifiedDate.ToString('yyyy.MM.dd')

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxNotepad {
    <#
    .SYNOPSIS
        Process Windows Notepad Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxNotepad
    $ProductName = 'Microsoft Notepad'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            $null = (New-Item -Path $TempFolder -ItemType Directory -Force)
            Push-Location $TempFolder
            tar -xf $OutFile
            $zipFile = Get-ChildItem -Path $TempFolder -Filter '*.zip' | Select-Object -First 1
            Expand-Archive -Path $zipFile.FullName -DestinationPath $TempFolder -Force
            Pop-Location

            # copy
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'WindowsNotepad.admx' | Select-Object -First 1).DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxClipchamp {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Microsoft Clipchamp ADMX files
    #>

    try {
        $DownloadId = 105674
        $urlVersion = "https://www.microsoft.com/en-us/download/details.aspx?id=$($DownloadId)"
        $JSONBlobPattern = '(?<scriptStart><script>[\w.]+__DLCDetails__=).*?(?<JSObject-scriptStart></script>)'

        # Load web page for scrapping url version
        $web = Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $urlVersion

        # Grab version
        $regEx = '(version\":")((?:\d+\.)+(?:\d+))"'
        $version = ('{0}.{1}' -f $DownloadId, ($web | Select-String -Pattern $regEx).Matches.Groups[2].Value)

        # Carve JSON from script tag
        $web = $web.Content | Select-String -Pattern $JSONBlobPattern | Select-Object -ExpandProperty Matches | ForEach-Object { $_.Groups['JSObject'].Value } | Select-Object -First 1 | ConvertFrom-Json
        $href = $web.dlcDetailsView.downloadFile | Where-Object { $_.url -like '*.zip' } | Select-Object -First 1

        # Return Evergreen object
        return @{ Version = $version; URI = $href.url }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxClipchamp {
    <#
    .SYNOPSIS
        Process Microsoft Clipchamp Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxClipchamp
    $ProductName = 'Microsoft Clipchamp'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            # copy
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter '*.admx' | Select-Object -First 1).DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxVisualStudio {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Microsoft Visual Studio ADMX files
    #>

    try {
        $DownloadId = 104405
        $urlVersion = "https://www.microsoft.com/en-us/download/details.aspx?id=$($DownloadId)"
        $JSONBlobPattern = '(?<scriptStart><script>[\w.]+__DLCDetails__=).*?(?<JSObject-scriptStart></script>)'

        # Load web page for scrapping url version
        $web = Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $urlVersion

        # Grab version
        $regEx = '(version\":")((?:\d+\.)+(?:\d+))"'
        $version = ('{0}.{1}' -f $DownloadId, ($web | Select-String -Pattern $regEx).Matches.Groups[2].Value)

        # Carve JSON from script tag
        $web = $web.Content | Select-String -Pattern $JSONBlobPattern | Select-Object -ExpandProperty Matches | ForEach-Object { $_.Groups['JSObject'].Value } | Select-Object -First 1 | ConvertFrom-Json
        $href = $web.dlcDetailsView.downloadFile | Where-Object { $_.url -like '*.exe' } | Select-Object -First 1

        # Return Evergreen object
        return @{ Version = $version; URI = $href.url }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxVisualStudio {
    <#
    .SYNOPSIS
        Process Microsoft Visual Studio Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxVisualStudio
    $ProductName = 'Microsoft Visual Studio'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            $null = (New-Item -Path $TempFolder -ItemType Directory -Force)
            Push-Location $TempFolder
            tar -xf $OutFile
            Pop-Location

            # copy
            $SourceAdmx = "$($TempFolder)\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxVSCode {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Visual Studio Code ADMX files
    #>

    try {
        $latest = Invoke-RestMethod -Uri 'https://update.code.visualstudio.com/api/update/win32-x64-archive/stable/latest'

        # Grab version
        $Version = $latest.productVersion
        # Grab uri
        $URI = $latest.url

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxVSCode {
    <#
    .SYNOPSIS
        Process Visual Studio Code Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxVSCode
    $ProductName = 'Microsoft VS Code'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            # locate policies folder and flatten locales if present
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter '*.admx' | Select-Object -First 1).DirectoryName
            $localesFolder = Join-Path $SourceAdmx 'locales'
            if (Test-Path -Path $localesFolder) {
                Get-ChildItem -Path $localesFolder | ForEach-Object {
                    Copy-Item -Path $_.FullName -Destination (Join-Path $SourceAdmx $_.Name) -Recurse -Force
                }
            }

            # copy
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmx1Password {
    <#
    .SYNOPSIS
        Returns url and latest Version for 1Password ADMX policy definition files.
    #>

    try {
        $url = 'https://support.1password.com/mobile-device-management/?windows='

        # Grab content
        $web = (Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing).Content

        # Extract ADMX templates zip url from support page
        $regEx = '(https\:\/\/c\.1password\.com\/dist\/1P\/win8\/1Password-admx-templates-[^\s"''<>]+\.zip)'
        $URI = (Select-String -Pattern $regEx -InputObject $web).Matches.Groups[1].Value
        if ([string]::IsNullOrWhiteSpace($URI)) {
            throw 'Unable to locate 1Password ADMX templates download link on the support page.'
        }

        # Prefer version embedded in filename (e.g. 8.12.4)
        $VersionMatch = Select-String -Pattern '1Password-admx-templates-(\d+(?:\.\d+)+)' -InputObject $URI
        if ($VersionMatch) {
            $Version = $VersionMatch.Matches.Groups[1].Value
        } else {
            $LastModifiedDate = (Resolve-Uri -Uri $URI).LastModified
            [version]$Version = $LastModifiedDate.ToString('yyyy.MM.dd')
        }

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmx1Password {
    <#
    .SYNOPSIS
        Process 1Password Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmx1Password
    $ProductName = '1Password'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            # copy (zip layout: 1Password Policies\PolicyDefinitions)
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter '1Password.admx' | Select-Object -First 1).DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxSlack {
    <#
    .SYNOPSIS
        Returns url and latest Slack policy definition files.
    #>

    try {
        $url = 'https://slack.com/help/articles/11906214948755-Manage-desktop-app-configurations'

        # Grab content
        $web = (Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing -DisableKeepAlive).Content

        # Extract url from ADMX download string (Zendesk attachment; path may include locale e.g. /hc/en-us/)
        $regEx = '(https\:\/\/slack\.zendesk\.com\/hc\/[^\s"'']+article_attachments\/[^\s"''<>]+)'
        $URI = (Select-String -Pattern $regEx -InputObject $web).Matches.Groups[1].Value
        if ([string]::IsNullOrWhiteSpace($URI)) {
            $attachmentId = (Select-String -Pattern 'article_attachments/(\d+)' -InputObject $web).Matches.Groups[1].Value
            if (-not $attachmentId) { throw 'Unable to locate Slack Group Policy Object template download link.' }
            $URI = "https://slack.zendesk.com/hc/en-us/article_attachments/$attachmentId"
        }

        # Grab version
        $LastModifiedDate = (Resolve-Uri -Uri $URI).LastModified
        [version]$Version = $LastModifiedDate.ToString('yyyy.MM.dd')

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxSlack {
    <#
    .SYNOPSIS
        Process Slack Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxSlack
    $ProductName = 'Slack'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process (Zendesk attachment has no .zip extension in URL)
        $OutFile = "$($WorkingDirectory)\downloads\SlackGPO.zip"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            # copy
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'slack.admx' | Select-Object -First 1).DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxTeamViewer {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for TeamViewer ADMX files
    #>

    try {
        # Define github repo
        $repo = 'systmworks/TeamViewer-ADMX'
        # Grab latest commit on main
        $latestCommit = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo)/commits/main" -UseBasicParsing

        # Grab version
        $Version = $latestCommit.sha.Substring(0, 7)
        # Grab uri
        $URI = "https://github.com/$($repo)/archive/refs/heads/main.zip"

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxTeamViewer {
    <#
    .SYNOPSIS
        Process TeamViewer Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxTeamViewer
    $ProductName = 'TeamViewer'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or $Evergreen.Version -ne $Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            # copy
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'TeamViewer.admx' | Select-Object -First 1).DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxAdobeDC {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Adobe DC ADMX files
    #>

    try {
        # Define github repo
        $repo = 'systmworks/Adobe-DC-ADMX'
        # Grab latest release properties
        $latest = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo)/releases/latest" -UseBasicParsing

        # Grab version
        $Version = ($latest.tag_name -replace '^v', '')
        # Grab uri
        $URI = ($latest.assets | Where-Object { $_.name -like '*.zip' }).browser_download_url

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxAdobeDC {
    <#
    .SYNOPSIS
        Process Adobe DC Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxAdobeDC
    $ProductName = 'Adobe DC'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            # Community package ships AdobeDC.adml at root; normalize into en-US for Copy-Admx
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter '*.admx' | Select-Object -First 1).DirectoryName
            $langFolder = Join-Path $SourceAdmx 'en-US'
            if (-not (Test-Path -Path $langFolder)) {
                $null = New-Item -Path $langFolder -ItemType Directory -Force
                Get-ChildItem -Path $SourceAdmx -Filter '*.adml' -File | ForEach-Object {
                    Move-Item -Path $_.FullName -Destination $langFolder -Force
                }
            }

            # copy
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxSecurityAdmx {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Security ADMX files
    #>

    try {
        # Define github repo
        $repo = 'Harvester57/Security-ADMX'
        # Grab latest tag
        $latestTag = (Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo)/tags?per_page=1" -UseBasicParsing)[0]

        # Grab version
        $Version = ($latestTag.name -replace '^v', '')
        # Grab uri
        $URI = "https://github.com/$($repo)/archive/refs/tags/$($latestTag.name).zip"

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxSecurityAdmx {
    <#
    .SYNOPSIS
        Process Security Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxSecurityAdmx
    $ProductName = 'Security ADMX'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            # copy
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Directory | Select-Object -First 1).FullName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxSchannel {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Schannel ADMX files
    #>

    try {
        $repo = 'Crosse/SchannelGroupPolicy'
        $latestCommit = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo)/commits/master" -UseBasicParsing

        $Version = $latestCommit.sha.Substring(0, 7)
        $URI = "https://github.com/$($repo)/archive/refs/heads/master.zip"

        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxSchannel {
    <#
    .SYNOPSIS
        Process Schannel Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxSchannel
    $ProductName = 'Schannel'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or $Evergreen.Version -ne $Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $OutFile = "$($WorkingDirectory)\downloads\SchannelGroupPolicy-master.zip"
        try {
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Directory -Filter 'template' | Select-Object -First 1).FullName
            if (-not $SourceAdmx) {
                throw "Unable to locate 'template' folder in Schannel Group Policy archive."
            }

            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxDellCommandUpdate {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Dell Command Update ADMX files via winget.
    #>

    try {
        $PackageId = 'Dell.CommandUpdate.Universal'

        if (-not (Get-Command -Name winget -ErrorAction SilentlyContinue)) {
            throw 'winget is required to download Dell Command Update ADMX files.'
        }

        # Resolve latest package version from winget
        $show = & winget show --id $PackageId --exact --accept-source-agreements 2>&1 | Out-String
        if ($show -notmatch '(?m)^Version:\s*(.+)$') {
            throw "Unable to determine version for winget package '$PackageId'."
        }
        $Version = $Matches[1].Trim()

        # Return evergreen object (URI is a winget package reference; download happens in Invoke-)
        return @{ Version = $Version; URI = "winget:$PackageId"; PackageId = $PackageId }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxDellCommandUpdate {
    <#
    .SYNOPSIS
        Process Dell Command Update Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxDellCommandUpdate
    $ProductName = 'Dell Command Update'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # Older GitHub-based entries used yyyy.MM.dd versioning; treat those as stale
    $isStaleGithubVersion = $false
    if ($Version) {
        try { $isStaleGithubVersion = ([version]$Version).Major -ge 2000 } catch { $isStaleGithubVersion = $true }
    }

    # see if this is a newer version
    if (-not $Version -or $isStaleGithubVersion -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $DownloadDir = Join-Path -Path "$($WorkingDirectory)\downloads" -ChildPath 'DellCommandUpdate'
        $TempFolder = Join-Path -Path $env:TEMP -ChildPath $ProductName
        try {
            if (Test-Path -Path $DownloadDir) { Remove-Item -Path $DownloadDir -Recurse -Force }
            $null = New-Item -Path $DownloadDir -ItemType Directory -Force
            if (Test-Path -Path $TempFolder) { Remove-Item -Path $TempFolder -Recurse -Force }
            $null = New-Item -Path $TempFolder -ItemType Directory -Force

            # Download latest installer with winget (no install)
            Write-Verbose "Downloading '$($Evergreen.PackageId)' with winget to '$($DownloadDir)'"
            $wingetArgs = @(
                'download'
                '--id', $Evergreen.PackageId
                '--exact'
                '--download-directory', $DownloadDir
                '--skip-dependencies'
                '--accept-package-agreements'
                '--accept-source-agreements'
            )
            $wingetOutput = & winget @wingetArgs 2>&1 | Out-String
            if ($LASTEXITCODE -ne 0) {
                throw "winget download failed for '$($Evergreen.PackageId)': $wingetOutput"
            }

            $Installer = Get-ChildItem -Path $DownloadDir -Filter '*.exe' -File -Recurse |
                Where-Object { $_.FullName -notmatch '\\Dependencies\\' } |
                Sort-Object Length -Descending |
                Select-Object -First 1
            if (-not $Installer) {
                throw "No Dell Command Update installer found in '$($DownloadDir)'."
            }

            # Prefer 7-Zip extraction (ADMX live under Templates\); fallback to Dell silent extract
            $sevenZip = @(
                (Join-Path -Path ${env:ProgramFiles} -ChildPath '7-Zip\7z.exe')
                (Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath '7-Zip\7z.exe')
                (Get-Command -Name '7z' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source)
            ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1

            if ($sevenZip) {
                Write-Verbose "Extracting '$($Installer.FullName)' with 7-Zip to '$($TempFolder)'"
                $null = & $sevenZip @('x', $Installer.FullName, "-o$TempFolder", '-y')
            } else {
                Write-Verbose "Extracting '$($Installer.FullName)' with Dell /passthrough to '$($TempFolder)'"
                $extractArgs = "/passthrough /X /B`"$TempFolder`""
                $process = Start-Process -FilePath $Installer.FullName -ArgumentList $extractArgs -Wait -PassThru -WindowStyle Hidden
                if ($process.ExitCode -notin @(0, $null)) {
                    throw "Dell Command Update extract failed with exit code $($process.ExitCode). Install 7-Zip or extract manually."
                }
            }

            $SourceAdmx = Join-Path -Path $TempFolder -ChildPath 'Templates'
            if (-not (Test-Path -Path $SourceAdmx) -or -not (Get-ChildItem -Path $SourceAdmx -Filter '*.admx' -File -ErrorAction SilentlyContinue)) {
                throw "Dell Command Update Templates folder with ADMX files was not found under '$($TempFolder)'."
            }

            # copy
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force -ErrorAction SilentlyContinue

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxWingetAutoUpdate {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Winget-AutoUpdate ADMX files
    #>

    try {
        # Define github repo
        $repo = 'Romanitho/Winget-AutoUpdate'
        # Grab latest release properties
        $latest = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo)/releases/latest" -UseBasicParsing

        # Grab version
        $Version = ($latest.tag_name -replace '^v', '')
        # Grab uri
        $URI = "https://github.com/$($repo)/archive/refs/tags/$($latest.tag_name).zip"

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxWingetAutoUpdate {
    <#
    .SYNOPSIS
        Process Winget-AutoUpdate Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxWingetAutoUpdate
    $ProductName = 'Winget-AutoUpdate'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            # copy
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'WAU.admx' | Select-Object -First 1).DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxWingetAutoUpdateIntune {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Winget-AutoUpdate-Intune ADMX files
    #>

    try {
        # Define github repo
        $repo = 'Weatherlights/Winget-AutoUpdate-Intune'
        # Grab latest commit on main
        $latestCommit = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo)/commits/main" -UseBasicParsing

        # Grab version
        $Version = $latestCommit.sha.Substring(0, 7)
        # Grab uri
        $URI = "https://github.com/$($repo)/archive/refs/heads/main.zip"

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxWingetAutoUpdateIntune {
    <#
    .SYNOPSIS
        Process Winget-AutoUpdate-Intune Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxWingetAutoUpdateIntune
    $ProductName = 'Winget-AutoUpdate-Intune'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or $Evergreen.Version -ne $Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            # copy
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'WinGet-AutoUpdate-Configurator.admx' | Select-Object -First 1).DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxPSAppDeployToolkit {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for PSAppDeployToolkit ADMX files
    #>

    try {
        # Define github repo
        $repo = 'PSAppDeployToolkit/PSAppDeployToolkit'
        # Grab latest release properties
        $latest = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo)/releases/latest" -UseBasicParsing

        # Grab version
        $Version = ($latest.tag_name -replace '^v', '')
        # Grab uri
        $URI = ($latest.assets | Where-Object { $_.name -eq 'PSAppDeployToolkit_ModuleOnly.zip' }).browser_download_url

        # Return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxPSAppDeployToolkit {
    <#
    .SYNOPSIS
        Process PSAppDeployToolkit Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxPSAppDeployToolkit
    $ProductName = 'PSAppDeployToolkit'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            # copy
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter '*.admx' | Select-Object -First 1).DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxDevolutionsRDM {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Devolutions Remote Desktop Manager ADMX files
    #>

    try {
        $url = 'https://devolutions.net/remote-desktop-manager/download/'

        # Grab content
        $web = (Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing).Content

        # Extract download urls and pick newest version
        $rdmBinLinks = [regex]::Matches($web, 'https://[^"''\s]+Devolutions\.RemoteDesktopManager\.Bin\.(\d+(?:\.\d+)+)\.zip')
        $latest = $rdmBinLinks | ForEach-Object {
            [PSCustomObject]@{
                Version = [version]$_.Groups[1].Value
                URI     = $_.Value
            }
        } | Sort-Object -Property Version -Descending | Select-Object -First 1

        # Return evergreen object
        return @{ Version = $latest.Version; URI = $latest.URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxDevolutionsRDM {
    <#
    .SYNOPSIS
        Process Devolutions Remote Desktop Manager Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxDevolutionsRDM
    $ProductName = 'Devolutions Remote Desktop Manager'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            # copy
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Directory -Filter 'Policies' | Select-Object -First 1).FullName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-EvergreenAdmxPowerToy {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Microsoft PowerToys ADMX files
    #>

    try {
        $repo = 'microsoft/PowerToys'
        $latest = ((Invoke-WebRequest -UseDefaultCredentials -Uri "https://api.github.com/repos/$($repo)/releases" -UseBasicParsing | ConvertFrom-Json) |
            Where-Object { $_.draft -eq $false -and $_.prerelease -eq $false -and $_.assets.browser_download_url -match 'GroupPolicyObjectFiles' })[0]

        $Version = ($latest.tag_name | Select-String -Pattern '(\d+(\.\d+){1,4})' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        $URI = $latest.assets.browser_download_url | Where-Object { $_ -like '*/GroupPolicyObjectFiles*.zip' } | Select-Object -First 1

        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxPowerToy {
    <#
    .SYNOPSIS
        Process Microsoft PowerToys Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxPowerToy
    $ProductName = 'Microsoft PowerToys'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'PowerToys.admx' | Select-Object -First 1).DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxWindowsTerminal {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Windows Terminal ADMX files
    #>

    try {
        $repo = 'microsoft/terminal'
        $latest = ((Invoke-WebRequest -UseDefaultCredentials -Uri "https://api.github.com/repos/$($repo)/releases" -UseBasicParsing | ConvertFrom-Json) |
            Where-Object { $_.draft -eq $false -and $_.prerelease -eq $false -and $_.assets.browser_download_url -match 'GroupPolicyTemplates' })[0]

        $Version = ($latest.tag_name | Select-String -Pattern '(\d+(\.\d+){1,4})' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        $URI = $latest.assets.browser_download_url | Where-Object { $_ -like '*/GroupPolicyTemplates*.zip' } | Select-Object -First 1

        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxWindowsTerminal {
    <#
    .SYNOPSIS
        Process Windows Terminal Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxWindowsTerminal
    $ProductName = 'Windows Terminal'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'WindowsTerminal.admx' | Select-Object -First 1).DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxThunderbird {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Mozilla Thunderbird ADMX files
    #>

    try {
        $repo = 'thunderbird/policy-templates'
        $latestCommit = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo)/commits/master" -UseBasicParsing

        $Version = $latestCommit.sha.Substring(0, 7)
        $URI = "https://github.com/$($repo)/archive/refs/heads/master.zip"

        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxThunderbird {
    <#
    .SYNOPSIS
        Process Mozilla Thunderbird Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxThunderbird
    $ProductName = 'Mozilla Thunderbird'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or $Evergreen.Version -ne $Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $OutFile = "$($WorkingDirectory)\downloads\thunderbird-policy-templates-master.zip"
        try {
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'thunderbird.admx' | Select-Object -First 1).DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxDropbox {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Dropbox ADMX files
    #>

    try {
        $repo = 'dropbox/GPO-Templates'
        $latestCommit = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo)/commits/main" -UseBasicParsing

        $Version = $latestCommit.sha.Substring(0, 7)
        $URI = "https://github.com/$($repo)/archive/refs/heads/main.zip"

        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxDropbox {
    <#
    .SYNOPSIS
        Process Dropbox Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxDropbox
    $ProductName = 'Dropbox'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or $Evergreen.Version -ne $Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $OutFile = "$($WorkingDirectory)\downloads\Dropbox-GPO-Templates-main.zip"
        try {
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            # Repo ships Dropbox.adml at root; normalize into en-US for Copy-Admx
            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'Dropbox.admx' | Select-Object -First 1).DirectoryName
            $langFolder = Join-Path $SourceAdmx 'en-US'
            if (-not (Test-Path -Path $langFolder)) {
                $null = New-Item -Path $langFolder -ItemType Directory -Force
                Get-ChildItem -Path $SourceAdmx -Filter '*.adml' -File | ForEach-Object {
                    Move-Item -Path $_.FullName -Destination $langFolder -Force
                }
            }

            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxFoxit {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Foxit PDF Reader/Editor ADMX files
    #>

    try {
        # Parent directory listing is 403; probe newest-first known version folders
        $candidates = @(
            '2026.2.0', '2026.1.5', '2026.1.4', '2026.1.3', '2026.1.2', '2026.1.1', '2026.1.0',
            '2025.4.0', '2025.3.0', '2025.2.0', '2025.1.0',
            '2024.4.0', '2024.3.0', '2024.2.2', '2024.2.0', '2024.1.0'
        )

        $latest = $null
        foreach ($ver in $candidates) {
            $base = "https://cdn01.foxitsoftware.com/product/phantomPDF/desktop/win/$ver/tools"
            $editorUri = "$base/Foxit%20PDF%20Editor_enu_admx%26adml.zip"
            try {
                $null = Invoke-WebRequest -UseDefaultCredentials -Uri $editorUri -Method Head -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
                $latest = [PSCustomObject]@{ Version = $ver; URI = $editorUri; Base = $base }
                break
            } catch {
                Write-Verbose "Foxit ADMX probe missed version '$ver': $($_.Exception.Message)"
            }
        }

        if (-not $latest) {
            throw 'Unable to locate a Foxit PDF ADMX tools package on the Foxit CDN.'
        }

        $readerUri = "$($latest.Base)/Foxit%20PDF%20Reader_enu_admx%26adml.zip"
        return @{ Version = $latest.Version; URI = $latest.URI; ReaderURI = $readerUri }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxFoxit {
    <#
    .SYNOPSIS
        Process Foxit PDF Reader and Editor Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxFoxit
    $ProductName = 'Foxit PDF'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $TempFolder = "$($env:TEMP)\$($ProductName)"
        $StagingFolder = Join-Path -Path $TempFolder -ChildPath 'staging'
        try {
            if (Test-Path -Path $TempFolder) { Remove-Item -Path $TempFolder -Recurse -Force }
            $null = New-Item -Path $StagingFolder -ItemType Directory -Force
            $null = New-Item -Path (Join-Path $StagingFolder 'en-US') -ItemType Directory -Force

            foreach ($item in @(
                    @{ Name = 'FoxitPDFEditor.zip'; URI = $Evergreen.URI }
                    @{ Name = 'FoxitPDFReader.zip'; URI = $Evergreen.ReaderURI }
                )) {
                $OutFile = "$($WorkingDirectory)\downloads\$($item.Name)"
                Write-Verbose "Downloading '$($item.URI)' to '$($OutFile)'"
                Invoke-FileDownload -Uri $item.URI -OutFile $OutFile -UseDefaultCredentials

                $ExtractPath = Join-Path -Path $TempFolder -ChildPath ([IO.Path]::GetFileNameWithoutExtension($item.Name))
                Expand-Archive -Path $OutFile -DestinationPath $ExtractPath -Force

                # Zips ship .admx and .adml in the same folder; normalize for Copy-Admx
                Get-ChildItem -Path $ExtractPath -Recurse -Filter '*.admx' -File | ForEach-Object {
                    Copy-Item -Path $_.FullName -Destination $StagingFolder -Force
                }
                Get-ChildItem -Path $ExtractPath -Recurse -Filter '*.adml' -File | ForEach-Object {
                    Copy-Item -Path $_.FullName -Destination (Join-Path $StagingFolder 'en-US') -Force
                }
            }

            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $StagingFolder -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxLibreOffice {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for LibreOffice / Collabora Office ADMX files
    #>

    try {
        $repo = 'CollaboraOnline/ADMX'
        $latestCommit = Invoke-RestMethod -Uri "https://api.github.com/repos/$($repo)/commits/master" -UseBasicParsing

        $Version = $latestCommit.sha.Substring(0, 7)
        $URI = "https://github.com/$($repo)/archive/refs/heads/master.zip"

        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxLibreOffice {
    <#
    .SYNOPSIS
        Process LibreOffice / Collabora Office Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxLibreOffice
    $ProductName = 'LibreOffice'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or $Evergreen.Version -ne $Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $OutFile = "$($WorkingDirectory)\downloads\CollaboraOnline-ADMX-master.zip"
        try {
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $TempFolder = "$($env:TEMP)\$($ProductName)"
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'Collabora-Office.admx' | Select-Object -First 1).DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxHPAnyware {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for HP Anyware (PCoIP Agent) ADMX files
    #>

    try {
        if (-not (Get-Command -Name winget -ErrorAction SilentlyContinue)) {
            throw 'winget is required to resolve the latest HP Anyware version.'
        }

        # Client package version tracks the same YY.MM.P train as the Standard Agent
        $show = & winget show --id HP.AnywarePCoIPClient --exact --accept-source-agreements 2>&1 | Out-String
        if ($show -notmatch '(?m)^Version:\s*(.+)$') {
            throw 'Unable to determine version for winget package HP.AnywarePCoIPClient.'
        }
        $wingetVersion = $Matches[1].Trim()

        # Map 26.5.3.0 -> 26.05.3 for the public agent download path
        $parts = $wingetVersion.Split('.')
        if ($parts.Count -lt 3) {
            throw "Unexpected HP Anyware winget version format: $wingetVersion"
        }
        $agentVersion = '{0}.{1:D2}.{2}' -f [int]$parts[0], [int]$parts[1], [int]$parts[2]
        $URI = "https://dl.anyware.hp.com/DeAdBCiUYInHcSTy/pcoip-agent/raw/names/pcoip-agent-standard-exe/versions/$agentVersion/pcoip-agent-standard_$agentVersion.exe"

        return @{ Version = $agentVersion; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxHPAnyware {
    <#
    .SYNOPSIS
        Process HP Anyware (PCoIP) Admx files from the Standard Agent installer
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxHPAnyware
    $ProductName = 'HP Anyware'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $OutFile = "$($WorkingDirectory)\downloads\pcoip-agent-standard_$($Evergreen.Version).exe"
        $TempFolder = "$($env:TEMP)\$($ProductName)"
        try {
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            if (Test-Path -Path $TempFolder) { Remove-Item -Path $TempFolder -Recurse -Force }
            $null = New-Item -Path $TempFolder -ItemType Directory -Force

            $sevenZip = @(
                (Join-Path -Path ${env:ProgramFiles} -ChildPath '7-Zip\7z.exe')
                (Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath '7-Zip\7z.exe')
                (Get-Command -Name '7z' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source)
            ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1

            if (-not $sevenZip) {
                throw '7-Zip is required to extract HP Anyware ADMX files from the Standard Agent installer.'
            }

            Write-Verbose "Extracting ADMX from '$($OutFile)' with 7-Zip to '$($TempFolder)'"
            $null = & $sevenZip @('x', $OutFile, "-o$TempFolder", '*.admx', '*.adml', '-r', '-y')

            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'PCoIP.admx' |
                Where-Object { $_.FullName -match 'configuration\\policyDefinitions' } |
                Select-Object -First 1).DirectoryName
            if (-not $SourceAdmx) {
                $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'PCoIP.admx' | Select-Object -First 1).DirectoryName
            }
            if (-not $SourceAdmx) {
                throw "PCoIP.admx was not found inside the HP Anyware Standard Agent installer."
            }

            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force -ErrorAction SilentlyContinue
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxSpecops {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Specops Authentication Client ADMX files
    #>

    try {
        $URI = 'https://download.specopssoft.com/Release/Client/Specops.Client.AdmxTemplates.zip'
        $AzureAdURI = 'https://download.specopssoft.com/Release/Client/Specops.Client.AzureAdJoinedComputer.AdmxTemplates.zip'

        $TempZip = Join-Path -Path $env:TEMP -ChildPath 'Specops.Client.AdmxTemplates.zip'
        Invoke-FileDownload -Uri $URI -OutFile $TempZip -UseDefaultCredentials

        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($TempZip)
        try {
            $entry = $zip.Entries | Where-Object { $_.Name -eq 'version.txt' } | Select-Object -First 1
            if (-not $entry) {
                throw 'version.txt was not found in Specops.Client.AdmxTemplates.zip'
            }
            $reader = New-Object System.IO.StreamReader($entry.Open())
            try {
                $Version = $reader.ReadToEnd().Trim()
            } finally {
                $reader.Close()
            }
        } finally {
            $zip.Dispose()
        }

        Remove-Item -Path $TempZip -Force -ErrorAction SilentlyContinue

        if ([string]::IsNullOrWhiteSpace($Version)) {
            throw 'Unable to read Specops ADMX version from version.txt'
        }

        return @{ Version = $Version; URI = $URI; AzureAdURI = $AzureAdURI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxSpecops {
    <#
    .SYNOPSIS
        Process Specops Authentication Client Admx files (on-prem + Entra ID)
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxSpecops
    $ProductName = 'Specops Authentication Client'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $TempFolder = "$($env:TEMP)\$($ProductName)"
        $StagingFolder = Join-Path -Path $TempFolder -ChildPath 'staging'
        try {
            if (Test-Path -Path $TempFolder) { Remove-Item -Path $TempFolder -Recurse -Force }
            $null = New-Item -Path $StagingFolder -ItemType Directory -Force
            $null = New-Item -Path (Join-Path $StagingFolder 'en-US') -ItemType Directory -Force

            foreach ($item in @(
                    @{ Name = 'Specops.Client.AdmxTemplates.zip'; URI = $Evergreen.URI }
                    @{ Name = 'Specops.Client.AzureAdJoinedComputer.AdmxTemplates.zip'; URI = $Evergreen.AzureAdURI }
                )) {
                $OutFile = "$($WorkingDirectory)\downloads\$($item.Name)"
                Write-Verbose "Downloading '$($item.URI)' to '$($OutFile)'"
                Invoke-FileDownload -Uri $item.URI -OutFile $OutFile -UseDefaultCredentials

                $ExtractPath = Join-Path -Path $TempFolder -ChildPath ([IO.Path]::GetFileNameWithoutExtension($item.Name))
                Expand-Archive -Path $OutFile -DestinationPath $ExtractPath -Force

                Get-ChildItem -Path $ExtractPath -Recurse -Filter '*.admx' -File | ForEach-Object {
                    Copy-Item -Path $_.FullName -Destination $StagingFolder -Force
                }
                Get-ChildItem -Path $ExtractPath -Recurse -Filter '*.adml' -File | ForEach-Object {
                    $langDir = $_.Directory.Name
                    $destLang = Join-Path $StagingFolder $langDir
                    if (-not (Test-Path -Path $destLang)) {
                        $null = New-Item -Path $destLang -ItemType Directory -Force
                    }
                    Copy-Item -Path $_.FullName -Destination $destLang -Force
                }
            }

            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $StagingFolder -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxWSL {
    <#
    .SYNOPSIS
        Returns latest Version (commit SHA) for Microsoft WSL Intune ADMX files
    #>

    try {
        $commits = Invoke-RestMethod -Uri 'https://api.github.com/repos/microsoft/WSL/commits?path=intune&per_page=1' -Headers @{ 'User-Agent' = 'EvergreenAdmx' } -UseBasicParsing
        $Version = $commits[0].sha.Substring(0, 7)
        return @{ Version = $Version; URI = 'https://api.github.com/repos/microsoft/WSL/contents/intune?ref=master' }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxWSL {
    <#
    .SYNOPSIS
        Process Microsoft WSL Intune Admx files from GitHub
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxWSL
    $ProductName = 'WSL'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or $Evergreen.Version -ne $Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $TempFolder = "$($env:TEMP)\$($ProductName)"
        $StagingFolder = Join-Path -Path $TempFolder -ChildPath 'staging'
        try {
            if (Test-Path -Path $TempFolder) { Remove-Item -Path $TempFolder -Recurse -Force }
            $null = New-Item -Path $StagingFolder -ItemType Directory -Force

            function Save-WSLGitHubContents {
                param(
                    [string]$ApiUri,
                    [string]$TargetRoot
                )

                $entries = Invoke-RestMethod -Uri $ApiUri -Headers @{ 'User-Agent' = 'EvergreenAdmx' } -UseBasicParsing
                foreach ($entry in $entries) {
                    if ($entry.type -eq 'dir') {
                        $subDir = Join-Path -Path $TargetRoot -ChildPath $entry.name
                        if (-not (Test-Path -Path $subDir)) {
                            $null = New-Item -Path $subDir -ItemType Directory -Force
                        }
                        Save-WSLGitHubContents -ApiUri $entry.url -TargetRoot $subDir
                    } elseif ($entry.type -eq 'file' -and $entry.name -match '\.(admx|adml)$') {
                        $outFile = Join-Path -Path $TargetRoot -ChildPath $entry.name
                        Write-Verbose "Downloading '$($entry.download_url)' to '$($outFile)'"
                        Invoke-FileDownload -Uri $entry.download_url -OutFile $outFile -UseDefaultCredentials
                    }
                }
            }

            Save-WSLGitHubContents -ApiUri $Evergreen.URI -TargetRoot $StagingFolder

            if (-not (Get-ChildItem -Path $StagingFolder -Filter 'WSL.admx' -File -ErrorAction SilentlyContinue)) {
                throw 'WSL.admx was not found in the microsoft/WSL intune folder.'
            }

            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $StagingFolder -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxLenovoDockManager {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Lenovo Dock Manager ADMX files
    #>

    try {
        $URI = 'https://download.lenovo.com/consumer/options/policy_setup.exe'
        $TempExe = Join-Path -Path $env:TEMP -ChildPath 'policy_setup.exe'
        Invoke-FileDownload -Uri $URI -OutFile $TempExe -UseDefaultCredentials

        $info = (Get-Item -LiteralPath $TempExe).VersionInfo
        $Version = $null
        if ($info.ProductName -match 'Package\s+(\d+(?:\.\d+)+)') {
            $Version = $Matches[1]
        } elseif (-not [string]::IsNullOrWhiteSpace($info.FileVersion)) {
            $Version = ($info.FileVersion).Trim()
        }

        Remove-Item -Path $TempExe -Force -ErrorAction SilentlyContinue

        if ([string]::IsNullOrWhiteSpace($Version)) {
            throw 'Unable to determine Lenovo Dock Manager ADMX package version from policy_setup.exe'
        }

        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxLenovoDockManager {
    <#
    .SYNOPSIS
        Process Lenovo Dock Manager Admx files from policy_setup.exe
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxLenovoDockManager
    $ProductName = 'Lenovo Dock Manager'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $OutFile = "$($WorkingDirectory)\downloads\policy_setup.exe"
        $TempFolder = "$($env:TEMP)\$($ProductName)"
        try {
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            if (Test-Path -Path $TempFolder) { Remove-Item -Path $TempFolder -Recurse -Force }
            $null = New-Item -Path $TempFolder -ItemType Directory -Force

            Write-Verbose "Extracting Inno Setup package to '$($TempFolder)'"
            $proc = Start-Process -FilePath $OutFile -ArgumentList @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', "/DIR=`"$TempFolder`"") -Wait -PassThru
            if ($proc.ExitCode -ne 0 -and $null -ne $proc.ExitCode) {
                Write-Verbose "Inno extract exit code: $($proc.ExitCode)"
            }

            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter '*.admx' -File -ErrorAction SilentlyContinue |
                Select-Object -First 1).DirectoryName
            if (-not $SourceAdmx) {
                $SourceAdmx = (Get-ChildItem -Path "$env:SystemRoot\PolicyDefinitions" -Filter '*Dock*' -File -ErrorAction SilentlyContinue |
                    Select-Object -First 1).DirectoryName
            }
            if (-not $SourceAdmx) {
                throw 'Lenovo Dock Manager ADMX files were not found after extracting policy_setup.exe'
            }

            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force -ErrorAction SilentlyContinue
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxPDFXChange {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for PDF-XChange ADMX files
    #>

    try {
        $URI = 'https://www.pdf-xchange.com/Tracker%5FAD%5FAdministrativeTemplates.zip'
        $LastModifiedDate = (Resolve-Uri -Uri $URI).LastModified
        [version]$Version = $LastModifiedDate.ToString('yyyy.MM.dd')
        return @{ Version = $Version.ToString(); URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxPDFXChange {
    <#
    .SYNOPSIS
        Process PDF-XChange Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxPDFXChange
    $ProductName = 'PDF-XChange'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $OutFile = "$($WorkingDirectory)\downloads\Tracker_AD_AdministrativeTemplates.zip"
        $TempFolder = "$($env:TEMP)\$($ProductName)"
        try {
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            if (Test-Path -Path $TempFolder) { Remove-Item -Path $TempFolder -Recurse -Force }
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'PDFXEditor.admx' | Select-Object -First 1).DirectoryName
            if (-not $SourceAdmx) {
                throw 'PDFXEditor.admx was not found in the PDF-XChange administrative templates zip.'
            }

            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxRealVNC {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for RealVNC Connect Server/Viewer ADMX files
    #>

    try {
        $ServerURI = 'https://downloads.realvnc.com/download/file/policy.files/RealVNC-server-admx-templates-Latest.zip'
        $ViewerURI = 'https://downloads.realvnc.com/download/file/policy.files/RealVNC-viewer-admx-templates-Latest.zip'

        $serverLm = (Resolve-Uri -Uri $ServerURI).LastModified
        $viewerLm = (Resolve-Uri -Uri $ViewerURI).LastModified
        $Latest = if ($serverLm -ge $viewerLm) { $serverLm } else { $viewerLm }
        [version]$Version = $Latest.ToString('yyyy.MM.dd')

        return @{ Version = $Version.ToString(); URI = $ServerURI; ViewerURI = $ViewerURI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxRealVNC {
    <#
    .SYNOPSIS
        Process RealVNC Connect Server and Viewer Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxRealVNC
    $ProductName = 'RealVNC Connect'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $TempFolder = "$($env:TEMP)\$($ProductName)"
        $StagingFolder = Join-Path -Path $TempFolder -ChildPath 'staging'
        try {
            if (Test-Path -Path $TempFolder) { Remove-Item -Path $TempFolder -Recurse -Force }
            $null = New-Item -Path $StagingFolder -ItemType Directory -Force
            $null = New-Item -Path (Join-Path $StagingFolder 'en-US') -ItemType Directory -Force

            # Viewer first, then Server so newer shared realvnc.admx/adml win
            foreach ($item in @(
                    @{ Name = 'RealVNC-viewer-admx-templates-Latest.zip'; URI = $Evergreen.ViewerURI }
                    @{ Name = 'RealVNC-server-admx-templates-Latest.zip'; URI = $Evergreen.URI }
                )) {
                $OutFile = "$($WorkingDirectory)\downloads\$($item.Name)"
                Write-Verbose "Downloading '$($item.URI)' to '$($OutFile)'"
                Invoke-FileDownload -Uri $item.URI -OutFile $OutFile -UseDefaultCredentials

                $ExtractPath = Join-Path -Path $TempFolder -ChildPath ([IO.Path]::GetFileNameWithoutExtension($item.Name))
                Expand-Archive -Path $OutFile -DestinationPath $ExtractPath -Force

                $AdmxRoot = Join-Path -Path $ExtractPath -ChildPath 'admx'
                if (-not (Test-Path -Path $AdmxRoot)) {
                    $AdmxRoot = $ExtractPath
                }

                Get-ChildItem -Path $AdmxRoot -Filter '*.admx' -File | ForEach-Object {
                    Copy-Item -Path $_.FullName -Destination $StagingFolder -Force
                }
                Get-ChildItem -Path $AdmxRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                    $destLang = Join-Path $StagingFolder $_.Name
                    if (-not (Test-Path -Path $destLang)) {
                        $null = New-Item -Path $destLang -ItemType Directory -Force
                    }
                    Get-ChildItem -Path $_.FullName -Filter '*.adml' -File | ForEach-Object {
                        Copy-Item -Path $_.FullName -Destination $destLang -Force
                    }
                }
            }

            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $StagingFolder -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxABBYYFineReader {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for ABBYY FineReader PDF ADMX files
    #>

    try {
        $pageUrl = 'https://help.abbyy.com/en-us/finereader/16/admin_guide/gpo_domain/'
        $admxUri = 'https://support.abbyy.com/hc/en-us/article_attachments/8277004669075/FineReader16.admx'
        $admlUri = 'https://support.abbyy.com/hc/en-us/article_attachments/8277017934739/FineReader16.adml'

        try {
            $web = (Invoke-WebRequest -UseDefaultCredentials -Uri $pageUrl -UseBasicParsing).Content
            $admxMatch = [regex]::Match($web, 'https://support\.abbyy\.com/hc/[^"''\s<>]+FineReader16\.admx')
            $admlMatch = [regex]::Match($web, 'https://support\.abbyy\.com/hc/[^"''\s<>]+FineReader16\.adml')
            if ($admxMatch.Success) { $admxUri = $admxMatch.Value }
            if ($admlMatch.Success) { $admlUri = $admlMatch.Value }
        } catch {
            Write-Verbose "ABBYY page scrape failed; using known attachment URLs: $($_.Exception.Message)"
        }

        $LastModifiedDate = (Resolve-Uri -Uri $admxUri).LastModified
        [version]$Version = $LastModifiedDate.ToString('yyyy.MM.dd')
        return @{ Version = $Version.ToString(); URI = $admxUri; AdmlURI = $admlUri }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxABBYYFineReader {
    <#
    .SYNOPSIS
        Process ABBYY FineReader PDF Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxABBYYFineReader
    $ProductName = 'ABBYY FineReader PDF'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $TempFolder = "$($env:TEMP)\$($ProductName)"
        $StagingFolder = Join-Path -Path $TempFolder -ChildPath 'staging'
        try {
            if (Test-Path -Path $TempFolder) { Remove-Item -Path $TempFolder -Recurse -Force }
            $null = New-Item -Path (Join-Path $StagingFolder 'en-US') -ItemType Directory -Force

            $OutAdmx = "$($WorkingDirectory)\downloads\FineReader16.admx"
            $OutAdml = "$($WorkingDirectory)\downloads\FineReader16.adml"
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutAdmx)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutAdmx -UseDefaultCredentials
            Write-Verbose "Downloading '$($Evergreen.AdmlURI)' to '$($OutAdml)'"
            Invoke-FileDownload -Uri $Evergreen.AdmlURI -OutFile $OutAdml -UseDefaultCredentials

            Copy-Item -Path $OutAdmx -Destination $StagingFolder -Force
            Copy-Item -Path $OutAdml -Destination (Join-Path $StagingFolder 'en-US') -Force

            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $StagingFolder -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxAdminByRequest {
    <#
    .SYNOPSIS
        Returns latest Version (content hash) and Uri for Admin By Request ADMX files
    #>

    try {
        $URI = 'https://www.adminbyrequest.com/ADMX'
        $TempZip = Join-Path -Path $env:TEMP -ChildPath 'AdminByRequest-Admx.zip'
        Invoke-FileDownload -Uri $URI -OutFile $TempZip -UseDefaultCredentials

        $hash = (Get-FileHash -Path $TempZip -Algorithm SHA256).Hash.ToLowerInvariant()
        $Version = $hash.Substring(0, 12)
        Remove-Item -Path $TempZip -Force -ErrorAction SilentlyContinue

        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxAdminByRequest {
    <#
    .SYNOPSIS
        Process Admin By Request Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxAdminByRequest
    $ProductName = 'Admin By Request'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or $Evergreen.Version -ne $Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $OutFile = "$($WorkingDirectory)\downloads\AdminByRequest-Admx.zip"
        $TempFolder = "$($env:TEMP)\$($ProductName)"
        try {
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            if (Test-Path -Path $TempFolder) { Remove-Item -Path $TempFolder -Recurse -Force }
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'AdminByRequest.admx' | Select-Object -First 1).DirectoryName
            if (-not $SourceAdmx) {
                throw 'AdminByRequest.admx was not found in the Admin By Request ADMX zip.'
            }

            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

function Get-EvergreenAdmxGoTo {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for GoTo app ADMX files
    #>

    try {
        $URI = 'https://goto-desktop.goto.com/GoToAppAdministrativeTemplates.zip'
        $LastModifiedDate = (Resolve-Uri -Uri $URI).LastModified
        [version]$Version = $LastModifiedDate.ToString('yyyy.MM.dd')
        return @{ Version = $Version.ToString(); URI = $URI }
    } catch {
        Throw $_
    }
}

function Invoke-EvergreenAdmxGoTo {
    <#
    .SYNOPSIS
        Process GoTo app Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxGoTo
    $ProductName = 'GoTo'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        $OutFile = "$($WorkingDirectory)\downloads\GoToAppAdministrativeTemplates.zip"
        $TempFolder = "$($env:TEMP)\$($ProductName)"
        try {
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            if (Test-Path -Path $TempFolder) { Remove-Item -Path $TempFolder -Recurse -Force }
            Expand-Archive -Path $OutFile -DestinationPath $TempFolder -Force

            $SourceAdmx = (Get-ChildItem -Path $TempFolder -Recurse -Filter 'GoTo.admx' | Select-Object -First 1).DirectoryName
            if (-not $SourceAdmx) {
                throw 'GoTo.admx was not found in GoToAppAdministrativeTemplates.zip'
            }

            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            Remove-Item -Path $TempFolder -Recurse -Force
            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        return $null
    }
}

# Download functions
function Invoke-EvergreenAdmxWindows {
    <#
    .SYNOPSIS
        Download Windows Admx policy definitions files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files

    .PARAMETER WindowsFeatureVersion
        Official WindowsFeatureVersion format

    .PARAMETER WindowsVersion
        Differentiate between Windows 10 and Windows 11

    .PARAMETER Languages
        Languages to check
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string]$WindowsFeatureVersion,
        [int]$WindowsVersion,
        [string[]]$Languages = $null
    )

    If ($WindowsVersion -eq 11 -or $WindowsVersion -eq 10) {
        $id = Get-WindowsDownloadId -WindowsVersion $WindowsVersion -WindowsFeatureVersion $WindowsFeatureVersion
        $ProductName = "Microsoft Windows $($WindowsVersion) $($WindowsFeatureVersion)"
    } elseif ($WindowsVersion -eq '2022' -or $WindowsVersion -eq '2025') {
        $id = Get-WindowsDownloadId -WindowsVersion $WindowsVersion
        $ProductName = "Microsoft Windows Server $($WindowsVersion)"
    }
    $Evergreen = Get-EvergreenAdmxWindows -DownloadId $id
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }
    $TempFolder = "$($env:TEMP)\$($ProductName)"

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # install
            Write-Verbose "Installing downloaded Windows $($WindowsVersion) Admx installer"
            $null = Start-Process -FilePath 'MsiExec.exe' -WorkingDirectory "$($WorkingDirectory)\downloads" -ArgumentList "/qn /norestart /a `"$($OutFile.split('\')[-1])`" TargetDir=`"$($TempFolder)`"" -PassThru -Wait

            # find installation path
            Write-Verbose "Grabbing installation path for Windows $($WindowsVersion) Admx installer"
            $InstallFolder = Get-ChildItem -Path "$($TempFolder)\Microsoft Group Policy"
            Write-Verbose "Found '$($InstallFolder.Name)'"

            # copy
            $SourceAdmx = "$($TempFolder)\Microsoft Group Policy\$($InstallFolder.Name)\PolicyDefinitions"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxEdge {
    <#
    .SYNOPSIS
        Process Microsoft Edge Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxEdge
    $ProductName = 'Microsoft Edge'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($ProductName).cab"
        $ZipFile = "$($WorkingDirectory)\downloads\MicrosoftEdgePolicyTemplates.zip"

        try {
            # Download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # Extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $null = (New-Item -Path "$($env:TEMP)\$($ProductName)" -ItemType Directory -Force)
            $null = (expand -F:* "$($OutFile)" "$($env:TEMP)\$($ProductName)" $ZipFile)
            Expand-Archive -Path $ZipFile -DestinationPath "$($env:TEMP)\$($ProductName)" -Force

            # Copy
            $SourceAdmx = "$($env:TEMP)\$($ProductName)\windows\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # Cleanup
            Remove-Item -Path $OutFile -Force
            Remove-Item -Path "$env:TEMP\$($ProductName)" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxOneDrive {
    <#
    .SYNOPSIS
        Process OneDrive Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files

    .PARAMETER PreferLocalOneDrive
        Prefer policy definitions from installed local version of MicrosoftOneDrive. If not specified, Microsoft OneDrive will be installed to extract the policy definitions.
    #>

    [CmdletBinding()]
    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [switch]$PreferLocalOneDrive,
        [string[]]$Languages = $null
    )

    if ($PreferLocalOneDrive) {
        $Evergreen = Get-EvergreenAdmxOneDrive -PreferLocalOneDrive
    } else {
        $Evergreen = Get-EvergreenAdmxOneDrive
    }

    $ProductName = 'Microsoft OneDrive'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        try {
            if (-not $PreferLocalOneDrive) {

                # Set the output file
                $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"

                # Download
                Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
                Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

                # Install
                Write-Verbose 'Installing downloaded OneDrive installer'
                $null = Start-Process -FilePath $OutFile -ArgumentList '/allusers /silent' -PassThru
                # Wait for setup to complete
                while (Get-Process -Name 'OneDriveSetup' -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
                # Check if OneDrive is running and close it if it is
                Write-Verbose 'Checking if OneDrive is running and stopping it if necessary'
                $process = Get-Process -Name 'OneDrive' -ErrorAction SilentlyContinue
                if ($process) {
                    Write-Verbose 'OneDrive process is running. Stopping it...'
                    try {
                        $process | Stop-Process -Force
                        Write-Verbose 'OneDrive process stopped successfully'
                    } catch {
                        Write-Warning "Failed to stop OneDrive process: $_"
                    }
                } else {
                    Write-Verbose 'No OneDrive process found running'
                }

                # Find uninstall info
                Write-Verbose 'Grabbing uninstallation info from registry for OneDrive installer'
                $uninstall = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\OneDriveSetup.exe'
                if ($null -eq $uninstall) {
                    $uninstall = Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\OneDriveSetup.exe'
                }
                if ($null -eq $uninstall) {
                    Write-Warning -Message 'Unable to find uninstall information for OneDrive.'
                } else {
                    Write-Verbose "Found '$($uninstall.DisplayName)'"
                    # Find OneDrive ADMX folder
                    Write-Verbose 'Grabbing installation path for OneDrive installer'
                    $installfolder = $uninstall.DisplayIcon.Substring(0, $uninstall.DisplayIcon.IndexOf('\OneDriveSetup.exe'))
                    Write-Verbose "Found '$($installfolder)'"
                }
            } else {
                $installfolder = $Evergreen.ADMXFolder
            }

            # Copy
            $SourceAdmx = "$($installfolder)\adm"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            if (-not (Test-Path -Path "$($TargetAdmx)")) { $null = (New-Item -Path "$($TargetAdmx)" -ItemType Directory -Force) }
            if ($PolicyStore -and (Test-Path -Path "$($SourceAdmx)\*.admx")) {
                Write-Verbose "Copying Admx files from '$($SourceAdmx)' to '$($PolicyStore)'"
                Copy-Item -Path "$($SourceAdmx)\*.admx" -Destination "$($PolicyStore)" -Force
                foreach ($language in $Languages) {
                    if (-not (Test-Path -Path "$($SourceAdmx)\$($language)") -and -not (Test-Path -Path "$($SourceAdmx)\$($language.Substring(0,2))")) {
                        if (-not (Test-Path -Path "$($PolicyStore)en-US")) { $null = (New-Item -Path "$($PolicyStore)en-US" -ItemType Directory -Force) }
                        Copy-Item -Path "$($SourceAdmx)\*.adml" -Destination "$($PolicyStore)en-US" -Force
                    } else {
                        $sourcelanguage = $language; if (-not (Test-Path -Path "$($SourceAdmx)\$($language)")) { $sourcelanguage = $language.Substring(0, 2) }
                        if (-not (Test-Path -Path "$($PolicyStore)$($language)")) { $null = (New-Item -Path "$($PolicyStore)$($language)" -ItemType Directory -Force) }
                        Copy-Item -Path "$($SourceAdmx)\$($sourcelanguage)\*.adml" -Destination "$($PolicyStore)$($language)" -Force
                    }
                }
            } elseIf (Test-Path -Path "$($SourceAdmx)\*.admx") {
                Write-Verbose "Copying Admx files from '$($SourceAdmx)' to '$($TargetAdmx)'"
                Copy-Item -Path "$($SourceAdmx)\*.admx" -Destination "$($TargetAdmx)" -Force
                foreach ($language in $Languages) {
                    if (-not (Test-Path -Path "$($SourceAdmx)\$($language)") -and -not (Test-Path -Path "$($SourceAdmx)\$($language.Substring(0,2))")) {
                        if ($language -notlike 'en-us') { Write-Warning "Language '$($language)' not found for '$($ProductName)'. Processing 'en-US' instead." }
                        if (-not (Test-Path -Path "$($TargetAdmx)\en-US")) { $null = (New-Item -Path "$($TargetAdmx)\en-US" -ItemType Directory -Force) }
                        Copy-Item -Path "$($SourceAdmx)\*.adml" -Destination "$($TargetAdmx)\en-US" -Force
                    } else {
                        $sourcelanguage = $language; if (-not (Test-Path -Path "$($SourceAdmx)\$($language)")) { $sourcelanguage = $language.Substring(0, 2) }
                        if (-not (Test-Path -Path "$($TargetAdmx)\$($language)")) { $null = (New-Item -Path "$($TargetAdmx)\$($language)" -ItemType Directory -Force) }
                        Copy-Item -Path "$($SourceAdmx)\$($sourcelanguage)\*.adml" -Destination "$($TargetAdmx)\$($language)" -Force
                    }
                }
            } else {
                Write-Warning "No ADMX files found for '$($ProductName)'"
            }

            if ($StampAdmxRevision) {
                $normalizedRevision = ConvertTo-AdmxRevisionString -Version $Evergreen.Version
                if ($normalizedRevision) {
                    if (Test-Path -LiteralPath $TargetAdmx) {
                        Set-AdmxRevision -Path $TargetAdmx -Revision $normalizedRevision
                    }
                    if ($PolicyStore) {
                        Set-AdmxRevision -Path $PolicyStore -Revision $normalizedRevision
                    }
                }
            }

            if (-not $PreferLocalOneDrive) {
                # Uninstall
                Write-Verbose 'Uninstalling Microsoft OneDrive installer'
                $null = Start-Process -FilePath "$($installfolder)\OneDriveSetup.exe" -ArgumentList '/uninstall /allusers' -PassThru -Wait
            }

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # Version already processed
        return $null
    }
}

function Invoke-EvergreenAdmx365App {
    <#
    .SYNOPSIS
        Download Microsoft 365 Apps policy definition files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files

    .PARAMETER Architecture
        Architecture (x86 or x64)
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string]$Architecture = 'x64',
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmx365App | Where-Object { $_.Architecture -like $Architecture }
    $ProductName = "Microsoft 365 Apps $($Architecture)"
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # See if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # Download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # Download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # Extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\office'"
            $null = Start-Process -FilePath $OutFile -ArgumentList "/quiet /norestart /extract:`"$($env:TEMP)\office`"" -PassThru -Wait

            # Copy
            $SourceAdmx = "$($env:TEMP)\office\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # Cleanup
            Remove-Item -Path "$($env:TEMP)\office" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # Version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxFSLogix {
    <#
    .SYNOPSIS
        Process Microsoft FSLogix Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null
    )

    $Evergreen = Get-EvergreenAdmxFSLogix
    $ProductName = 'Microsoft FSLogix'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\$($ProductName)" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\$($ProductName)"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            if (-not (Test-Path -Path "$($TargetAdmx)\en-US")) { $null = (New-Item -Path "$($TargetAdmx)\en-US" -ItemType Directory -Force) }

            Write-Verbose "Copying Admx files from '$($SourceAdmx)' to '$($TargetAdmx)'"
            Copy-Item -Path "$($SourceAdmx)\*.admx" -Destination "$($TargetAdmx)" -Force
            Copy-Item -Path "$($SourceAdmx)\*.adml" -Destination "$($TargetAdmx)\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from '$($SourceAdmx)' to '$($PolicyStore)'"
                Copy-Item -Path "$($SourceAdmx)\*.admx" -Destination "$($PolicyStore)" -Force
                if (-not (Test-Path -Path "$($PolicyStore)en-US")) { $null = (New-Item -Path "$($PolicyStore)en-US" -ItemType Directory -Force) }
                Copy-Item -Path "$($SourceAdmx)\*.adml" -Destination "$($PolicyStore)en-US" -Force
            }

            if ($StampAdmxRevision) {
                $normalizedRevision = ConvertTo-AdmxRevisionString -Version $Evergreen.Version
                if ($normalizedRevision) {
                    Set-AdmxRevision -Path $TargetAdmx -Revision $normalizedRevision
                    if ($PolicyStore) {
                        Set-AdmxRevision -Path $PolicyStore -Revision $normalizedRevision
                    }
                }
            }

            # cleanup
            Remove-Item -Path "$($env:TEMP)\$($ProductName)" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # Version already processed
        return $null
    }
}

Function Invoke-EvergreenAdmxChrome {
    <#
    .SYNOPSIS
        Download Google Chrome policy definition files.

    .PARAMETER Version
        Get version of Google Chrome policy definition files.

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxChrome
    $ProductName = 'Google Chrome'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # See if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"
        $OutFile = "$($WorkingDirectory)\downloads\googlechromeadmx.zip"

        try {
            # Download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # Extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\chromeadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\chromeadmx" -Force

            # Copy
            $SourceAdmx = "$($env:TEMP)\chromeadmx\windows\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # Cleanup
            Remove-Item -Path "$($env:TEMP)\chromeadmx" -Recurse -Force

            # Chrome update admx is a separate download
            $url = 'https://dl.google.com/dl/update2/enterprise/googleupdateadmx.zip'

            # Download
            $OutFile = "$($WorkingDirectory)\downloads\googlechromeupdateadmx.zip"
            Write-Verbose "Downloading '$($url)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $url -OutFile $OutFile -UseDefaultCredentials

            # Extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\chromeupdateadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\chromeupdateadmx" -Force

            # Copy
            $SourceAdmx = "$($env:TEMP)\chromeupdateadmx\GoogleUpdateAdmx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Quiet -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # Cleanup
            Remove-Item -Path "$($env:TEMP)\chromeupdateadmx" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # Version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxAdobeAcrobat {
    <#
    .SYNOPSIS
        Process Adobe Acrobat Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null
    )

    $Evergreen = Get-EvergreenAdmxAdobeAcrobat
    $ProductName = 'Adobe Acrobat'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\$($ProductName)" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\$($ProductName)"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path "$($env:TEMP)\$($ProductName)" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxAdobeReader {
    <#
    .SYNOPSIS
        Process Adobe Reader Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null
    )

    $Evergreen = Get-EvergreenAdmxAdobeReader
    $ProductName = 'Adobe Reader'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\AdobeReader'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\AdobeReader" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\AdobeReader"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path "$($env:TEMP)\AdobeReader" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxWorkspaceApp {
    <#
    .SYNOPSIS
        Process Citrix Workspace App Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxWorkplaceApp
    $ProductName = 'Citrix Workspace App'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('?')[0].Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\citrixworkspaceapp'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\citrixworkspaceapp" -Force

            # copy
            # $SourceAdmx = "$($env:TEMP)\citrixworkspaceapp\$($Evergreen.URI.Split("/")[-2].Split("?")[0].SubString(0,$Evergreen.URI.Split("/")[-2].Split("?")[0].IndexOf(".")))"
            $SourceAdmx = (Get-ChildItem -Path "$($env:TEMP)\citrixworkspaceapp\$($Evergreen.URI.Split('/')[-2].Split('?')[0].SubString(0,$Evergreen.URI.Split('/')[-2].Split('?')[0].IndexOf('.')))" -Include '*.admx' -Recurse)[0].DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path "$($env:TEMP)\citrixworkspaceapp" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxFirefox {
    <#
    .SYNOPSIS
        Process Mozilla Firefox Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxFirefox
    $ProductName = 'Mozilla Firefox'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\firefoxadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\firefoxadmx" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\firefoxadmx\windows"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path "$($env:TEMP)\firefoxadmx" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxZoom {
    <#
    .SYNOPSIS
        Process Zoom Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null
    )

    $Evergreen = Get-EvergreenAdmxZoom
    $ProductName = 'Zoom'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\$($ProductName)" -Force

            # cleanup folder structure
            $SourceAdmx = Get-ChildItem -Path "$($env:TEMP)\$($ProductName)\" -Exclude *.adm -Include *.admx -Recurse | Where-Object { -Not $_.PSIsContainer }
            $SourceAdml = Get-ChildItem -Path "$($env:TEMP)\$($ProductName)\" -Exclude *.adm -Include *.adml -Recurse | Where-Object { -Not $_.PSIsContainer }
            $null = (New-Item -Path "$($env:TEMP)\clean-$($ProductName)\" -ItemType Directory -Force)
            $null = (New-Item -Path "$($env:TEMP)\clean-$($ProductName)\en-us" -ItemType Directory -Force)
            Copy-Item -Path $SourceAdmx -Destination "$($env:TEMP)\clean-$($ProductName)\" -Force
            Copy-Item -Path $SourceAdml -Destination "$($env:TEMP)\clean-$($ProductName)\en-us" -Force
            Remove-Item -Path "$($env:TEMP)\$($ProductName)" -Recurse -Force
            Rename-Item -Path "$($env:TEMP)\clean-$($ProductName)" -NewName "$($ProductName)" -Force
            # copy
            $SourceAdmx = "$($env:TEMP)\$($ProductName)"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path "$($env:TEMP)\$($ProductName)" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxZoomVDI {
    <#
    .SYNOPSIS
        Process Zoom VDI Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null
    )

    $Evergreen = Get-EvergreenAdmxZoomVDI
    $ProductName = 'Zoom VDI'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\$($ProductName)" -Force

            # cleanup folder structure
            $SourceAdmx = Get-ChildItem -Path "$($env:TEMP)\$($ProductName)\" -Exclude *.adm -Include *.admx -Recurse | Where-Object { -Not $_.PSIsContainer }
            $SourceAdml = Get-ChildItem -Path "$($env:TEMP)\$($ProductName)\" -Exclude *.adm -Include *.adml -Recurse | Where-Object { -Not $_.PSIsContainer }
            $null = (New-Item -Path "$($env:TEMP)\clean-$($ProductName)\" -ItemType Directory -Force)
            $null = (New-Item -Path "$($env:TEMP)\clean-$($ProductName)\en-us" -ItemType Directory -Force)
            Copy-Item -Path $SourceAdmx -Destination "$($env:TEMP)\clean-$($ProductName)\" -Force
            Copy-Item -Path $SourceAdml -Destination "$($env:TEMP)\clean-$($ProductName)\en-us" -Force
            Remove-Item -Path "$($env:TEMP)\$($ProductName)" -Recurse -Force
            Rename-Item -Path "$($env:TEMP)\clean-$($ProductName)" -NewName "$($ProductName)" -Force
            # copy
            $SourceAdmx = "$($env:TEMP)\$($ProductName)"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path "$($env:TEMP)\$($ProductName)" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxBISF {
    <#
    .SYNOPSIS
        Process BIS-F Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null
    )

    $Evergreen = Get-EvergreenAdmxBISF
    $ProductName = 'BIS-F'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\bis-f.$($Evergreen.Version).zip"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\bisfadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\bisfadmx" -Force

            # find extraction folder
            Write-Verbose 'Finding extraction folder'
            $folder = (Get-ChildItem -Path "$($env:TEMP)\bisfadmx" | Sort-Object LastWriteTime -Descending)[0].Name

            # copy
            $SourceAdmx = "$($env:TEMP)\bisfadmx\$($folder)\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path "$($env:TEMP)\bisfadmx" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxCustomPolicy {
    <#
    .SYNOPSIS
        Process Custom Policy Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string]$CustomPolicyStore,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-CustomPolicyOnline -CustomPolicyStore $CustomPolicyStore
    $ProductName = 'Custom Policy Store'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        try {
            # copy
            $SourceAdmx = "$($Evergreen.URI)"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName "$($ProductName)" -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxAvd {
    <#
    .SYNOPSIS
    Process Microsoft AVD Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxAVD
    $ProductName = 'Microsoft AVD'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($ProductName).cab"
        $ZipFile = "$($WorkingDirectory)\downloads\AVDGPTemplate.zip"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            #Invoke-Download -URL $Evergreen.URI -Destination "$($WorkingDirectory)\downloads" -FileName "$($ProductName).cab"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $null = (New-Item -Path "$($env:TEMP)\$($ProductName)" -ItemType Directory -Force)
            $null = (expand -F:* "$($OutFile)" "$($env:TEMP)\$($ProductName)" $ZipFile)
            Expand-Archive -Path $ZipFile -DestinationPath "$($env:TEMP)\$($ProductName)" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\$($ProductName)"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path $OutFile -Force
            Remove-Item -Path "$($env:TEMP)\$($ProductName)" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxWinget {
    <#
    .SYNOPSIS
        Process Winget-cli Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxWinget
    $ProductName = 'Microsoft Winget'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\wingetadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\wingetadmx" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\wingetadmx\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path "$($env:TEMP)\wingetadmx" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Invoke-EvergreenAdmxBrave {
    <#
    .SYNOPSIS
        Process Brave Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-EvergreenAdmxBrave
    $ProductName = 'Brave Browser'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-FileDownload -Uri $Evergreen.URI -OutFile $OutFile -UseDefaultCredentials

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\braveadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\braveadmx" -Force

            # fix policyNamespaces to support Intune ingest
            Write-Verbose 'Fixing policyNamespaces in brave.admx'
            [xml]$xml = (Get-Content -Path "$($env:TEMP)\braveadmx\windows\admx\brave.admx") -replace 'Brave:Cat_Brave', 'brave:Cat_Brave' | Where-Object { $_ -notmatch '^\s*<using' }
            $newCategory = $xml.CreateElement('category')
            $newCategory.SetAttribute('displayName', '$(string.brave)')
            $newCategory.SetAttribute('name', 'Cat_Brave')
            $xml.policyDefinitions.categories.AppendChild($newCategory)
            $xml.Save("$($env:TEMP)\braveadmx\windows\admx\brave.admx")

            # copy
            $SourceAdmx = "$($env:TEMP)\braveadmx\windows\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages -Revision $(if ($StampAdmxRevision) { $Evergreen.Version })

            # cleanup
            Remove-Item -Path "$($env:TEMP)\braveadmx" -Recurse -Force

            return $Evergreen
        } catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

# Helper function to update ADMX versions
function Update-AdmxVersion {
    <#
    .SYNOPSIS
        Updates the ADMX versions object with new version information

    .PARAMETER AdmxVersions
        The ADMX versions object to update

    .PARAMETER ProductKey
        The product key to update

    .PARAMETER AdmxData
        The new ADMX data
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AdmxVersions,

        [Parameter(Mandatory = $true)]
        [string]$ProductKey,

        [Parameter(Mandatory = $false)]
        $AdmxData
    )

    if ($null -ne $AdmxData) {
        if ($PSCmdlet.ShouldProcess("$ProductKey ADMX version", 'Update')) {
            # Check if AdmxVersions.Value is null and initialize if needed
            if ($null -eq $AdmxVersions.Value) {
                $AdmxVersions.Value = @{}
            }

            # Check if the product key exists
            if ($AdmxVersions.Value.ContainsKey($ProductKey)) {
                $AdmxVersions.Value.$ProductKey = @{
                    Version = $AdmxData.Version
                    URI = $AdmxData.URI
                }
            } else {
                # Add new product key if it doesn't exist
                $AdmxVersions.Value += @{
                    $ProductKey = @{
                        Version = $AdmxData.Version
                        URI = $AdmxData.URI
                    }
                }
            }
        }
    }
}
#endregion

#region execution
# Ensure Central Policy Store exists when explicitly requested
if ($PolicyStoreSpecified) {
    Initialize-PolicyStore -PolicyStore $PolicyStore -Languages $Languages
}

# Cleanup-only mode skips downloads
if ($CleanPolicyStoreOnly) {
    Write-Verbose "`nCleanPolicyStoreOnly: skipping Admx downloads"
    $null = Clear-ObsoleteAdmx -PolicyStore $PolicyStore -Languages $Languages
    return
}

# Custom Policy Store
if ($Include -notcontains 'Custom Policy Store') {
    Write-Verbose "`nSkipping Custom Policy Store"
} else {
    Write-Verbose "`nProcessing Admx files for Custom Policy Store"
    $currentversion = $null
    if ($AdmxVersions.PSObject.properties -match 'CustomPolicyStore') { $currentversion = $AdmxVersions.CustomPolicyStore.Version }
    $admx = Invoke-EvergreenAdmxCustomPolicy -Version $currentversion -PolicyStore $PolicyStore -CustomPolicyStore $CustomPolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'CustomPolicyStore' -AdmxData $admx
}

# Windows 10
if ($Include -notcontains 'Windows 10') {
    Write-Verbose "`nSkipping Windows 10"
} else {
    Write-Verbose "`nProcessing Admx files for Windows 10 $($WindowsFeatureVersion)"
    $pkey = 'Windows10'
    $admx = Invoke-EvergreenAdmxWindows -Version $AdmxVersions[$pkey].Version -PolicyStore $PolicyStore -WindowsFeatureVersion $WindowsFeatureVersion -WindowsVersion 10 -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey $pkey -AdmxData $admx
}

# Windows 11
if ($Include -notcontains 'Windows 11') {
    Write-Verbose "`nSkipping Windows 11"
} else {
    Write-Verbose "`nProcessing Admx files for Windows 11 $($WindowsFeatureVersion)"
    $pkey = 'Windows11'
    $admx = Invoke-EvergreenAdmxWindows -Version $AdmxVersions[$pkey].Version -PolicyStore $PolicyStore -WindowsFeatureVersion $WindowsFeatureVersion -WindowsVersion 11 -Languages $Languages
    if ($null -ne $admx) {
        Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey $pkey -AdmxData $admx
    } else {
        Write-Warning 'Failed to retrieve Windows 11 ADMX files. Skipping update.'
    }
}

# Windows 2022
if ($Include -notcontains 'Windows 2022') {
    Write-Verbose "`nSkipping Windows Server 2022"
} else {
    Write-Verbose "`nProcessing Admx files for Windows Server 2022"
    $pkey = 'Windows2022'
    $admx = Invoke-EvergreenAdmxWindows -Version $AdmxVersions[$pkey].Version -PolicyStore $PolicyStore -WindowsVersion 2022 -Languages $Languages
    if ($null -ne $admx) {
        Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey $pkey -AdmxData $admx
    } else {
        Write-Warning 'Failed to retrieve Windows Server 2022 ADMX files. Skipping update.'
    }
}

# Windows 2025
if ($Include -notcontains 'Windows 2025') {
    Write-Verbose "`nSkipping Windows Server 2025"
} else {
    Write-Verbose "`nProcessing Admx files for Windows Server 2025"
    $pkey = 'Windows2025'
    $admx = Invoke-EvergreenAdmxWindows -Version $AdmxVersions[$pkey].Version -PolicyStore $PolicyStore -WindowsVersion 2025 -Languages $Languages
    if ($null -ne $admx) {
        Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey $pkey -AdmxData $admx
    } else {
        Write-Warning 'Failed to retrieve Windows Server 2025 ADMX files. Skipping update.'
    }
}

# Microsoft Edge
if ($Include -notcontains 'Microsoft Edge') {
    Write-Verbose "`nSkipping Microsoft Edge"
} else {
    Write-Verbose "`nProcessing Admx files for Microsoft Edge"
    $admx = Invoke-EvergreenAdmxEdge -Version $AdmxVersions.Edge.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Edge' -AdmxData $admx
}

# Microsoft OneDrive
if ($Include -notcontains 'Microsoft OneDrive') {
    Write-Verbose "`nSkipping Microsoft OneDrive"
} else {
    Write-Verbose "`nProcessing Admx files for Microsoft OneDrive"
    If ($PreferLocalOneDrive) {
        $admx = Invoke-EvergreenAdmxOneDrive -Version $AdmxVersions.OneDrive.Version -PolicyStore $PolicyStore -PreferLocalOneDrive -Languages $Languages
    } else {
        $admx = Invoke-EvergreenAdmxOneDrive -Version $AdmxVersions.OneDrive.Version -PolicyStore $PolicyStore -Languages $Languages
    }
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'OneDrive' -AdmxData $admx
}

# Microsoft 365
if ($Include -notcontains 'Microsoft 365 Apps') {
    Write-Verbose "`nSkipping Microsoft 365 Apps"
} else {
    Write-Verbose "`nProcessing Admx files for Microsoft 365 Apps"
    $admx = Invoke-EvergreenAdmx365App -Version $AdmxVersions['365Apps'].Version -PolicyStore $PolicyStore -Architecture 'x64' -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey '365Apps' -AdmxData $admx
}

# Microsoft FSLogix
if ($Include -notcontains 'Microsoft FSLogix') {
    Write-Verbose "`nSkipping Microsoft FSLogix"
} else {
    Write-Verbose "`nProcessing Admx files for Microsoft FSLogix"
    $admx = Invoke-EvergreenAdmxFSLogix -Version $AdmxVersions.FSLogix.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'FSLogix' -AdmxData $admx
}

# Adobe Acrobat
if ($Include -notcontains 'Adobe Acrobat') {
    Write-Verbose "`nSkipping Adobe Acrobat"
} else {
    Write-Verbose "`nProcessing Admx files for Adobe Acrobat"
    $admx = Invoke-EvergreenAdmxAdobeAcrobat -Version $AdmxVersions.AdobeAcrobat.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'AdobeAcrobat' -AdmxData $admx
}

# Adobe Reader
if ($Include -notcontains 'Adobe Reader') {
    Write-Verbose "`nSkipping Adobe Reader"
} else {
    Write-Verbose "`nProcessing Admx files for Adobe Reader"
    $admx = Invoke-EvergreenAdmxAdobeReader -Version $AdmxVersions.AdobeReader.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'AdobeReader' -AdmxData $admx
}

# BIS-F
if ($Include -notcontains 'BIS-F') {
    Write-Verbose "`nSkipping BIS-F"
} else {
    Write-Verbose "`nProcessing Admx files for BIS-F"
    $admx = Invoke-EvergreenAdmxBISF -Version $AdmxVersions.BISF.Version -PolicyStore $PolicyStore
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'BISF' -AdmxData $admx
}

# Citrix Workspace App
if ($Include -notcontains 'Citrix Workspace App') {
    Write-Verbose "`nSkipping Citrix Workspace App"
} else {
    Write-Verbose "`nProcessing Admx files for Citrix Workspace App"
    $admx = Invoke-EvergreenAdmxWorkspaceApp -Version $AdmxVersions.CitrixWorkspaceApp.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'CitrixWorkspaceApp' -AdmxData $admx
}

# Google Chrome
if ($Include -notcontains 'Google Chrome') {
    Write-Verbose "`nSkipping Google Chrome"
} else {
    Write-Verbose "`nProcessing Admx files for Google Chrome"
    $admx = Invoke-EvergreenAdmxChrome -Version $AdmxVersions.GoogleChrome.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'GoogleChrome' -AdmxData $admx
}

# Mozilla Firefox
if ($Include -notcontains 'Mozilla Firefox') {
    Write-Verbose "`nSkipping Mozilla Firefox"
} else {
    Write-Verbose "`nProcessing Admx files for Mozilla Firefox"
    $admx = Invoke-EvergreenAdmxFirefox -Version $AdmxVersions.MozillaFirefox.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'MozillaFirefox' -AdmxData $admx
}

# Zoom
if ($Include -notcontains 'Zoom') {
    Write-Verbose "`nSkipping Zoom"
} else {
    Write-Verbose "`nProcessing Admx files for Zoom"
    $admx = Invoke-EvergreenAdmxZoom -Version $AdmxVersions.Zoom.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Zoom' -AdmxData $admx
}

# Zoom VDI
if ($Include -notcontains 'Zoom VDI') {
    Write-Verbose "`nSkipping Zoom VDI"
} else {
    Write-Verbose "`nProcessing Admx files for Zoom VDI"
    $pkey = 'Zoom VDI'
    $admx = Invoke-EvergreenAdmxZoomVDI -Version $AdmxVersions[$pkey].Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey $pkey -AdmxData $admx
}

# Microsoft AVD
if ($Include -notcontains 'Microsoft AVD') {
    Write-Verbose "`nSkipping Microsoft Azure Virtual Desktop"
} else {
    Write-Verbose "`nProcessing Admx files for Microsoft Azure Virtual Desktop"
    $admx = Invoke-EvergreenAdmxAvd -Version $AdmxVersions.AzureVirtualDesktop.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'AzureVirtualDesktop' -AdmxData $admx
}

# Microsoft Winget
if ($Include -notcontains 'Microsoft Winget') {
    Write-Verbose "`nSkipping Microsoft Winget"
} else {
    Write-Verbose "`nProcessing Admx files for Microsoft Winget"
    $admx = Invoke-EvergreenAdmxWinget -Version $AdmxVersions.Winget.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Winget' -AdmxData $admx
}

# Brave Browser
if ($Include -notcontains 'Brave Browser') {
    Write-Verbose "`nSkipping Brave Browser"
} else {
    Write-Verbose "`nProcessing Admx files for Brave Browser"
    $admx = Invoke-EvergreenAdmxBrave -Version $AdmxVersions.Brave.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Brave' -AdmxData $admx
}

# Microsoft Notepad
if ($Include -notcontains 'Microsoft Notepad') {
    Write-Verbose "`nSkipping Microsoft Notepad"
} else {
    Write-Verbose "`nProcessing Admx files for Microsoft Notepad"
    $admx = Invoke-EvergreenAdmxNotepad -Version $AdmxVersions.Notepad.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Notepad' -AdmxData $admx
}

# Microsoft Clipchamp
if ($Include -notcontains 'Microsoft Clipchamp') {
    Write-Verbose "`nSkipping Microsoft Clipchamp"
} else {
    Write-Verbose "`nProcessing Admx files for Microsoft Clipchamp"
    $admx = Invoke-EvergreenAdmxClipchamp -Version $AdmxVersions.Clipchamp.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Clipchamp' -AdmxData $admx
}

# Microsoft Visual Studio
if ($Include -notcontains 'Microsoft Visual Studio') {
    Write-Verbose "`nSkipping Microsoft Visual Studio"
} else {
    Write-Verbose "`nProcessing Admx files for Microsoft Visual Studio"
    $admx = Invoke-EvergreenAdmxVisualStudio -Version $AdmxVersions.VisualStudio.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'VisualStudio' -AdmxData $admx
}

# Microsoft VS Code
if ($Include -notcontains 'Microsoft VS Code') {
    Write-Verbose "`nSkipping Microsoft VS Code"
} else {
    Write-Verbose "`nProcessing Admx files for Microsoft VS Code"
    $admx = Invoke-EvergreenAdmxVSCode -Version $AdmxVersions.VSCode.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'VSCode' -AdmxData $admx
}

# Slack
if ($Include -notcontains 'Slack') {
    Write-Verbose "`nSkipping Slack"
} else {
    Write-Verbose "`nProcessing Admx files for Slack"
    $admx = Invoke-EvergreenAdmxSlack -Version $AdmxVersions.Slack.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Slack' -AdmxData $admx
}

# 1Password
if ($Include -notcontains '1Password') {
    Write-Verbose "`nSkipping 1Password"
} else {
    Write-Verbose "`nProcessing Admx files for 1Password"
    $admx = Invoke-EvergreenAdmx1Password -Version $AdmxVersions.'1Password'.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey '1Password' -AdmxData $admx
}

# TeamViewer
if ($Include -notcontains 'TeamViewer') {
    Write-Verbose "`nSkipping TeamViewer"
} else {
    Write-Verbose "`nProcessing Admx files for TeamViewer"
    $admx = Invoke-EvergreenAdmxTeamViewer -Version $AdmxVersions.TeamViewer.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'TeamViewer' -AdmxData $admx
}

# Adobe DC
if ($Include -notcontains 'Adobe DC') {
    Write-Verbose "`nSkipping Adobe DC"
} else {
    Write-Verbose "`nProcessing Admx files for Adobe DC"
    $admx = Invoke-EvergreenAdmxAdobeDC -Version $AdmxVersions.AdobeDC.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'AdobeDC' -AdmxData $admx
}

# Security ADMX
if ($Include -notcontains 'Security ADMX') {
    Write-Verbose "`nSkipping Security ADMX"
} else {
    Write-Verbose "`nProcessing Admx files for Security ADMX"
    $admx = Invoke-EvergreenAdmxSecurityAdmx -Version $AdmxVersions.SecurityAdmx.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'SecurityAdmx' -AdmxData $admx
}

# Schannel
if ($Include -notcontains 'Schannel') {
    Write-Verbose "`nSkipping Schannel"
} else {
    Write-Verbose "`nProcessing Admx files for Schannel"
    $admx = Invoke-EvergreenAdmxSchannel -Version $AdmxVersions.Schannel.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Schannel' -AdmxData $admx
}

# Dell Command Update
if ($Include -notcontains 'Dell Command Update') {
    Write-Verbose "`nSkipping Dell Command Update"
} else {
    Write-Verbose "`nProcessing Admx files for Dell Command Update"
    $admx = Invoke-EvergreenAdmxDellCommandUpdate -Version $AdmxVersions.DellCommandUpdate.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'DellCommandUpdate' -AdmxData $admx
}

# Winget-AutoUpdate
if ($Include -notcontains 'Winget-AutoUpdate') {
    Write-Verbose "`nSkipping Winget-AutoUpdate"
} else {
    Write-Verbose "`nProcessing Admx files for Winget-AutoUpdate"
    $admx = Invoke-EvergreenAdmxWingetAutoUpdate -Version $AdmxVersions.WingetAutoUpdate.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'WingetAutoUpdate' -AdmxData $admx
}

# Winget-AutoUpdate-Intune
if ($Include -notcontains 'Winget-AutoUpdate-Intune') {
    Write-Verbose "`nSkipping Winget-AutoUpdate-Intune"
} else {
    Write-Verbose "`nProcessing Admx files for Winget-AutoUpdate-Intune"
    $admx = Invoke-EvergreenAdmxWingetAutoUpdateIntune -Version $AdmxVersions.WingetAutoUpdateIntune.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'WingetAutoUpdateIntune' -AdmxData $admx
}

# PSAppDeployToolkit
if ($Include -notcontains 'PSAppDeployToolkit') {
    Write-Verbose "`nSkipping PSAppDeployToolkit"
} else {
    Write-Verbose "`nProcessing Admx files for PSAppDeployToolkit"
    $admx = Invoke-EvergreenAdmxPSAppDeployToolkit -Version $AdmxVersions.PSAppDeployToolkit.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'PSAppDeployToolkit' -AdmxData $admx
}

# Devolutions Remote Desktop Manager
if ($Include -notcontains 'Devolutions Remote Desktop Manager') {
    Write-Verbose "`nSkipping Devolutions Remote Desktop Manager"
} else {
    Write-Verbose "`nProcessing Admx files for Devolutions Remote Desktop Manager"
    $admx = Invoke-EvergreenAdmxDevolutionsRDM -Version $AdmxVersions.DevolutionsRDM.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'DevolutionsRDM' -AdmxData $admx
}

# Microsoft PowerToys
if ($Include -notcontains 'Microsoft PowerToys') {
    Write-Verbose "`nSkipping Microsoft PowerToys"
} else {
    Write-Verbose "`nProcessing Admx files for Microsoft PowerToys"
    $admx = Invoke-EvergreenAdmxPowerToy -Version $AdmxVersions.PowerToys.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'PowerToys' -AdmxData $admx
}

# Windows Terminal
if ($Include -notcontains 'Windows Terminal') {
    Write-Verbose "`nSkipping Windows Terminal"
} else {
    Write-Verbose "`nProcessing Admx files for Windows Terminal"
    $admx = Invoke-EvergreenAdmxWindowsTerminal -Version $AdmxVersions.WindowsTerminal.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'WindowsTerminal' -AdmxData $admx
}

# Mozilla Thunderbird
if ($Include -notcontains 'Mozilla Thunderbird') {
    Write-Verbose "`nSkipping Mozilla Thunderbird"
} else {
    Write-Verbose "`nProcessing Admx files for Mozilla Thunderbird"
    $admx = Invoke-EvergreenAdmxThunderbird -Version $AdmxVersions.MozillaThunderbird.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'MozillaThunderbird' -AdmxData $admx
}

# Dropbox
if ($Include -notcontains 'Dropbox') {
    Write-Verbose "`nSkipping Dropbox"
} else {
    Write-Verbose "`nProcessing Admx files for Dropbox"
    $admx = Invoke-EvergreenAdmxDropbox -Version $AdmxVersions.Dropbox.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Dropbox' -AdmxData $admx
}

# Foxit PDF
if ($Include -notcontains 'Foxit PDF') {
    Write-Verbose "`nSkipping Foxit PDF"
} else {
    Write-Verbose "`nProcessing Admx files for Foxit PDF"
    $admx = Invoke-EvergreenAdmxFoxit -Version $AdmxVersions.FoxitPDF.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'FoxitPDF' -AdmxData $admx
}

# LibreOffice
if ($Include -notcontains 'LibreOffice') {
    Write-Verbose "`nSkipping LibreOffice"
} else {
    Write-Verbose "`nProcessing Admx files for LibreOffice"
    $admx = Invoke-EvergreenAdmxLibreOffice -Version $AdmxVersions.LibreOffice.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'LibreOffice' -AdmxData $admx
}

# HP Anyware
if ($Include -notcontains 'HP Anyware') {
    Write-Verbose "`nSkipping HP Anyware"
} else {
    Write-Verbose "`nProcessing Admx files for HP Anyware"
    $admx = Invoke-EvergreenAdmxHPAnyware -Version $AdmxVersions.HPAnyware.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'HPAnyware' -AdmxData $admx
}


# Specops Authentication Client
if ($Include -notcontains 'Specops Authentication Client') {
    Write-Verbose "`nSkipping Specops Authentication Client"
} else {
    Write-Verbose "`nProcessing Admx files for Specops Authentication Client"
    $admx = Invoke-EvergreenAdmxSpecops -Version $AdmxVersions.SpecopsAuthenticationClient.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'SpecopsAuthenticationClient' -AdmxData $admx
}

# WSL
if ($Include -notcontains 'WSL') {
    Write-Verbose "`nSkipping WSL"
} else {
    Write-Verbose "`nProcessing Admx files for WSL"
    $admx = Invoke-EvergreenAdmxWSL -Version $AdmxVersions.WSL.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'WSL' -AdmxData $admx
}

# Lenovo Dock Manager
if ($Include -notcontains 'Lenovo Dock Manager') {
    Write-Verbose "`nSkipping Lenovo Dock Manager"
} else {
    Write-Verbose "`nProcessing Admx files for Lenovo Dock Manager"
    $admx = Invoke-EvergreenAdmxLenovoDockManager -Version $AdmxVersions.LenovoDockManager.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'LenovoDockManager' -AdmxData $admx
}

# PDF-XChange
if ($Include -notcontains 'PDF-XChange') {
    Write-Verbose "`nSkipping PDF-XChange"
} else {
    Write-Verbose "`nProcessing Admx files for PDF-XChange"
    $admx = Invoke-EvergreenAdmxPDFXChange -Version $AdmxVersions.PDFXChange.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'PDFXChange' -AdmxData $admx
}

# RealVNC Connect
if ($Include -notcontains 'RealVNC Connect') {
    Write-Verbose "`nSkipping RealVNC Connect"
} else {
    Write-Verbose "`nProcessing Admx files for RealVNC Connect"
    $admx = Invoke-EvergreenAdmxRealVNC -Version $AdmxVersions.RealVNCConnect.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'RealVNCConnect' -AdmxData $admx
}

# ABBYY FineReader PDF
if ($Include -notcontains 'ABBYY FineReader PDF') {
    Write-Verbose "`nSkipping ABBYY FineReader PDF"
} else {
    Write-Verbose "`nProcessing Admx files for ABBYY FineReader PDF"
    $admx = Invoke-EvergreenAdmxABBYYFineReader -Version $AdmxVersions.ABBYYFineReaderPDF.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'ABBYYFineReaderPDF' -AdmxData $admx
}

# Admin By Request
if ($Include -notcontains 'Admin By Request') {
    Write-Verbose "`nSkipping Admin By Request"
} else {
    Write-Verbose "`nProcessing Admx files for Admin By Request"
    $admx = Invoke-EvergreenAdmxAdminByRequest -Version $AdmxVersions.AdminByRequest.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'AdminByRequest' -AdmxData $admx
}

# GoTo
if ($Include -notcontains 'GoTo') {
    Write-Verbose "`nSkipping GoTo"
} else {
    Write-Verbose "`nProcessing Admx files for GoTo"
    $admx = Invoke-EvergreenAdmxGoTo -Version $AdmxVersions.GoTo.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'GoTo' -AdmxData $admx
}
if ($CleanPolicyStore -and $PolicyStoreSpecified) {
    Write-Verbose "`nCleaning obsolete Admx files from Policy Store '$($PolicyStore)'"
    $null = Clear-ObsoleteAdmx -PolicyStore $PolicyStore -Languages $Languages
}

Write-Verbose "`nSaving Admx versions to '$($WorkingDirectory)\AdmxVersions.xml'"
$AdmxVersions | Export-Clixml -Path "$($WorkingDirectory)\AdmxVersions.xml" -Force
#endregion
