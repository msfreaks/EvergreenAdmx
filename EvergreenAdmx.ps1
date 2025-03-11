#Requires -RunAsAdministrator

#region init
<#PSScriptInfo

.AUTHOR Arjan Mensch & Jonathan Pitre

.TAGS GroupPolicy GPO Admx Evergreen Automation

.LICENSEURI https://github.com/msfreaks/EvergreenAdmx/blob/main/LICENSE

#>

<#
.SYNOPSIS
    Script to automatically download latest Admx files for several products.

.DESCRIPTION
    Script to automatically download latest Admx files for several products.
    Optionally copies the latest Admx files to a folder of your choosing, for example a Policy Store.

.PARAMETER WindowsVersion
    Specifies Windows major version. Supports 10, 11, 2022 or 2025.
    Default is 11.

.PARAMETER WindowsFeatureVersion
    Specifies Windows 10 or 11 feature version to get the Admx files for.
    Valid values are: 1903, 1909, 2004, 20H2, 21H1, 21H2, 22H2 for Windows 10.
    Valid values are: 21H2, 22H2, 23H2, 24H2 for Windows 11.
    Defaults to 24H2.

    Note: Windows 11 23H2 policy definitions now supports Windows 10.

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
    Valid values are: "Windows 10", "Windows 11", "Windows 2022", "Windows 2025", "Microsoft Edge", "Microsoft OneDrive", "Microsoft 365 Apps", "Microsoft FSLogix", "Adobe Acrobat", "Adobe Reader", "BIS-F", "Citrix Workspace App", "Google Chrome", "Microsoft Desktop Optimization Pack", "Mozilla Firefox", "Zoom", "Zoom VDI", "Microsoft AVD", "Microsoft Winget", "Brave Browser".
    Defaults to "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft 365 Apps".

.PARAMETER PreferLocalOneDrive
    Microsoft OneDrive Admx files are only available after installing OneDrive.
    If this script is running on a machine that has OneDrive installed locally, use this switch to prevent automatically uninstalling OneDrive.

.EXAMPLE
    .\EvergreenAdmx.ps1

    Downloads the latest admx files for Windows 11, Microsoft Edge, Microsoft OneDrive, and Microsoft 365 Apps to the current folder.

.EXAMPLE
    .\EvergreenAdmx.ps1 -WindowsVersion 2025

    Downloads the latest admx files for Windows 2025, Microsoft Edge, Microsoft OneDrive, and Microsoft 365 Apps to the current folder.

.EXAMPLE
    .\EvergreenAdmx.ps1 -WorkingDirectory "C:\Temp\EvergreenAdmx" -Include @('Windows 11', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps', 'Microsoft FSLogix')

    Downloads the latest admx files for the specified products to C:\Temp\EvergreenAdmx folder.

.EXAMPLE
    .\EvergreenAdmx.ps1 -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions" -Languages @("en-US", "nl-NL") -UseProductFolders

    Downloads the default set of products policy definitions files, stores them in product folders for both English and Dutch languages, and copies them to the specified Policy store.

.LINK
    https://github.com/msfreaks/EvergreenAdmx

.LINK
    https://msfreaks.wordpress.com

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $False, Position = 0)]
    [ValidateSet('10', '11', '2022', '2025')]
    [System.String] $WindowsVersion = '11',
    [Alias('WindowsFeatureEdition')]
    [ValidateSet('1903', '1909', '2004', '20H2', '21H1', '21H2', '22H2', '23H2', '24H2')]
    [System.String] $WindowsFeatureVersion = $(
        switch ($WindowsVersion) {
            '10' { '22H2' }
            '11' { '24H2' }
            default { '24H2' }
        }
    ),
    [Parameter(Mandatory = $False)]
    [System.String] $WorkingDirectory,
    [Parameter(Mandatory = $False)]
    [System.String] $PolicyStore = $null,
    [Parameter(Mandatory = $False)]
    [System.String[]] $Languages = @('en-US'),
    [Parameter(Mandatory = $False)]
    [switch] $UseProductFolders,
    [Parameter(Mandatory = $False)]
    [System.String] $CustomPolicyStore = $null,
    [Parameter(Mandatory = $False)]
    [switch] $PreferLocalOneDrive,
    [ValidateSet('Custom Policy Store', 'Windows 10', 'Windows 11', 'Windows 2022', 'Windows 2025', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps', 'Microsoft FSLogix', 'Adobe Acrobat', 'Adobe Reader', 'BIS-F', 'Citrix Workspace App', 'Google Chrome', 'Microsoft Desktop Optimization Pack', 'Mozilla Firefox', 'Zoom', 'Zoom VDI', 'Microsoft AVD', 'Microsoft Winget', 'Brave Browser')]
    [System.String[]] $Include = $(
        switch ($WindowsVersion) {
            '10' { @('Windows 10', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps') }
            '11' { @('Windows 11', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps') }
            '2022' { @('Windows 2022', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps') }
            '2025' { @('Windows 2025', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps') }
            default { @('Windows 11', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps') }
        }
    )
)

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
if (-not (Test-Path -Path "$($WorkingDirectory)\admx")) { $null = New-Item -Path "$($WorkingDirectory)\admx" -ItemType Directory -Force }
if (-not (Test-Path -Path "$($WorkingDirectory)\downloads")) { $null = New-Item -Path "$($WorkingDirectory)\downloads" -ItemType Directory -Force }
if ($PolicyStore -and -not $PolicyStore.EndsWith('\')) { $PolicyStore += '\' }
elseif ($null -eq $PolicyStore) { $PolicyStore = $PWD }
if ($Languages -notmatch '([A-Za-z]{2})-([A-Za-z]{2})$') { Write-Warning "Language not in expected format: $($Languages -notmatch '([A-Za-z]{2})-([A-Za-z]{2})$')" }
if ($CustomPolicyStore -and -not (Test-Path -Path "$($CustomPolicyStore)")) { throw "'$($CustomPolicyStore)' is not a valid path." }
if ($CustomPolicyStore -and -not $CustomPolicyStore.EndsWith('\')) { $CustomPolicyStore += '\' }
if ($CustomPolicyStore -and (Get-ChildItem -Path $CustomPolicyStore -Directory) -notmatch '([A-Za-z]{2})-([A-Za-z]{2})$') { throw "'$($CustomPolicyStore)' does not contain at least one subfolder matching the language format (e.g 'en-US')." }
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

function Copy-Admx {
    param (
        [string]$SourceFolder,
        [string]$TargetFolder,
        [string]$PolicyStore = $null,
        [string]$ProductName,
        [switch]$Quiet,
        [string[]]$Languages = $null
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

function Get-EvergreenAdmx365Apps {
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
        Specifies Windows client feature edition. Default is 24H2.

    .EXAMPLE
        Get-WindowsDownloadId -WindowsVersion 11 -WindowsFeatureVersion 24H2
    #>

    param (
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [ValidateSet('10', '11', '2022', '2025')]
        [ValidateNotNullOrEmpty()]
        [int]$WindowsVersion = '11',
        [Parameter(Position = 1, ValueFromPipeline = $true)]
        [ValidateScript({
                if ($WindowsVersion -eq '10' -and $_ -in @('1903', '1909', '2004', '20H2', '21H1', '21H2', '22H2')) {
                    return $true
                } elseif ($WindowsVersion -eq '11' -and $_ -in @('21H2', '22H2', '23H2', '24H2')) {
                    return $true
                } elseif ($WindowsVersion -eq '2022' -or $WindowsVersion -eq '2025') {
                    return $true
                } else {
                    throw "Invalid Windows Feature Version '$_' for Windows $WindowsVersion. Windows 10 supports: 1903, 1909, 2004, 20H2, 21H1, 21H2, 22H2. Windows 11 supports: 21H2, 22H2, 23H2, 24H2. Windows 2022 and 2025 has no Windows Feature Versions."
                }
            })]
        [ValidateNotNullOrEmpty()]
        [string]$WindowsFeatureVersion = '24H2'
    )

    switch ($WindowsVersion) {
        10 {
            return (@( @{ '1903' = '58495' }, @{ '1909' = '100591' }, @{ '2004' = '101445' }, @{ '20H2' = '102157' }, @{ '21H1' = '103124' }, @{ '21H2' = '104042' }, @{ '22H2' = '104677' } ).$WindowsFeatureVersion)
            break
        }
        11 {
            return (@( @{ '21H2' = '103507' }, @{ '22H2' = '104593' }, @{ '23H2' = '105667' }, @{ '24H2' = '106254' } ).$WindowsFeatureVersion)
            break
        }
        2022 {
            return @( '104003' )
            break
        }
        2025 {
            return @( '106295' )
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

        If (-not [string]::IsNullOrWhiteSpace($UserInstall)) {
            Write-Verbose "User OneDrive install found: $($UserInstall.DisplayVersion)"
            $isOneDriveInstalled = $true
            $OneDriveInstalledVersion = $UserInstall.DisplayVersion
            $global:oneDriveADMXFolder = (Get-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\OneDrive').CurrentVersionPath
        }
        If (-not [string]::IsNullOrWhiteSpace($Systemx64Install)) {
            Write-Verbose "System x64 OneDrive install found: $($Systemx64Install.DisplayVersion)"
            $isOneDriveInstalled = $true
            $OneDriveInstalledVersion = $Systemx64Install.DisplayVersion
            $global:oneDriveADMXFolder = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\OneDrive').CurrentVersionPath
        }
        If (-not [string]::IsNullOrWhiteSpace($Systemx86Install)) {
            Write-Verbose "System x86 OneDrive install found: $($Systemx86Install.DisplayVersion )"
            $isOneDriveInstalled = $true
            $OneDriveInstalledVersion = $Systemx86Install.DisplayVersion
            $global:oneDriveADMXFolder = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\OneDrive').CurrentVersionPath
        } else {
            $isOneDriveInstalled = $false
        }

        if ($PreferLocalOneDrive) {
            If ($isOneDriveInstalled) {
                return @{ Version = $OneDriveInstalledVersion }
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
        Returns latest Version and Uri for the Adobe Acrobat Continuous track Admx files. Use this for Acrobat , 64-bit Reader, and the unified installer.

    .PARAMETER Track
        Specifies Adobe Acrobat track (Example: Continuous)
    #>

    param (
        [Parameter()]
        [ValidateSet('Continuous', 'Classic2020', 'Classic2017')]
        [ValidateNotNullOrEmpty()]
        [string]$Track = 'Continuous'
    )

    switch ($Track) {
        Continuous {
            $URL = 'https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/misc/AcrobatADMTemplate.zip'
        }
        Classic2020 {
            $URL = 'https://ardownload2.adobe.com/pub/adobe/acrobat/win/Acrobat2020/misc/AcrobatADMTemplate.zip'
        }
        Classic2017 {
            $URL = 'https://ardownload2.adobe.com/pub/adobe/acrobat/win/Acrobat2017/misc/AcrobatADMTemplate.zip'
        }
    }

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
        Returns latest Version and Uri for the Adobe Reader admx files

    .PARAMETER Track
        Specifies Adobe Reader track (Example: Continuous)
    #>

    param (
        [Parameter()]
        [ValidateSet('Continuous', 'Classic2020', 'Classic2017')]
        [ValidateNotNullOrEmpty()]
        [string]$Track = 'Continuous'
    )

    switch ($Track) {
        Continuous {
            $URL = 'https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/misc/ReaderADMTemplate.zip'
        }
        Classic2020 {
            $URL = 'https://ardownload2.adobe.com/pub/adobe/reader/win/Acrobat2020/misc/ReaderADMTemplate.zip'
        }
        Classic2017 {
            $URL = 'https://ardownload2.adobe.com/pub/adobe/reader/win/Acrobat2017/misc/ReaderADMTemplate.zip'
        }
    }

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

function Get-EvergreenAdmxMDOP {
    <#
    .SYNOPSIS
        Returns latest Version and Uri for the Desktop Optimization Pack Admx files (both x64 and x86)
    #>

    $id = '55531'
    $urlversion = "https://www.microsoft.com/en-us/download/details.aspx?id=$($id)"
    $JSONBlobPattern = '(?<scriptStart><script>[\w.]+__DLCDetails__=).*?(?<JSObject-scriptStart></script>)'
    try {
        # Load web page for scrapping admx version
        $web = Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $urlversion
        # grab version
        $regEx = '(version\":")((?:\d+\.)+(?:\d+))"'
        $version = ($web | Select-String -Pattern $regEx).Matches.Groups[2].Value

        # carve JSON from script tag
        $web = $web.Content | Select-String -Pattern $JSONBlobPattern | Select-Object -ExpandProperty Matches | ForEach-Object { $_.Groups['JSObject'].Value } | Select-Object -First 1 | ConvertFrom-Json
        # grab download url
        $href = $web.dlcDetailsView.downloadFile

        # return evergreen object
        return @{ Version = $Version; URI = $href.url }
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

function Get-EvergreenAdmxSlack {
    <#
    .SYNOPSIS
        Returns url and latest Slack policy definition files.
    #>

    try {
        $url = 'https://slack.com/help/articles/11906214948755-Manage-desktop-app-configurations'

        # grab content
        $web = (Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing -DisableKeepAlive).RawContent
        # find line with ADMX download
        $str = ($web -split "`r`n")
        # extract url from ADMX download string
        $regEx = '(https\:\/\/[^\s,]+(?=))(\"\>Group Policy Object template)'
        $URI = (Select-String -Pattern $regEx -Input $str).Matches.Groups[1].Value

        # grab version

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    } catch {
        Throw $_
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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

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
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # Extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $null = (New-Item -Path "$($env:TEMP)\$($ProductName)" -ItemType Directory -Force)
            $null = (expand -F:* "$($OutFile)" "$($env:TEMP)\$($ProductName)" $ZipFile)
            Expand-Archive -Path $ZipFile -DestinationPath "$($env:TEMP)\$($ProductName)" -Force

            # Copy
            $SourceAdmx = "$($env:TEMP)\$($ProductName)\windows\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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
                Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

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
                $installfolder = $oneDriveADMXFolder
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

function Invoke-EvergreenAdmx365Apps {
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

    $Evergreen = Get-EvergreenAdmx365Apps | Where-Object { $_.Architecture -like $Architecture }
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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # Extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\office'"
            $null = Start-Process -FilePath $OutFile -ArgumentList "/quiet /norestart /extract:`"$($env:TEMP)\office`"" -PassThru -Wait

            # Copy
            $SourceAdmx = "$($env:TEMP)\office\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # Extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\chromeadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\chromeadmx" -Force

            # Copy
            $SourceAdmx = "$($env:TEMP)\chromeadmx\windows\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

            # Cleanup
            Remove-Item -Path "$($env:TEMP)\chromeadmx" -Recurse -Force

            # Chrome update admx is a separate download
            $url = 'https://dl.google.com/dl/update2/enterprise/googleupdateadmx.zip'

            # Download
            $OutFile = "$($WorkingDirectory)\downloads\googlechromeupdateadmx.zip"
            Write-Verbose "Downloading '$($url)' to '$($OutFile)'"
            Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing -OutFile $OutFile

            # Extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\chromeupdateadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\chromeupdateadmx" -Force

            # Copy
            $SourceAdmx = "$($env:TEMP)\chromeupdateadmx\GoogleUpdateAdmx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Quiet -Languages $Languages

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
            Invoke-WebRequest -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\$($ProductName)" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\$($ProductName)"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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
            Invoke-WebRequest -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\AdobeReader'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\AdobeReader" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\AdobeReader"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\citrixworkspaceapp'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\citrixworkspaceapp" -Force

            # copy
            # $SourceAdmx = "$($env:TEMP)\citrixworkspaceapp\$($Evergreen.URI.Split("/")[-2].Split("?")[0].SubString(0,$Evergreen.URI.Split("/")[-2].Split("?")[0].IndexOf(".")))"
            $SourceAdmx = (Get-ChildItem -Path "$($env:TEMP)\citrixworkspaceapp\$($Evergreen.URI.Split('/')[-2].Split('?')[0].SubString(0,$Evergreen.URI.Split('/')[-2].Split('?')[0].IndexOf('.')))" -Include '*.admx' -Recurse)[0].DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\firefoxadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\firefoxadmx" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\firefoxadmx\windows"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

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
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

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
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\bisfadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\bisfadmx" -Force

            # find extraction folder
            Write-Verbose 'Finding extraction folder'
            $folder = (Get-ChildItem -Path "$($env:TEMP)\bisfadmx" | Sort-Object LastWriteTime -Descending)[0].Name

            # copy
            $SourceAdmx = "$($env:TEMP)\bisfadmx\$($folder)\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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

function Invoke-EvergreenAdmxMDOP {
    <#
    .SYNOPSIS
        Process MDOP Admx files

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

    $Evergreen = Get-EvergreenAdmxMDOP
    $ProductName = 'Microsoft Desktop Optimization Pack'
    $ProductFolder = ''; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split('/')[-1])"
        try {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\mdopadmx'"
            $null = (New-Item -Path "$($env:TEMP)\mdopadmx" -ItemType Directory -Force)
            $null = (expand "$($OutFile)" -F:* "$($env:TEMP)\mdopadmx")

            # find app-v folder
            Write-Verbose 'Finding App-V folder'
            $appvfolder = (Get-ChildItem -Path "$($env:TEMP)\mdopadmx" -Filter 'App-V*' | Sort-Object Name -Descending)[0].Name

            Write-Verbose 'Finding MBAM folder'
            $mbamfolder = (Get-ChildItem -Path "$($env:TEMP)\mdopadmx" -Filter 'MBAM*' | Sort-Object Name -Descending)[0].Name

            Write-Verbose 'Finding UE-V folder'
            $uevfolder = (Get-ChildItem -Path "$($env:TEMP)\mdopadmx" -Filter 'UE-V*' | Sort-Object Name -Descending)[0].Name

            # copy
            $SourceAdmx = "$($env:TEMP)\mdopadmx\$($appvfolder)"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName "$($ProductName) - App-V" -Languages $Languages
            $SourceAdmx = "$($env:TEMP)\mdopadmx\$($mbamfolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName "$($ProductName) - MBAM" -Languages $Languages
            $SourceAdmx = "$($env:TEMP)\mdopadmx\$($uevfolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName "$($ProductName) - UE-V" -Languages $Languages

            # cleanup
            Remove-Item -Path "$($env:TEMP)\mdopadmx" -Recurse -Force

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
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName "$($ProductName)" -Languages $Languages

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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $null = (New-Item -Path "$($env:TEMP)\$($ProductName)" -ItemType Directory -Force)
            $null = (expand -F:* "$($OutFile)" "$($env:TEMP)\$($ProductName)" $ZipFile)
            Expand-Archive -Path $ZipFile -DestinationPath "$($env:TEMP)\$($ProductName)" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\$($ProductName)"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\wingetadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\wingetadmx" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\wingetadmx\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

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
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

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
    $admx = Invoke-EvergreenAdmxWindows -Version $AdmxVersions.Windows.Version -PolicyStore $PolicyStore -WindowsFeatureVersion $WindowsFeatureVersion -WindowsVersion 10 -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Windows10' -AdmxData $admx
}

# Windows 11
if ($Include -notcontains 'Windows 11') {
    Write-Verbose "`nSkipping Windows 11"
} else {
    Write-Verbose "`nProcessing Admx files for Windows 11 $($WindowsFeatureVersion)"
    $admx = Invoke-EvergreenAdmxWindows -Version $AdmxVersions.Windows.Version -PolicyStore $PolicyStore -WindowsFeatureVersion $WindowsFeatureVersion -WindowsVersion 11 -Languages $Languages
    if ($null -ne $admx) {
        Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Windows11' -AdmxData $admx
    } else {
        Write-Warning 'Failed to retrieve Windows 11 ADMX files. Skipping update.'
    }
}

# Windows 2022
if ($Include -notcontains 'Windows 2022') {
    Write-Verbose "`nSkipping Windows Server 2022"
} else {
    Write-Verbose "`nProcessing Admx files for Windows Server 2022"
    $admx = Invoke-EvergreenAdmxWindows -Version $AdmxVersions.Windows.Version -PolicyStore $PolicyStore -WindowsVersion 2022 -Languages $Languages
    if ($null -ne $admx) {
        Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Windows2022' -AdmxData $admx
    } else {
        Write-Warning 'Failed to retrieve Windows Server 2022 ADMX files. Skipping update.'
    }
}

# Windows 2025
if ($Include -notcontains 'Windows 2025') {
    Write-Verbose "`nSkipping Windows Server 2025"
} else {
    Write-Verbose "`nProcessing Admx files for Windows Server 2025"
    $admx = Invoke-EvergreenAdmxWindows -Version $AdmxVersions.Windows.Version -PolicyStore $PolicyStore -WindowsVersion 2025 -Languages $Languages
    if ($null -ne $admx) {
        Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Windows2025' -AdmxData $admx
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
    $admx = Invoke-EvergreenAdmx365Apps -Version $AdmxVersions['365Apps'].Version -PolicyStore $PolicyStore -Architecture 'x64' -Languages $Languages
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

# Microsoft Desktop Optimization Pack
if ($Include -notcontains 'Microsoft Desktop Optimization Pack') {
    Write-Verbose "`nSkipping Microsoft Desktop Optimization Pack"
} else {
    Write-Verbose "`nProcessing Admx files for Microsoft Desktop Optimization Pack"
    $admx = Invoke-EvergreenAdmxMDOP -Version $AdmxVersions.MDOP.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'MDOP' -AdmxData $admx
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
    $admx = Invoke-EvergreenAdmxZoomVDI -Version $AdmxVersions.ZoomVDI.Version -PolicyStore $PolicyStore -Languages $Languages
    Update-AdmxVersion -AdmxVersions ([ref]$AdmxVersions) -ProductKey 'Zoom VDI' -AdmxData $admx
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

Write-Verbose "`nSaving Admx versions to '$($WorkingDirectory)\AdmxVersions.xml'"
$AdmxVersions | Export-Clixml -Path "$($WorkingDirectory)\AdmxVersions.xml" -Force
#endregion
