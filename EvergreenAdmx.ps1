#Requires -RunAsAdministrator

#region init
<#PSScriptInfo

.VERSION 2411.1

.GUID 999952b7-1337-4018-a1b9-499fad48e734

.AUTHOR Arjan Mensch & Jonathan Pitre

.COMPANYNAME IT-WorXX

.TAGS GroupPolicy GPO Admx Evergreen Automation

.LICENSEURI https://github.com/msfreaks/EvergreenAdmx/blob/main/LICENSE

.PROJECTURI https://github.com/msfreaks/EvergreenAdmx
#>

<#
.SYNOPSIS
    Script to automatically download latest Admx files for several products.

.DESCRIPTION
    Script to automatically download latest Admx files for several products.
    Optionally copies the latest Admx files to a folder of your chosing, for example a Policy Store.

.PARAMETER Windows10Version
    The Windows 10 version to get the Admx files for. This value will be ignored if 'Windows 10' is
    not specified with -Include parameter.
    If the -Include parameter contains 'Windows 10', the latest Windows 10 version will be used.
    Defaults to "Windows11Version" if omitted.

 Note: Windows 11 23H2 policy definitions now supports Windows 10.

.PARAMETER Windows11Version
    The Windows 11 version to get the Admx files for. This value will be ignored if 'Windows 10' is
    not specified with -Include parameter.
    If omitted, defaults to latest version available .

.PARAMETER WorkingDirectory
    Optionally provide a Working Directory for the script.
    The script will store Admx files in a subdirectory called "admx".
    The script will store downloaded files in a subdirectory called "downloads".
    If omitted the script will treat the script's folder as the working directory.

.PARAMETER PolicyStore
    Optionally provide a Policy Store location to copy the Admx files to after processing.

.PARAMETER Languages
    Optionally provide an array of languages to process. Entries must be in 'xy-XY' format.
    If omitted the script will default to 'en-US'.

.PARAMETER UseProductFolders
    When specified the extracted Admx files are copied to their respective product folders in a subfolder of 'Admx' in the WorkingDirectory.

.PARAMETER CustomPolicyStore
    When specified processes a location for custom policy files. Can be UNC format or local folder.
    The script will expect to find .admx files in this location, and at least one language folder holding the .adml file(s).
    Versioning will be done based on the newest file found recursively in this location (any .admx or .adml).
    Note that if any file has changed the script will process all files found in location.

.PARAMETER Include
    Array containing Admx products to include when checking for updates.
    Defaults to "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft Office" if omitted.

.PARAMETER PreferLocalOneDrive
    Microsoft OneDrive Admx files are only available after installing OneDrive.
    If this script is running on a machine that has OneDrive installed, this switch will prevent automatic uninstallation of OneDrive.

.EXAMPLE
    .\EvergreenAdmx.ps1 -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions" -Languages @("en-US", "nl-NL") -UseProductFolders
    Get policy default set of products, storing results in product folders, for both en-us and nl-NL languages, and copies the files to the Policy store.

.LINK
    https://github.com/msfreaks/EvergreenAdmx
    https://msfreaks.wordpress.com

#>

[CmdletBinding(DefaultParameterSetName = 'Windows11Version')]
param(
    [Parameter(Mandatory = $False, ParameterSetName = "Windows10Version", Position = 0)][ValidateSet("1903", "1909", "2004", "20H2", "21H1", "21H2", "22H2")]
    [System.String] $Windows10Version = "22H2",
    [Parameter(Mandatory = $False, ParameterSetName = "Windows11Version", Position = 0)][ValidateSet("21H2", "22H2", "23H2", "24H2")]
    [Alias("WindowsVersion")]
    [System.String] $Windows11Version = "24H2",
    [Parameter(Mandatory = $False)]
    [System.String] $WorkingDirectory = $null,
    [Parameter(Mandatory = $False)]
    [System.String] $PolicyStore = $null,
    [Parameter(Mandatory = $False)]
    [System.String[]] $Languages = @("en-US"),
    [Parameter(Mandatory = $False)]
    [switch] $UseProductFolders,
    [Parameter(Mandatory = $False)]
    [System.String] $CustomPolicyStore = $null,
    [Parameter(Mandatory = $False)][ValidateSet("Custom Policy Store", "Windows 10", "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft Office", "FSLogix", "Adobe Acrobat", "Adobe Reader", "BIS-F", "Citrix Workspace App", "Google Chrome", "Microsoft Desktop Optimization Pack", "Mozilla Firefox", "Zoom Desktop Client", "Azure Virtual Desktop", "Microsoft Winget")]
    [System.String[]] $Include = @("Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft Office"),
    [Parameter(Mandatory = $False)]
    [switch] $PreferLocalOneDrive = $False
)
$ProgressPreference = "SilentlyContinue"
$ErrorActionPreference = "SilentlyContinue"

$admxversions = $null
if (-not $WorkingDirectory) { $WorkingDirectory = $PSScriptRoot }
if (Test-Path -Path "$($WorkingDirectory)\admxversions.xml") { $admxversions = Import-Clixml -Path "$($WorkingDirectory)\admxversions.xml" }
if (-not (Test-Path -Path "$($WorkingDirectory)\admx")) { $null = New-Item -Path "$($WorkingDirectory)\admx" -ItemType Directory -Force }
if (-not (Test-Path -Path "$($WorkingDirectory)\downloads")) { $null = New-Item -Path "$($WorkingDirectory)\downloads" -ItemType Directory -Force }
if ($PolicyStore -and -not $PolicyStore.EndsWith("\")) { $PolicyStore += "\" }
if ($Languages -notmatch "([A-Za-z]{2})-([A-Za-z]{2})$") { Write-Warning "Language not in expected format: $($Languages -notmatch "([A-Za-z]{2})-([A-Za-z]{2})$")" }
if ($CustomPolicyStore -and -not (Test-Path -Path "$($CustomPolicyStore)")) { throw "'$($CustomPolicyStore)' is not a valid path." }
if ($CustomPolicyStore -and -not $CustomPolicyStore.EndsWith("\")) { $CustomPolicyStore += "\" }
if ($CustomPolicyStore -and (Get-ChildItem -Path $CustomPolicyStore -Directory) -notmatch "([A-Za-z]{2})-([A-Za-z]{2})$") { throw "'$($CustomPolicyStore)' does not contain at least one subfolder matching the language format (e.g 'en-US')." }
if ($PreferLocalOneDrive -and $Include -notcontains "Microsoft OneDrive") { Write-Warning "PreferLocalOneDrive is used, but Microsoft OneDrive is not in the list of included products to process." }
$oneDriveADMXFolder = $null
if ((Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\OneDrive").CurrentVersionPath)
{
    $oneDriveADMXFolder = (Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\OneDrive").CurrentVersionPath
}
if ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OneDrive").CurrentVersionPath)
{
    $oneDriveADMXFolder = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OneDrive").CurrentVersionPath
}
if ($PreferLocalOneDrive -and $Include -contains "Microsoft OneDrive" -and $null -eq $oneDriveADMXFolder)
{
    throw "PreferLocalOneDrive will only work if OneDrive is machine installed. User installed OneDrive is not supported.`nLocal machine installed OneDrive not found."
    break
}
Write-Verbose "Windows 10 Version:`t'$($Windows10Version)'"
Write-Verbose "Windows 11 Version:`t'$($Windows11Version)'"
Write-Verbose "WorkingDirectory:`t`t'$($WorkingDirectory)'"
Write-Verbose "PolicyStore:`t`t`t'$($PolicyStore)'"
Write-Verbose "CustomPolicyStore:`t`t'$($CustomPolicyStore)'"
Write-Verbose "Languages:`t`t`t`t'$($Languages)'"
Write-Verbose "Use product folders:`t'$($UseProductFolders)'"
Write-Verbose "Admx path:`t`t`t`t'$($WorkingDirectory)\admx'"
Write-Verbose "Download path:`t`t`t'$($WorkingDirectory)\downloads'"
Write-Verbose "Included:`t`t`t`t'$($Include -join ', ')'"
Write-Verbose "PreferLocalOneDrive:`t'$($PreferLocalOneDrive)'"
#endregion

#region functions
function Get-Link
{
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
        Uri              = $Uri
        Method           = 'GET'
        UseBasicParsing  = $True
        DisableKeepAlive = $True
        ErrorAction      = 'Stop'
    }

    if ($UserAgent)
    {
        $ParamHash.UserAgent = $UserAgent
    }

    if ($Headers)
    {
        $ParamHash.Headers = $Headers
    }

    try
    {
        $Response = Invoke-WebRequest @ParamHash

        foreach ($CurrentPattern in $Pattern)
        {
            $Link = $Response.Links | Where-Object $MatchProperty -Match $CurrentPattern | Select-Object -First 1 -ExpandProperty $ReturnProperty

            if ($PrefixDomain)
            {
                $BaseURL = ($Uri -split '/' | Select-Object -First 3) -join '/'
                $Link = Set-UriPrefix -Uri $Link -Prefix $BaseURL
            }
            elseif ($PrefixParent)
            {
                $BaseURL = ($Uri -split '/' | Select-Object -SkipLast 1) -join '/'
                $Link = Set-UriPrefix -Uri $Link -Prefix $BaseURL
            }

            $Link

        }
    }
    catch
    {
        Write-Error "$($MyInvocation.MyCommand): $($_.Exception.Message)"
    }

}

function Get-Version
{
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

    begin
    {

    }

    process
    {

        if ($PsCmdlet.ParameterSetName -eq 'Uri')
        {

            $ProgressPreference = 'SilentlyContinue'

            try
            {
                $ParamHash = @{
                    Uri              = $Uri
                    Method           = 'GET'
                    UseBasicParsing  = $True
                    DisableKeepAlive = $True
                    ErrorAction      = 'Stop'
                }

                if ($UserAgent)
                {
                    $ParamHash.UserAgent = $UserAgent
                }

                $String = (Invoke-WebRequest @ParamHash).Content
            }
            catch
            {
                Write-Error "Unable to query URL '$Uri': $($_.Exception.Message)"
            }

        }

        foreach ($CurrentString in $String)
        {
            if ($ReplaceWithDot)
            {
                $CurrentString = $CurrentString.Replace('-', '.').Replace('+', '.').Replace('_', '.')
            }
            if ($CurrentString -match $Pattern)
            {
                $matches[1]
            }
            else
            {
                Write-Warning "No version found within $CurrentString using pattern $Pattern"
            }

        }

    }

    end
    {
    }

}

# Replace Get-RedirectedUrl function
function Resolve-Uri
{
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

    begin
    {
        $ProgressPreference = 'SilentlyContinue'
    }

    process
    {

        foreach ($UriToResolve in $Uri)
        {

            try
            {

                $ParamHash = @{
                    Uri              = $UriToResolve
                    Method           = 'Head'
                    UseBasicParsing  = $True
                    DisableKeepAlive = $True
                    ErrorAction      = 'Stop'
                }

                if ($UserAgent)
                {
                    $ParamHash.UserAgent = $UserAgent
                }

                if ($Headers)
                {
                    $ParamHash.Headers = $Headers
                }

                $Response = Invoke-WebRequest @ParamHash

                if ($IsCoreCLR)
                {
                    $ResolvedUri = $Response.BaseResponse.RequestMessage.RequestUri.AbsoluteUri
                }
                else
                {
                    $ResolvedUri = $Response.BaseResponse.ResponseUri.AbsoluteUri
                }

                Write-Verbose "$($MyInvocation.MyCommand): URI resolved to: $ResolvedUri"

                #PowerShell 7 returns each header value as single unit arrays instead of strings which messes with the -match operator coming up, so use Select-Object:
                $ContentDisposition = $Response.Headers.'Content-Disposition' | Select-Object -First 1

                if ($ContentDisposition -match 'filename="?([^\\/:\*\?"<>\|]+)')
                {
                    $FileName = $matches[1]
                    Write-Verbose "$($MyInvocation.MyCommand): Content-Disposition header found: $ContentDisposition"
                    Write-Verbose "$($MyInvocation.MyCommand): File name determined from Content-Disposition header: $FileName"
                }
                else
                {
                    $Slug = [uri]::UnescapeDataString($ResolvedUri.Split('?')[0].Split('/')[-1])
                    if ($Slug -match '^[^\\/:\*\?"<>\|]+\.[^\\/:\*\?"<>\|]+$')
                    {
                        Write-Verbose "$($MyInvocation.MyCommand): URI slug is a valid file name: $FileName"
                        $FileName = $Slug
                    }
                    else
                    {
                        $FileName = $null
                    }
                }

                try
                {
                    $LastModified = [DateTime]($Response.Headers.'Last-Modified' | Select-Object -First 1)
                    Write-Verbose "$($MyInvocation.MyCommand): Last modified date: $LastModified"
                }
                catch
                {
                    Write-Verbose "$($MyInvocation.MyCommand): Unable to parse date from last modified header: $($Response.Headers.'Last-Modified')"
                    $LastModified = $null
                }

            }
            catch
            {
                Throw "$($MyInvocation.MyCommand): Unable to resolve URI: $($_.Exception.Message)"
            }

            if ($ResolvedUri)
            {
                [PSCustomObject]@{
                    Uri          = $ResolvedUri
                    FileName     = $FileName
                    LastModified = $LastModified
                }
            }

        }
    }

    end
    {
    }

}

# Replace Invoke-WebRequest (FASTER DOWNLOADS!)
function Invoke-Download
{
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
        [Alias('URI')]
        [ValidateNotNullOrEmpty()]
        [string]$URL,

        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Destination = $PWD.Path,

        [Parameter(Position = 2)]
        [string]$FileName,

        [string[]]$UserAgent = @('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36', 'Googlebot/2.1 (+http://www.google.com/bot.html)'),

        [string]$TempPath = [System.IO.Path]::GetTempPath(),

        [switch]$IgnoreDate,
        [switch]$BlockFile,
        [switch]$NoClobber,
        [switch]$NoProgress,
        [switch]$PassThru
    )

    begin
    {
        # Required on Windows Powershell only
        if ($PSEdition -eq 'Desktop')
        {
            Add-Type -AssemblyName System.Net.Http
            Add-Type -AssemblyName System.Web
        }

        # Enable TLS 1.2 in addition to whatever is pre-configured
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

        # Create one single client object for the pipeline
        $HttpClient = New-Object System.Net.Http.HttpClient
    }

    process
    {

        Write-Verbose "Requesting headers from URL '$URL'"

        foreach ($UserAgentString in $UserAgent)
        {
            $HttpClient.DefaultRequestHeaders.Remove('User-Agent') | Out-Null
            if ($UserAgentString)
            {
                Write-Verbose "Using UserAgent '$UserAgentString'"
                $HttpClient.DefaultRequestHeaders.Add('User-Agent', $UserAgentString)
            }

            # This sends a GET request but only retrieves the headers
            $ResponseHeader = $HttpClient.GetAsync($URL, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).Result

            # Exit the foreach if success
            if ($ResponseHeader.IsSuccessStatusCode)
            {
                break
            }
        }

        if ($ResponseHeader.IsSuccessStatusCode)
        {
            Write-Verbose 'Successfully retrieved headers'

            if ($ResponseHeader.RequestMessage.RequestUri.AbsoluteUri -ne $URL)
            {
                Write-Verbose "URL '$URL' redirects to '$($ResponseHeader.RequestMessage.RequestUri.AbsoluteUri)'"
            }

            try
            {
                $FileSize = $null
                $FileSize = [int]$ResponseHeader.Content.Headers.GetValues('Content-Length')[0]
                $FileSizeReadable = switch ($FileSize)
                {
                    { $_ -gt 1TB } { '{0:n2} TB' -f ($_ / 1TB); Break }
                    { $_ -gt 1GB } { '{0:n2} GB' -f ($_ / 1GB); Break }
                    { $_ -gt 1MB } { '{0:n2} MB' -f ($_ / 1MB); Break }
                    { $_ -gt 1KB } { '{0:n2} KB' -f ($_ / 1KB); Break }
                    default { '{0} B' -f $_ }
                }
                Write-Verbose "File size: $FileSize bytes ($FileSizeReadable)"
            }
            catch
            {
                Write-Verbose 'Unable to determine file size'
            }

            # Try to get the last modified date from the "Last-Modified" header, use error handling in case string is in invalid format
            try
            {
                $LastModified = $null
                $LastModified = [DateTime]::ParseExact($ResponseHeader.Content.Headers.GetValues('Last-Modified')[0], 'r', [System.Globalization.CultureInfo]::InvariantCulture)
                Write-Verbose "Last modified: $($LastModified.ToString())"
            }
            catch
            {
                Write-Verbose 'Last-Modified header not found'
            }

            if ($FileName)
            {
                $FileName = $FileName.Trim()
                Write-Verbose "Will use supplied filename '$FileName'"
            }
            else
            {
                # Get the file name from the "Content-Disposition" header if available
                try
                {
                    $ContentDispositionHeader = $null
                    $ContentDispositionHeader = $ResponseHeader.Content.Headers.GetValues('Content-Disposition')[0]
                    Write-Verbose "Content-Disposition header found: $ContentDispositionHeader"
                }
                catch
                {
                    Write-Verbose 'Content-Disposition header not found'
                }
                if ($ContentDispositionHeader)
                {
                    $ContentDispositionRegEx = @'
^.*filename\*?\s*=\s*"?(?:UTF-8|iso-8859-1)?(?:'[^']*?')?([^";]+)
'@
                    if ($ContentDispositionHeader -match $ContentDispositionRegEx)
                    {
                        # GetFileName ensures we are not getting a full path with slashes. UrlDecode will convert characters like %20 back to spaces.
                        $FileName = [System.IO.Path]::GetFileName([System.Web.HttpUtility]::UrlDecode($matches[1]))
                        # If any further invalid filename characters are found, convert them to spaces.
                        [IO.Path]::GetinvalidFileNameChars() | ForEach-Object { $FileName = $FileName.Replace($_, ' ') }
                        $FileName = $FileName.Trim()
                        Write-Verbose "Extracted filename '$FileName' from Content-Disposition header"
                    }
                    else
                    {
                        Write-Verbose 'Failed to extract filename from Content-Disposition header'
                    }
                }

                if ([string]::IsNullOrEmpty($FileName))
                {
                    # If failed to parse Content-Disposition header or if it's not available, extract the file name from the absolute URL to capture any redirections.
                    # GetFileName ensures we are not getting a full path with slashes. UrlDecode will convert characters like %20 back to spaces. The URL is split with ? to ensure we can strip off any API parameters.
                    $FileName = [System.IO.Path]::GetFileName([System.Web.HttpUtility]::UrlDecode($ResponseHeader.RequestMessage.RequestUri.AbsoluteUri.Split('?')[0]))
                    [IO.Path]::GetinvalidFileNameChars() | ForEach-Object { $FileName = $FileName.Replace($_, ' ') }
                    $FileName = $FileName.Trim()
                    Write-Verbose "Extracted filename '$FileName' from absolute URL '$($ResponseHeader.RequestMessage.RequestUri.AbsoluteUri)'"
                }
            }

        }
        else
        {
            Write-Verbose 'Failed to retrieve headers'
        }

        if ([string]::IsNullOrEmpty($FileName))
        {
            # If still no filename set, extract the file name from the original URL.
            # GetFileName ensures we are not getting a full path with slashes. UrlDecode will convert characters like %20 back to spaces. The URL is split with ? to ensure we can strip off any API parameters.
            $FileName = [System.IO.Path]::GetFileName([System.Web.HttpUtility]::UrlDecode($URL.Split('?')[0]))
            [System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object { $FileName = $FileName.Replace($_, ' ') }
            $FileName = $FileName.Trim()
            Write-Verbose "Extracted filename '$FileName' from original URL '$URL'"
        }

        $DestinationFilePath = Join-Path $Destination $FileName

        # Exit if -NoClobber specified and file exists.
        if ($NoClobber -and (Test-Path -LiteralPath $DestinationFilePath -PathType Leaf))
        {
            Write-Error 'NoClobber switch specified and file already exists'
            return
        }

        # Open the HTTP stream
        $ResponseStream = $HttpClient.GetStreamAsync($URL).Result

        if ($ResponseStream.CanRead)
        {

            # Check TempPath exists and create it if not
            if (-not (Test-Path -LiteralPath $TempPath -PathType Container))
            {
                Write-Verbose "Temp folder '$TempPath' does not exist"
                try
                {
                    New-Item -Path $Destination -ItemType Directory -Force | Out-Null
                    Write-Verbose "Created temp folder '$TempPath'"
                }
                catch
                {
                    Write-Error "Unable to create temp folder '$TempPath': $_"
                    return
                }
            }

            # Generate temp file name
            $TempFileName = (New-Guid).ToString('N') + ".tmp"
            $TempFilePath = Join-Path $TempPath $TempFileName

            # Check Destiation exists and create it if not
            if (-not (Test-Path -LiteralPath $Destination -PathType Container))
            {
                Write-Verbose "Output folder '$Destination' does not exist"
                try
                {
                    New-Item -Path $Destination -ItemType Directory -Force | Out-Null
                    Write-Verbose "Created output folder '$Destination'"
                }
                catch
                {
                    Write-Error "Unable to create output folder '$Destination': $_"
                    return
                }
            }

            # Open file stream
            try
            {
                $FileStream = [System.IO.File]::Create($TempFilePath)
            }
            catch
            {
                Write-Error "Unable to create file '$TempFilePath': $_"
                return
            }

            if ($FileStream.CanWrite)
            {
                Write-Verbose "Downloading to temp file '$TempFilePath'..."

                $Buffer = New-Object byte[] 64KB
                $BytesDownloaded = 0
                $ProgressIntervalMs = 250
                $ProgressTimer = (Get-Date).AddMilliseconds(-$ProgressIntervalMs)

                while ($true)
                {
                    try
                    {
                        # Read stream into buffer
                        $ReadBytes = $ResponseStream.Read($Buffer, 0, $Buffer.Length)

                        # Track bytes downloaded and display progress bar if enabled and file size is known
                        $BytesDownloaded += $ReadBytes
                        if (!$NoProgress -and (Get-Date) -gt $ProgressTimer.AddMilliseconds($ProgressIntervalMs))
                        {
                            if ($FileSize)
                            {
                                $PercentComplete = [System.Math]::Floor($BytesDownloaded / $FileSize * 100)
                                Write-Progress -Activity "Downloading $FileName" -Status "$BytesDownloaded of $FileSize bytes ($PercentComplete%)" -PercentComplete $PercentComplete
                            }
                            else
                            {
                                Write-Progress -Activity "Downloading $FileName" -Status "$BytesDownloaded of ? bytes" -PercentComplete 0
                            }
                            $ProgressTimer = Get-Date
                        }

                        # If end of stream
                        if ($ReadBytes -eq 0)
                        {
                            Write-Progress -Activity "Downloading $FileName" -Completed
                            $FileStream.Close()
                            $FileStream.Dispose()
                            try
                            {
                                Write-Verbose "Moving temp file to destination '$DestinationFilePath'"
                                $DownloadedFile = Move-Item -LiteralPath $TempFilePath -Destination $DestinationFilePath -Force -PassThru
                            }
                            catch
                            {
                                Write-Error "Error moving file from '$TempFilePath' to '$DestinationFilePath': $_"
                                return
                            }
                            if ($IsWindows)
                            {
                                if ($BlockFile)
                                {
                                    Write-Verbose 'Marking file as downloaded from the internet'
                                    Set-Content -LiteralPath $DownloadedFile -Stream 'Zone.Identifier' -Value "[ZoneTransfer]`nZoneId=3"
                                }
                                else
                                {
                                    Unblock-File -LiteralPath $DownloadedFile
                                }
                            }
                            if ($LastModified -and -not $IgnoreDate)
                            {
                                Write-Verbose 'Setting Last Modified date'
                                $DownloadedFile.LastWriteTime = $LastModified
                            }
                            Write-Verbose 'Download complete!'
                            if ($PassThru)
                            {
                                $DownloadedFile
                            }
                            break
                        }
                        $FileStream.Write($Buffer, 0, $ReadBytes)
                    }
                    catch
                    {
                        Write-Error "Error downloading file: $_"
                        Write-Progress -Activity "Downloading $FileName" -Completed
                        $FileStream.Close()
                        $FileStream.Dispose()
                        break
                    }
                }

            }
        }
        else
        {
            Write-Error 'Failed to start download'
        }

        # Reset this to avoid reusing the same name when fed multiple URLs via the pipeline
        $FileName = $null
    }

    end
    {
        $HttpClient.Dispose()
    }
}

function Copy-Admx
{
    param (
        [string]$SourceFolder,
        [string]$TargetFolder,
        [string]$PolicyStore = $null,
        [string]$ProductName,
        [switch]$Quiet,
        [string[]]$Languages = $null
    )
    if (-not (Test-Path -Path "$($TargetFolder)")) { $null = (New-Item -Path "$($TargetFolder)" -ItemType Directory -Force) }
    if (-not $Languages -or $Languages -eq "") { $Languages = @('en-US') }

    Write-Verbose "Copying Admx files from '$($SourceFolder)' to '$($TargetFolder)'"
    Copy-Item -Path "$($SourceFolder)\*.admx" -Destination "$($TargetFolder)" -Force
    foreach ($language in $Languages)
    {
        if (-not (Test-Path -Path "$($SourceFolder)\$($language)"))
        {
            Write-Verbose "$($language) not found"
            if (-not $Quiet) { Write-Warning "Language '$($language)' not found for '$($ProductName)'. Processing 'en-US' instead." }
            $language = "en-US"
        }
        if (-not (Test-Path -Path "$($TargetFolder)\$($language)"))
        {
            Write-Verbose "'$($TargetFolder)\$($language)' does not exist, creating folder"
            $null = (New-Item -Path "$($TargetFolder)\$($language)" -ItemType Directory -Force)
        }
        Write-Verbose "Copying '$($SourceFolder)\$($language)\*.adml' to '$($TargetFolder)\$($language)'"
        Copy-Item -Path "$($SourceFolder)\$($language)\*.adml" -Destination "$($TargetFolder)\$($language)" -Force
    }
    if ($PolicyStore)
    {
        Write-Verbose "Copying Admx files from '$($SourceFolder)' to '$($PolicyStore)'"
        Copy-Item -Path "$($SourceFolder)\*.admx" -Destination "$($PolicyStore)" -Force
        foreach ($language in $Languages)
        {
            if (-not (Test-Path -Path "$($SourceFolder)\$($language)")) { $language = "en-US" }
            if (-not (Test-Path -Path "$($PolicyStore)$($language)")) { $null = (New-Item -Path "$($PolicyStore)$($language)" -ItemType Directory -Force) }
            Copy-Item -Path "$($SourceFolder)\$($language)\*.adml" -Destination "$($PolicyStore)$($language)" -Force
        }
    }
}

function Get-FSLogixOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for FSLogix
    #>

    try
    {
        # grab URI (redirected url)
        $URL = 'https://aka.ms/fslogix/download'
        $URI = (Resolve-Uri -Uri $URL).URI
        # grab version
        $Version = Get-Version -String $URI -Pattern "(\d+(\.\d+){1,4})"

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch
    {
        Throw $_
    }
}

function Get-MicrosoftOfficeAdmxOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for the Office Admx files (both x64 and x86)
    #>

    $id = "49030"
    $urlVersion = "https://www.microsoft.com/en-us/download/details.aspx?id=$($id)"
    $urlDownload = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=$($id)"

    try
    {

        # load page for version scrape
        $web = (Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $urlVersion -MaximumRedirection 0 -UserAgent 'Googlebot/2.1 (+http://www.google.com/bot.html)').RawContent
        # grab version
        $regEx = '(version\":")((?:\d+\.)+(?:\d+))"'
        $version = ($web | Select-String -Pattern $regEx).Matches.Groups[2].Value

        # load page for uri scrape
        $web = Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $urlDownload -MaximumRedirection 0 -UserAgent 'Googlebot/2.1 (+http://www.google.com/bot.html)'
        # grab x64 version
        $hrefx64 = $web.Links | Where-Object { $_.outerHTML -like "*click here to download manually*" -and $_.href -like "*.exe" -and $_.href -like "*x64*" } | Select-Object -First 1
        # grab x86 version
        $hrefx86 = $web.Links | Where-Object { $_.outerHTML -like "*click here to download manually*" -and $_.href -like "*.exe" -and $_.href -like "*x86*" } | Select-Object -First 1

        # return evergreen object
        return @( @{ Version = $version; URI = $hrefx64.href; Architecture = "x64" }, @{ Version = $version; URI = $hrefx86.href; Architecture = "x86" })
    }
    catch
    {
        Throw $_
    }
}

function Get-WindowsAdmxDownloadId
{
    <#
    .SYNOPSIS
        Returns Windows admx download Id

    .PARAMETER WindowsEdition
        Specifies Windows edition (Example: 11)

    .PARAMETER WindowsVersion
        Specifies Windows version (Example: 23H2)
    #>

    param (
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [ValidateSet("10", "11")]
        [ValidateNotNullOrEmpty()]
        [int]$WindowsEdition = "11",
        [Parameter(Position = 1, ValueFromPipeline = $true)]
        [ValidateSet("1903", "1909", "2004", "20H2", "21H1", "21H2", "22H2", "23H2", "24H2")]
        [ValidateNotNullOrEmpty()]
        [string]$WindowsVersion = "24H2"
    )

    switch ($WindowsEdition)
    {
        10
        {
            return (@( @{ "1903" = "58495" }, @{ "1909" = "100591" }, @{ "2004" = "101445" }, @{ "20H2" = "102157" }, @{ "21H1" = "103124" }, @{ "21H2" = "104042" }, @{ "22H2" = "104677" } ).$WindowsVersion)
            break
        }
        11
        {
            return (@( @{ "21H2" = "103507" }, @{ "22H2" = "104593" }, @{ "23H2" = "105667" }, @{ "24H2" = "106254" } ).$WindowsVersion)
            break
        }
    }
}

function Get-WindowsAdmxOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for the Windows 10 or Windows 11 Admx files

    .PARAMETER DownloadId
        Id returned from Get-WindowsAdmxDownloadId
    #>

    [CmdletBinding()]
    param(
        [int]$DownloadId = (Get-WindowsAdmxDownloadId)
    )

    $urlVersion = "https://www.microsoft.com/en-us/download/details.aspx?id=$($DownloadId)"
    $urlDownload = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=$($DownloadId)"

    try
    {

        # load page for version scrape
        $web = Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $urlVersion -UserAgent 'Googlebot/2.1 (+http://www.google.com/bot.html)'

        # grab version
        $regEx = '(version\":")((?:\d+\.)+(?:\d+))"'
        $version = ('{0}.{1}' -f $DownloadId, ($web | Select-String -Pattern $regEx).Matches.Groups[2].Value)

        # load page for uri scrape
        $web = Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $urlDownload -MaximumRedirection 0 -UserAgent 'Googlebot/2.1 (+http://www.google.com/bot.html)'
        $href = $web.Links | Where-Object { $_.outerHTML -like "*click here to download manually*" -and $_.href -like "*.msi" } | Select-Object -First 1

        # return evergreen object
        return @{ Version = $version; URI = $href.href }
    }
    catch
    {
        Throw $_
    }
}

function Get-OneDriveOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for OneDrive
    #>

    [CmdletBinding()]
    param (
        [bool]$PreferLocalOneDrive
    )

    # detect if OneDrive is installed
    $localOneDrive = (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*OneDrive*" }) | Sort-Object -Property Version -Descending | Select-Object -First 1
    If (-Not $localOneDrive )
    {
        $localOneDrive = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*OneDrive*" }) | Sort-Object -Property Version -Descending | Select-Object -First 1
    }
    If (-Not $localOneDrive )
    {
        $localOneDrive = (Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*OneDrive*" }) | Sort-Object -Property Version -Descending | Select-Object -First 1
    }


    if (($PreferLocalOneDrive) -and [bool]($localOneDrive))
    {
        $URI = "$((Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\OneDrive").CurrentVersionPath)"
        $Version = $localOneDrive.DisplayVersion
        return @{ Version = $Version; URI = $URI }
    }
    elseIf (-Not $PreferLocalOneDrive)
    {
        try
        {
            $url = "https://evergreen-api.stealthpuppy.com/app/MicrosoftOneDrive"
            $architecture = "x64"
            $ring = "Insider"
            $type = "exe"
            $Evergreen = Invoke-RestMethod -Uri $url
            $Evergreen = $Evergreen | Where-Object { $_.Architecture -eq $architecture -and $_.Ring -eq $ring -and $_.Type -eq $type } | `
                    Sort-Object -Property @{ Expression = { [System.Version]$_.Version }; Descending = $true } | Select-Object -First 1

            # grab download uri
            $URI = $Evergreen.URI

            # grab version
            $Version = $Evergreen.Version

            # return evergreen object
            return @{ Version = $Version; URI = $URI }
        }
        catch
        {
            Throw $_
        }
    }
}

function Get-MicrosoftEdgePolicyOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for the Microsoft Edge Admx files
    #>

    try
    {

        $url = "https://edgeupdates.microsoft.com/api/products?view=enterprise"
        # grab json containing product info
        $json = Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing -MaximumRedirection 0 | ConvertFrom-Json
        # filter out the newest release
        $release = ($json | Where-Object { $_.Product -like "Policy" }).Releases | Sort-Object ProductVersion -Descending | Select-Object -First 1
        # grab version
        $Version = $release.ProductVersion
        # grab uri
        $URI = $release.Artifacts[0].Location

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch
    {
        Throw $_
    }

}

function Get-GoogleChromeAdmxOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for the Google Chrome Admx files
    #>

    try
    {

        $URI = "https://dl.google.com/dl/edgedl/chrome/policy/policy_templates.zip"
        # download the file
        Invoke-WebRequest -UseDefaultCredentials -Uri $URI -OutFile "$($env:TEMP)\policy_templates.zip"
        # extract the file
        Expand-Archive -Path "$($env:TEMP)\policy_templates.zip" -DestinationPath "$($env:TEMP)\chromeadmx" -Force

        # open the version file
        $versionfile = (Get-Content -Path "$($env:TEMP)\chromeadmx\VERSION").Split('=')
        $Version = "$($versionfile[1]).$($versionfile[3]).$($versionfile[5]).$($versionfile[7])"

        # cleanup
        Remove-Item -Path "$($env:TEMP)\policy_templates.zip" -Force
        Remove-Item -Path "$($env:TEMP)\chromeadmx" -Recurse -Force

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch
    {
        Throw $_
    }
}

function Get-AdobeAcrobatAdmxOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for the Adobe Acrobat Continuous track Admx files. Use this for Acrobat , 64-bit Reader, and the unified installer.

    .PARAMETER Track
        Specifies Adobe Acrobat track (Example: Continuous)
    #>

    param (
        [Parameter()]
        [ValidateSet("Continuous", "Classic2020", "Classic2017")]
        [ValidateNotNullOrEmpty()]
        [string]$Track = "Continuous"
    )

    switch ($Track)
    {
        Continuous
        {
            $URL = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/misc/AcrobatADMTemplate.zip"
        }
        Classic2020
        {
            $URL = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/Acrobat2020/misc/AcrobatADMTemplate.zip"
        }
        Classic2017
        {
            $URL = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/Acrobat2017/misc/AcrobatADMTemplate.zip"
        }
    }

    try
    {
        # grab uri
        $URI = (Resolve-Uri -Uri $URL).URI

        # grab version
        $LastModifiedDate = (Resolve-Uri -Uri $URL).LastModified
        [version]$Version = $LastModifiedDate.ToString("yyyy.MM.dd")

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch
    {
        Throw $_
    }
}

function Get-AdobeReaderAdmxOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for the Adobe Reader admx files

    .PARAMETER Track
        Specifies Adobe Reader track (Example: Continuous)
    #>

    param (
        [Parameter()]
        [ValidateSet("Continuous", "Classic2020", "Classic2017")]
        [ValidateNotNullOrEmpty()]
        [string]$Track = "Continuous"
    )

    switch ($Track)
    {
        Continuous
        {
            $URL = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/misc/ReaderADMTemplate.zip"
        }
        Classic2020
        {
            $URL = "https://ardownload2.adobe.com/pub/adobe/reader/win/Acrobat2020/misc/ReaderADMTemplate.zip"
        }
        Classic2017
        {
            $URL = "https://ardownload2.adobe.com/pub/adobe/reader/win/Acrobat2017/misc/ReaderADMTemplate.zip"
        }
    }

    try
    {
        # grab uri
        $URI = (Resolve-Uri -Uri $URL).URI

        # grab version
        $LastModifiedDate = (Resolve-Uri -Uri $URL).LastModified
        [version]$Version = $LastModifiedDate.ToString("yyyy.MM.dd")

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch
    {
        Throw $_
    }
}

function Get-CitrixWorkspaceAppAdmxOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Citrix Workspace App ADMX files
    #>

    try
    {

        $url = "https://www.citrix.com/downloads/workspace-app/windows/workspace-app-for-windows-latest.html"
        # grab content
        $web = (Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing -DisableKeepAlive).RawContent
        # find line with ADMX download
        $str = ($web -split "`r`n" | Select-String -Pattern "_ADMX_")[0].ToString().Trim()
        # extract url from ADMX download string
        $URI = "https:$(((Select-String '(\/\/)([^\s,]+)(?=")' -Input $str).Matches.Value))"
        # grab version
        $VersionRegEx = "Version\: ((?:\d+\.)+(?:\d+)) \((.+)\)"
        $Version = ($web | Select-String -Pattern $VersionRegEx).Matches.Groups[1].Value

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch
    {
        Throw $_
    }
}

function Get-MozillaFirefoxAdmxOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Mozilla Firefox ADMX files
    #>

    try
    {

        # define github repo
        $repo = "mozilla/policy-templates"
        # grab latest release properties
        $latest = (Invoke-WebRequest -UseDefaultCredentials -Uri "https://api.github.com/repos/$($repo)/releases" -UseBasicParsing | ConvertFrom-Json)[0]

        # grab version
        $Version = ($latest.tag_name | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        # grab uri
        $URI = $latest.assets.browser_download_url

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch
    {
        Throw $_
    }
}

function Get-BIS-FAdmxOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for BIS-F ADMX files
    #>

    try
    {

        # define github repo
        $repo = "EUCweb/BIS-F"
        # grab latest release properties
        $latest = (Invoke-WebRequest -UseDefaultCredentials -Uri "https://api.github.com/repos/$($repo)/releases" -UseBasicParsing | ConvertFrom-Json)[0]

        # grab version
        $Version = ($latest.tag_name | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        # grab uri
        $URI = $latest.zipball_url

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch
    {
        Throw $_
    }
}

function Get-MDOPAdmxOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for the Desktop Optimization Pack Admx files (both x64 and x86)
    #>

    $id = "55531"
    $urlversion = "https://www.microsoft.com/en-us/download/details.aspx?id=$($id)"
    $urldownload = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=$($id)"
    try
    {
        # load page for version scrape
        $web = Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $urlversion
        # grab version
        $regEx = '(version\":")((?:\d+\.)+(?:\d+))"'
        $version = ($web | Select-String -Pattern $regEx).Matches.Groups[2].Value

        # load page for uri scrape
        $web = Invoke-WebRequest -UseDefaultCredentials -UseBasicParsing -Uri $urldownload -MaximumRedirection 0
        # grab download url
        $href = $web.Links | Where-Object { $_.outerHTML -like "*click here to download manually*" }

        # return evergreen object
        return @{ Version = $Version; URI = $href.href }
    }
    catch
    {
        Throw $_
    }
}

function Get-ZoomDesktopClientAdmxOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Zoom Desktop Client ADMX files
    #>

    try
    {
        $url = "https://support.zoom.com/hc/en/article?id=zm_kb&sysparm_article=KB0065466"

        # grab content
        $web = Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing -UserAgent 'Googlebot/2.1 (+http://www.google.com/bot.html)'
        # find ADMX download
        $URI = (($web.Links | Where-Object { $_.href -like "*msi-templates*.zip" })[-1]).href
        # grab version
        $Version = ($URI.Split("/")[-1] | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch
    {
        Throw $_
    }
}

function Get-CustomPolicyOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Custom Policies

    .PARAMETER CustomPolicyStore
        Folder where Custom Policies can be found
    #>

    param(
        [string] $CustomPolicyStore
    )

    $newestFileDate = Get-Date -Date ((Get-ChildItem -Path $CustomPolicyStore -Include "*.admx", "*.adml" -Recurse | Sort-Object LastWriteTime -Descending) | Select-Object -First 1).LastWriteTime

    $version = Get-Date -Date $newestFileDate -Format "yyMM.dd.HHmmss"

    return @{ Version = $version; URI = $CustomPolicyStore }
}

function Get-AzureVirtualDesktopAdmxOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Azure Virtual Desktop ADMX files
    #>

    try
    {
        $URL = "https://aka.ms/avdgpo"

        # grab uri
        $URI = (Resolve-Uri -Uri $URL).URI

        # grab version
        $LastModifiedDate = (Resolve-Uri -Uri $URL).LastModified
        [version]$Version = $LastModifiedDate.ToString("yyyy.MM.dd")

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch
    {
        Throw $_
    }
}

function Get-WingetAdmxOnline
{
    <#
    .SYNOPSIS
        Returns latest Version and Uri for Winget-cli ADMX files
    #>

    try
    {

        # define github repo
        $repo = "microsoft/winget-cli"
        # grab latest release properties
        $latest = ((Invoke-WebRequest -UseDefaultCredentials -Uri "https://api.github.com/repos/$($repo)/releases" -UseBasicParsing | ConvertFrom-Json) | Where-Object { $_.name -notlike '*-preview' -and $_.draft -eq $false -and $_.assets.browser_download_url -match 'DesktopAppInstallerPolicies.zip' })[0]

        # grab version
        $Version = ($latest.tag_name | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        # grab uri
        $URI = $latest.assets.browser_download_url | Where-Object { $_ -like '*/DesktopAppInstallerPolicies.zip'}

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch
    {
        Throw $_
    }
}

function Get-FSLogixAdmx
{
    <#
    .SYNOPSIS
        Process FSLogix Admx files

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

    $Evergreen = Get-FSLogixOnline
    $ProductName = "FSLogix"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split("/")[-1])"
        try
        {
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
            if ($PolicyStore)
            {
                Write-Verbose "Copying Admx files from '$($SourceAdmx)' to '$($PolicyStore)'"
                Copy-Item -Path "$($SourceAdmx)\*.admx" -Destination "$($PolicyStore)" -Force
                if (-not (Test-Path -Path "$($PolicyStore)en-US")) { $null = (New-Item -Path "$($PolicyStore)en-US" -ItemType Directory -Force) }
                Copy-Item -Path "$($SourceAdmx)\*.adml" -Destination "$($PolicyStore)en-US" -Force
            }

            # cleanup
            Remove-Item -Path "$($env:TEMP)\$($ProductName)" -Recurse -Force

            return $Evergreen
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-MicrosoftOfficeAdmx
{
    <#
    .SYNOPSIS
        Process Office Admx files

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
        [string]$Architecture = "x64",
        [string[]]$Languages = $null
    )

    $Evergreen = Get-MicrosoftOfficeAdmxOnline | Where-Object { $_.Architecture -like $Architecture }
    $ProductName = "Microsoft Office $($Architecture)"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split("/")[-1])"
        try
        {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\office'"
            $null = Start-Process -FilePath $OutFile -ArgumentList "/quiet /norestart /extract:`"$($env:TEMP)\office`"" -PassThru -Wait

            # copy
            $SourceAdmx = "$($env:TEMP)\office\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

            # cleanup
            Remove-Item -Path "$($env:TEMP)\office" -Recurse -Force

            return $Evergreen
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-WindowsAdmx
{
    <#
    .SYNOPSIS
        Process Windows 10 or Windows 11 Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files

    .PARAMETER WindowsVersion
        Official WindowsVersion format

    .PARAMETER WindowsEdition
        Differentiate between Windows 10 and Windows 11

    .PARAMETER Languages
        Languages to check
    #>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string]$WindowsVersion,
        [int]$WindowsEdition,
        [string[]]$Languages = $null
    )

    $id = Get-WindowsAdmxDownloadId -WindowsVersion $WindowsVersion -WindowsEdition $WindowsEdition
    $Evergreen = Get-WindowsAdmxOnline -DownloadId $id
    $ProductName = "Microsoft Windows $($WindowsEdition) $($WindowsVersion)"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }
    $TempFolder = "$($env:TEMP)\$($ProductName)"

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split("/")[-1])"
        try
        {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # install
            Write-Verbose "Installing downloaded Windows $($WindowsEdition) Admx installer"
            $null = Start-Process -FilePath "MsiExec.exe" -WorkingDirectory "$($WorkingDirectory)\downloads" -ArgumentList "/qn /norestart /a `"$($OutFile.split('\')[-1])`" TargetDir=`"$($TempFolder)`"" -PassThru -Wait

            # find installation path
            Write-Verbose "Grabbing installation path for Windows $($WindowsEdition) Admx installer"
            $InstallFolder = Get-ChildItem -Path "$($TempFolder)\Microsoft Group Policy"
            Write-Verbose "Found '$($InstallFolder.Name)'"

            # copy
            $SourceAdmx = "$($TempFolder)\Microsoft Group Policy\$($InstallFolder.Name)\PolicyDefinitions"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

            # cleanup
            Remove-Item -Path $TempFolder -Recurse -Force

            return $Evergreen
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-OneDriveAdmx
{
    <#
    .SYNOPSIS
        Process OneDrive Admx files

    .PARAMETER Version
        Current Version present

    .PARAMETER PolicyStore
        Destination for the Admx files

    .PARAMETER PreferLocalOneDrive
        Check locally only
    #>

    [CmdletBinding()]
    param(
        [string] $Version,
        [string] $PolicyStore = $null,
        [bool] $PreferLocalOneDrive,
        [string[]]$Languages = $null
    )

    if ($PreferLocalOneDrive)
    {
        $Evergreen = Get-OneDriveOnline -PreferLocalOneDrive $PreferLocalOneDrive
    }
    else
    {
        $Evergreen = Get-OneDriveOnline
    }

    $ProductName = "Microsoft OneDrive"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split("/")[-1])"
        try
        {
            if (-not $PreferLocalOneDrive)
            {
                # download
                Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
                Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

                # install
                Write-Verbose "Installing downloaded OneDrive installer"
                $null = Start-Process -FilePath $OutFile -ArgumentList "/allusers /silent" -PassThru
                # wait for setup to complete
                while (Get-Process -Name "OneDriveSetup") { Start-Sleep -Seconds 10 }
                # onedrive starts automatically after setup. kill!
                Stop-Process -Name "OneDrive" -Force

                # find uninstall info
                Write-Verbose "Grabbing uninstallation info from registry for OneDrive installer"
                $uninstall = Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\OneDriveSetup.exe"
                if ($null -eq $uninstall)
                {
                    $uninstall = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\OneDriveSetup.exe"
                }
                if ($null -eq $uninstall)
                {
                    Write-Warning -Message "Unable to find uninstall information for OneDrive."
                }
                else
                {
                    Write-Verbose "Found '$($uninstall.DisplayName)'"

                    # find installation path
                    Write-Verbose "Grabbing installation path for OneDrive installer"
                    $installfolder = $uninstall.DisplayIcon.Substring(0, $uninstall.DisplayIcon.IndexOf("\OneDriveSetup.exe"))
                    Write-Verbose "Found '$($installfolder)'"
                }
            }
            else
            {
                $installfolder = $Evergreen.URI
            }
            # copy
            $SourceAdmx = "$($installfolder)\adm"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            if (-not (Test-Path -Path "$($TargetAdmx)")) { $null = (New-Item -Path "$($TargetAdmx)" -ItemType Directory -Force) }

            Write-Verbose "Copying Admx files from '$($SourceAdmx)' to '$($TargetAdmx)'"
            Copy-Item -Path "$($SourceAdmx)\*.admx" -Destination "$($TargetAdmx)" -Force
            foreach ($language in $Languages)
            {
                if (-not (Test-Path -Path "$($SourceAdmx)\$($language)") -and -not (Test-Path -Path "$($SourceAdmx)\$($language.Substring(0,2))"))
                {
                    if ($language -notlike "en-us") { Write-Warning "Language '$($language)' not found for '$($ProductName)'. Processing 'en-US' instead." }
                    if (-not (Test-Path -Path "$($TargetAdmx)\en-US")) { $null = (New-Item -Path "$($TargetAdmx)\en-US" -ItemType Directory -Force) }
                    Copy-Item -Path "$($SourceAdmx)\*.adml" -Destination "$($TargetAdmx)\en-US" -Force
                }
                else
                {
                    $sourcelanguage = $language; if (-not (Test-Path -Path "$($SourceAdmx)\$($language)")) { $sourcelanguage = $language.Substring(0, 2) }
                    if (-not (Test-Path -Path "$($TargetAdmx)\$($language)")) { $null = (New-Item -Path "$($TargetAdmx)\$($language)" -ItemType Directory -Force) }
                    Copy-Item -Path "$($SourceAdmx)\$($sourcelanguage)\*.adml" -Destination "$($TargetAdmx)\$($language)" -Force
                }
            }

            if ($PolicyStore)
            {
                Write-Verbose "Copying Admx files from '$($SourceAdmx)' to '$($PolicyStore)'"
                Copy-Item -Path "$($SourceAdmx)\*.admx" -Destination "$($PolicyStore)" -Force
                foreach ($language in $Languages)
                {
                    if (-not (Test-Path -Path "$($SourceAdmx)\$($language)") -and -not (Test-Path -Path "$($SourceAdmx)\$($language.Substring(0,2))"))
                    {
                        if (-not (Test-Path -Path "$($PolicyStore)en-US")) { $null = (New-Item -Path "$($PolicyStore)en-US" -ItemType Directory -Force) }
                        Copy-Item -Path "$($SourceAdmx)\*.adml" -Destination "$($PolicyStore)en-US" -Force
                    }
                    else
                    {
                        $sourcelanguage = $language; if (-not (Test-Path -Path "$($SourceAdmx)\$($language)")) { $sourcelanguage = $language.Substring(0, 2) }
                        if (-not (Test-Path -Path "$($PolicyStore)$($language)")) { $null = (New-Item -Path "$($PolicyStore)$($language)" -ItemType Directory -Force) }
                        Copy-Item -Path "$($SourceAdmx)\$($sourcelanguage)\*.adml" -Destination "$($PolicyStore)$($language)" -Force
                    }
                }
            }

            if (-not $PreferLocalOneDrive)
            {
                # uninstall
                Write-Verbose "Uninstalling OneDrive installer"
                $null = Start-Process -FilePath "$($installfolder)\OneDriveSetup.exe" -ArgumentList "/uninstall /allusers" -PassThru -Wait
            }

            return $Evergreen
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-MicrosoftEdgeAdmx
{
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

    $Evergreen = Get-MicrosoftEdgePolicyOnline
    $ProductName = "Microsoft Edge"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($ProductName).cab"
        $ZipFile = "$($WorkingDirectory)\downloads\MicrosoftEdgePolicyTemplates.zip"

        try
        {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\$($ProductName)'"
            $null = (New-Item -Path "$($env:TEMP)\$($ProductName)" -ItemType Directory -Force)
            $null = (expand -F:* "$($OutFile)" "$($env:TEMP)\$($ProductName)" $ZipFile)
            Expand-Archive -Path $ZipFile -DestinationPath "$($env:TEMP)\$($ProductName)" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\$($ProductName)\windows\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

            # cleanup
            Remove-Item -Path $OutFile -Force
            Remove-Item -Path "$env:TEMP\$($ProductName)" -Recurse -Force

            return $Evergreen
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-GoogleChromeAdmx
{
    <#
    .SYNOPSIS
        Process Google Chrome Admx files

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

    $Evergreen = Get-GoogleChromeAdmxOnline
    $ProductName = "Google Chrome"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\googlechromeadmx.zip"
        try
        {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\chromeadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\chromeadmx" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\chromeadmx\windows\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

            # cleanup
            Remove-Item -Path "$($env:TEMP)\chromeadmx" -Recurse -Force

            # chrome update admx is a seperate download
            $url = "https://dl.google.com/dl/update2/enterprise/googleupdateadmx.zip"

            # download
            $OutFile = "$($WorkingDirectory)\downloads\googlechromeupdateadmx.zip"
            Write-Verbose "Downloading '$($url)' to '$($OutFile)'"
            Invoke-WebRequest -UseDefaultCredentials -Uri $url -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\chromeupdateadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\chromeupdateadmx" -Force

            # copy
            $SourceAdmx = "$($env:TEMP)\chromeupdateadmx\GoogleUpdateAdmx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Quiet -Languages $Languages

            # cleanup
            Remove-Item -Path "$($env:TEMP)\chromeupdateadmx" -Recurse -Force

            return $Evergreen
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-AdobeAcrobatAdmx
{
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
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-AdobeAcrobatAdmxOnline
    $ProductName = "Adobe Acrobat"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split("/")[-1])"
        try
        {
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
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-AdobeReaderAdmx
{
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
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-AdobeReaderAdmxOnline
    $ProductName = "Adobe Reader"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split("/")[-1])"
        try
        {
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
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-CitrixWorkspaceAppAdmx
{
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

    $Evergreen = Get-CitrixWorkspaceAppAdmxOnline
    $ProductName = "Citrix Workspace App"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split("?")[0].Split("/")[-1])"
        try
        {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\citrixworkspaceapp'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\citrixworkspaceapp" -Force

            # copy
            # $SourceAdmx = "$($env:TEMP)\citrixworkspaceapp\$($Evergreen.URI.Split("/")[-2].Split("?")[0].SubString(0,$Evergreen.URI.Split("/")[-2].Split("?")[0].IndexOf(".")))"
            $SourceAdmx = (Get-ChildItem -Path "$($env:TEMP)\citrixworkspaceapp\$($Evergreen.URI.Split("/")[-2].Split("?")[0].SubString(0,$Evergreen.URI.Split("/")[-2].Split("?")[0].IndexOf(".")))" -Include "*.admx" -Recurse)[0].DirectoryName
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

            # cleanup
            Remove-Item -Path "$($env:TEMP)\citrixworkspaceapp" -Recurse -Force

            return $Evergreen
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-MozillaFirefoxAdmx
{
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

    $Evergreen = Get-MozillaFirefoxAdmxOnline
    $ProductName = "Mozilla Firefox"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split("/")[-1])"
        try
        {
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
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-ZoomDesktopClientAdmx
{
    <#
    .SYNOPSIS
        Process Zoom Desktop Client Admx files

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

    $Evergreen = Get-ZoomDesktopClientAdmxOnline
    $ProductName = "Zoom Desktop Client"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split("/")[-1])"
        try
        {
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
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-BIS-FAdmx
{
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
        [string]$PolicyStore = $null,
        [string[]]$Languages = $null
    )

    $Evergreen = Get-BIS-FAdmxOnline
    $ProductName = "BIS-F"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\bis-f.$($Evergreen.Version).zip"
        try
        {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\bisfadmx'"
            Expand-Archive -Path $OutFile -DestinationPath "$($env:TEMP)\bisfadmx" -Force

            # find extraction folder
            Write-Verbose "Finding extraction folder"
            $folder = (Get-ChildItem -Path "$($env:TEMP)\bisfadmx" | Sort-Object LastWriteTime -Descending)[0].Name

            # copy
            $SourceAdmx = "$($env:TEMP)\bisfadmx\$($folder)\admx"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName $ProductName -Languages $Languages

            # cleanup
            Remove-Item -Path "$($env:TEMP)\bisfadmx" -Recurse -Force

            return $Evergreen
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-MDOPAdmx
{
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

    $Evergreen = Get-MDOPAdmxOnline
    $ProductName = "Microsoft Desktop Optimization Pack"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split("/")[-1])"
        try
        {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

            # extract
            Write-Verbose "Extracting '$($OutFile)' to '$($env:TEMP)\mdopadmx'"
            $null = (New-Item -Path "$($env:TEMP)\mdopadmx" -ItemType Directory -Force)
            $null = (expand "$($OutFile)" -F:* "$($env:TEMP)\mdopadmx")

            # find app-v folder
            Write-Verbose "Finding App-V folder"
            $appvfolder = (Get-ChildItem -Path "$($env:TEMP)\mdopadmx" -Filter "App-V*" | Sort-Object Name -Descending)[0].Name

            Write-Verbose "Finding MBAM folder"
            $mbamfolder = (Get-ChildItem -Path "$($env:TEMP)\mdopadmx" -Filter "MBAM*" | Sort-Object Name -Descending)[0].Name

            Write-Verbose "Finding UE-V folder"
            $uevfolder = (Get-ChildItem -Path "$($env:TEMP)\mdopadmx" -Filter "UE-V*" | Sort-Object Name -Descending)[0].Name

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
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-CustomPolicyAdmx
{
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
    $ProductName = "Custom Policy Store"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        try
        {
            # copy
            $SourceAdmx = "$($Evergreen.URI)"
            $TargetAdmx = "$($WorkingDirectory)\admx$($ProductFolder)"
            Copy-Admx -SourceFolder $SourceAdmx -TargetFolder $TargetAdmx -PolicyStore $PolicyStore -ProductName "$($ProductName)" -Languages $Languages

            return $Evergreen
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-AzureVirtualDesktopAdmx
{
    <#
    .SYNOPSIS
    Process Azure Virtual Desktop Admx files

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

    $Evergreen = Get-AzureVirtualDesktopAdmxOnline
    $ProductName = "Azure Virtual Desktop"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($ProductName).cab"
        $ZipFile = "$($WorkingDirectory)\downloads\AVDGPTemplate.zip"
        try
        {
            # download
            Write-Verbose "Downloading '$($Evergreen.URI)' to '$($OutFile)'"
            Invoke-Download -URL $Evergreen.URI -Destination "$($WorkingDirectory)\downloads" -FileName "$($ProductName).cab"
            #Invoke-WebRequest -UseDefaultCredentials -Uri $Evergreen.URI -UseBasicParsing -OutFile $OutFile

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
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

function Get-WingetAdmx
{
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

    $Evergreen = Get-WingetAdmxOnline
    $ProductName = "Microsoft winget-cli"
    $ProductFolder = ""; if ($UseProductFolders) { $ProductFolder = "\$($ProductName)" }

    # see if this is a newer version
    if (-not $Version -or [version]$Evergreen.Version -gt [version]$Version)
    {
        Write-Verbose "Found new version $($Evergreen.Version) for '$($ProductName)'"

        # download and process
        $OutFile = "$($WorkingDirectory)\downloads\$($Evergreen.URI.Split("/")[-1])"
        try
        {
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
        }
        catch
        {
            Throw $_
        }
    }
    else
    {
        # version already processed
        return $null
    }
}

#endregion

#region execution
# Custom Policy Store
if ($Include -notcontains 'Custom Policy Store')
{
    Write-Verbose "`nSkipping Custom Policy Store"
}
else
{
    Write-Verbose "`nProcessing Admx files for Custom Policy Store"
    $currentversion = $null
    if ($admxversions.PSObject.properties -match 'CustomPolicyStore') { $currentversion = $admxversions.CustomPolicyStore.Version }
    $admx = Get-CustomPolicyAdmx -Version $currentversion -PolicyStore $PolicyStore -CustomPolicyStore $CustomPolicyStore -Languages $Languages
    if ($admx) { if ($admxversions.CustomPolicyStore) { $admxversions.CustomPolicyStore = $admx } else { $admxversions += @{ CustomPolicyStore = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Windows 10
if ($Include -notcontains 'Windows 10')
{
    Write-Verbose "`nSkipping Windows 10"
}
else
{
    Write-Verbose "`nProcessing Admx files for Windows 10 $($Windows10Version)"
    $admx = Get-WindowsAdmx -Version $admxversions.Windows10.Version -PolicyStore $PolicyStore -WindowsVersion $Windows10Version -WindowsEdition 10 -Languages $Languages
    if ($admx) { if ($admxversions.Windows10) { $admxversions.Windows10 = $admx } else { $admxversions += @{ Windows10 = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Windows 11
if ($Include -notcontains 'Windows 11')
{
    Write-Verbose "`nSkipping Windows 11"
}
else
{
    Write-Verbose "`nProcessing Admx files for Windows 11 $($Windows11Version)"
    $admx = Get-WindowsAdmx -Version $admxversions.Windows11.Version -PolicyStore $PolicyStore -WindowsVersion $Windows11Version -WindowsEdition 11 -Languages $Languages
    if ($admx) { if ($admxversions.Windows11) { $admxversions.Windows11 = $admx } else { $admxversions += @{ Windows11 = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Microsoft Edge
if ($Include -notcontains 'Microsoft Edge')
{
    Write-Verbose "`nSkipping Microsoft Edge"
}
else
{
    Write-Verbose "`nProcessing Admx files for Microsoft Edge"
    $admx = Get-MicrosoftEdgeAdmx -Version $admxversions.Edge.Version -PolicyStore $PolicyStore -Languages $Languages
    if ($admx) { if ($admxversions.Edge) { $admxversions.Edge = $admx } else { $admxversions += @{ Edge = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Microsoft OneDrive
if ($Include -notcontains 'Microsoft OneDrive')
{
    Write-Verbose "`nSkipping Microsoft OneDrive"
}
else
{
    Write-Verbose "`nProcessing Admx files for Microsoft OneDrive"
    $admx = Get-OneDriveAdmx -Version $admxversions.OneDrive.Version -PolicyStore $PolicyStore -PreferLocalOneDrive $PreferLocalOneDrive -Languages $Languages
    if ($admx) { if ($admxversions.OneDrive) { $admxversions.OneDrive = $admx } else { $admxversions += @{ OneDrive = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Microsoft Office
if ($Include -notcontains 'Microsoft Office')
{
    Write-Verbose "`nSkipping Microsoft Office"
}
else
{
    Write-Verbose "`nProcessing Admx files for Microsoft Office"
    $admx = Get-MicrosoftOfficeAdmx -Version $admxversions.Office.Version -PolicyStore $PolicyStore -Architecture "x64" -Languages $Languages
    if ($admx) { if ($admxversions.Office) { $admxversions.Office = $admx } else { $admxversions += @{ Office = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# FSLogix
if ($Include -notcontains 'FSLogix')
{
    Write-Verbose "`nSkipping FSLogix"
}
else
{
    Write-Verbose "`nProcessing Admx files for FSLogix"
    $admx = Get-FSLogixAdmx -Version $admxversions.FSLogix.Version -PolicyStore $PolicyStore -Languages $Languages
    if ($admx) { if ($admxversions.FSLogix) { $admxversions.FSLogix = $admx } else { $admxversions += @{ FSLogix = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Adobe Acrobat
if ($Include -notcontains 'Adobe Acrobat')
{
    Write-Verbose "`nSkipping Adobe Acrobat"
}
else
{
    Write-Verbose "`nProcessing Admx files for Adobe Acrobat"
    $admx = Get-AdobeAcrobatAdmx -Version $admxversions.AdobeAcrobat.Version -PolicyStore $PolicyStore -Languages $Languages
    if ($admx) { if ($admxversions.AdobeAcrobat) { $admxversions.AdobeAcrobat = $admx } else { $admxversions += @{ AdobeAcrobat = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Adobe Reader
if ($Include -notcontains 'Adobe Reader')
{
    Write-Verbose "`nSkipping Adobe Reader"
}
else
{
    Write-Verbose "`nProcessing Admx files for Adobe Reader"
    $admx = Get-AdobeReaderAdmx -Version $admxversions.AdobeReader.Version -PolicyStore $PolicyStore -Languages $Languages
    if ($admx) { if ($admxversions.AdobeReader) { $admxversions.AdobeReader = $admx } else { $admxversions += @{ AdobeReader = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# BIS-F
if ($Include -notcontains 'BIS-F')
{
    Write-Verbose "`nSkipping BIS-F"
}
else
{
    Write-Verbose "`nProcessing Admx files for BIS-F"
    $admx = Get-BIS-FAdmx -Version $admxversions.BISF.Version -PolicyStore $PolicyStore
    if ($admx) { if ($admxversions.BISF) { $admxversions.BISF = $admx } else { $admxversions += @{ BISF = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Citrix Workspace App
if ($Include -notcontains 'Citrix Workspace App')
{
    Write-Verbose "`nSkipping Citrix Workspace App"
}
else
{
    Write-Verbose "`nProcessing Admx files for Citrix Workspace App"
    $admx = Get-CitrixWorkspaceAppAdmx -Version $admxversions.CitrixWorkspaceApp.Version -PolicyStore $PolicyStore -Languages $Languages
    if ($admx) { if ($admxversions.CitrixWorkspaceApp) { $admxversions.CitrixWorkspaceApp = $admx } else { $admxversions += @{ CitrixWorkspaceApp = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Google Chrome
if ($Include -notcontains 'Google Chrome')
{
    Write-Verbose "`nSkipping Google Chrome"
}
else
{
    Write-Verbose "`nProcessing Admx files for Google Chrome"
    $admx = Get-GoogleChromeAdmx -Version $admxversions.GoogleChrome.Version -PolicyStore $PolicyStore -Languages $Languages
    if ($admx) { if ($admxversions.GoogleChrome) { $admxversions.GoogleChrome = $admx } else { $admxversions += @{ GoogleChrome = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Microsoft Desktop Optimization Pack
if ($Include -notcontains 'Microsoft Desktop Optimization Pack')
{
    Write-Verbose "`nSkipping Microsoft Desktop Optimization Pack"
}
else
{
    Write-Verbose "`nProcessing Admx files for Microsoft Desktop Optimization Pack"
    $admx = Get-MDOPAdmx -Version $admxversions.MDOP.Version -PolicyStore $PolicyStore -Languages $Languages
    if ($admx) { if ($admxversions.MDOP) { $admxversions.MDOP = $admx } else { $admxversions += @{ MDOP = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Mozilla Firefox
if ($Include -notcontains 'Mozilla Firefox')
{
    Write-Verbose "`nSkipping Mozilla Firefox"
}
else
{
    Write-Verbose "`nProcessing Admx files for Mozilla Firefox"
    $admx = Get-MozillaFirefoxAdmx -Version $admxversions.MozillaFirefox.Version -PolicyStore $PolicyStore -Languages $Languages
    if ($admx) { if ($admxversions.MozillaFirefox) { $admxversions.MozillaFirefox = $admx } else { $admxversions += @{ MozillaFirefox = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Zoom Desktop Client
if ($Include -notcontains 'Zoom Desktop Client')
{
    Write-Verbose "`nSkipping Zoom Desktop Client"
}
else
{
    Write-Verbose "`nProcessing Admx files for Zoom Desktop Client"
    $admx = Get-ZoomDesktopClientAdmx -Version $admxversions.ZoomDesktopClient.Version -PolicyStore $PolicyStore -Languages $Languages
    if ($admx) { if ($admxversions.ZoomDesktopClient) { $admxversions.ZoomDesktopClient = $admx } else { $admxversions += @{ ZoomDesktopClient = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Azure Virtual Desktop
if ($Include -notcontains 'Azure Virtual Desktop')
{
    Write-Verbose "`nSkipping Azure Virtual Desktop"
}
else
{
    Write-Verbose "`nProcessing Admx files for Azure Virtual Desktop"
    $admx = Get-AzureVirtualDesktopAdmx -Version $admxversions.AzureVirtualDesktop.Version -PolicyStore $PolicyStore -Languages $Languages
    if ($admx) { if ($admxversions.AzureVirtualDesktop) { $admxversions.AzureVirtualDesktop = $admx } else { $admxversions += @{ AzureVirtualDesktop = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

# Microsoft winget-cli
if ($Include -notcontains 'Microsoft Winget')
{
    Write-Verbose "`nSkipping Microsoft Winget"
}
else
{
    Write-Verbose "`nProcessing Admx files for Microsoft Winget"
    $admx = Get-WingetAdmx -Version $admxversions.Winget.Version -PolicyStore $PolicyStore -Languages $Languages
    if ($admx) { if ($admxversions.Winget) { $admxversions.Winget = $admx } else { $admxversions += @{ Winget = @{ Version = $admx.Version; URI = $admx.URI } } } }
}

Write-Verbose "`nSaving Admx versions to '$($WorkingDirectory)\admxversions.xml'"
$admxversions | Export-Clixml -Path "$($WorkingDirectory)\admxversions.xml" -Force
#endregion
