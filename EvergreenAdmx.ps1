<#PSScriptInfo

.VERSION 2012.6

.GUID 999952b7-1337-4018-a1b9-499fad48e734

.AUTHOR Arjan Mensch

.COMPANYNAME IT-WorXX

.TAGS GroupPolicy GPO Admx Evergreen Automation

.LICENSEURI https://github.com/msfreaks/EvergreenAdmx/blob/main/LICENSE

.PROJECTURI https://github.com/msfreaks/EvergreenAdmx

.DESCRIPTION
 Script to automatically download latest Admx files for several products.
 Optionally copies the latest Admx files to a folder of your chosing, for example a Policy Store.

#> 
<#
.SYNOPSIS
 Script to automatically download latest Admx files for several products.

.DESCRIPTION
 Script to automatically download latest Admx files for several products.
 Optionally copies the latest Admx files to a folder of your chosing, for example a Policy Store.

.PARAMETER WindowsVersion
 The Windows 10 version to get the Admx files for.
 If omitted the newest version supported by this script will be used.

.PARAMETER WorkingDirectory
 Optionally provide a Working Directory for the script.
 The script will store Admx files in a subdirectory called "admx".
 The script will store downloaded files in a subdirectory called "downloads".
 If omitted the script will treat the script's folder as the working directory.

.PARAMETER PolicyStore
 Optionally provide a Policy Store location to copy the Admx files to after processing.

.PARAMETER Languages
 Optionally provide an array of languages to process. Entries must be in 'xy-XY' format.
 If omitted the script will process 'en-US'.

.PARAMETER UseProductFolders
 When specified the extracted Admx files are copied to their respective product folders in a subfolder of 'Admx' in the WorkingDirectory.

.EXAMPLE
 .\EvergreenAdmx.ps1 -WindowsVersion "20H2" -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions" -Languages @("en-US", "nl-NL") -UseProductFolders

.LINK
 https://github.com/msfreaks/EvergreenAdmx
 https://msfreaks.wordpress.com

#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)][ValidateSet("1903", "1909", "2004", "20H2")]
    [string]$WindowsVersion = "20H2",
    [Parameter(Mandatory = $false)]
    [string]$WorkingDirectory = $null,
    [Parameter(Mandatory = $false)]
    [string]$PolicyStore = $null,
    [Parameter(Mandatory = $false)]
    [string[]]$Languages = @("en-US"),
    [Parameter(Mandatory = $false)]
    [switch]$UseProductFolders
)

#region init
$admxversions = $null
if (-not $WorkingDirectory) { $WorkingDirectory = $PSScriptRoot }
if (Test-Path -Path "$($WorkingDirectory)\admxversions.xml") { $admxversions = Import-Clixml -Path "$($WorkingDirectory)\admxversions.xml" }
if (-not (Test-Path -Path "$($WorkingDirectory)\admx")) { $null = mkdir -Path "$($WorkingDirectory)\admx" -Force }
if (-not (Test-Path -Path "$($WorkingDirectory)\downloads")) { $null = mkdir -Path "$($WorkingDirectory)\downloads" -Force }
if ($PolicyStore -and -not $PolicyStore.EndsWith("\")) { $policypath += "\" }
if ($Languages -notmatch "([A-Za-z]{2})-([A-Za-z]{2})$") { Write-Warning "Language not in expected format: $($Languages -notmatch "([A-Za-z]{2})-([A-Za-z]{2})$")" }

Write-Verbose "Windows Version:`t`t$($WindowsVersion)"
Write-Verbose "WorkingDirectory:`t`t$($WorkingDirectory)"
Write-Verbose "PolicyStore:`t`t`t$($PolicyStore)"
Write-Verbose "Languages:`t`t`t`t$($Languages)"
Write-Verbose "Use product folders:`t$($UseProductFolders)"
Write-Verbose "Admx path:`t`t`t`t$($WorkingDirectory)\admx"
Write-Verbose "Download path:`t`t`t$($WorkingDirectory)\downloads"
#endregion

#region functions
function Get-Windows10AdmxDownloadId {
<#
    .SYNOPSIS
    Returns download Id for Admx file based on Windows 10 version

    .PARAMETER WindowsVersion
    Official WindowsVersion format
#>

    param (
        [string]$WindowsVersion
    )

    return (@( @{ "1903" = "58495" }, @{ "1909" = "100591" }, @{ "2004" = "101445" }, @{ "20H2" = "102157" } ).$WindowsVersion)
}

function Copy-Admx {
    param (
        [string]$SourceFolder,
        [string]$TargetFolder,
        [string]$PolicyStore = $null,
        [string]$ProductName,
        [switch]$Quiet
    )
    if (-not (Test-Path -Path "$($TargetFolder)")) { $null = (mkdir -Path "$($TargetFolder)" -Force) }

    Write-Verbose "Copying Admx files from '$($SourceFolder)' to '$($TargetFolder)'"
    Copy-Item -Path "$($SourceFolder)\*.admx" -Destination "$($TargetFolder)" -Force
    foreach ($language in $Languages) {
        if (-not (Test-Path -Path "$($SourceFolder)\$($language)")) {
            if (-not $Quiet) { Write-Warning "Language '$($language)' not found for '$($ProductName)'. Processing 'en-US' instead." }
            $language = "en-US"
        }
        if (-not (Test-Path -Path "$($TargetFolder)\$($language)")) { $null = (mkdir -Path "$($TargetFolder)\$($language)" -Force) }
        Copy-Item -Path "$($SourceFolder)\$($language)\*.adml" -Destination "$($TargetFolder)\$($language)" -Force
    }
    if ($PolicyStore) {
        Write-Verbose "Copying Admx files from '$($SourceFolder)' to '$($PolicyStore)'"
        Copy-Item -Path "$($SourceFolder)\*.admx" -Destination "$($PolicyStore)" -Force
        foreach ($language in $Languages) {
            if (-not (Test-Path -Path "$($SourceFolder)\$($language)")) { $language = "en-US" }
            if (-not (Test-Path -Path "$($PolicyStore)\$($language)")) { $null = (mkdir -Path "$($PolicyStore)\$($language)" -Force) }
            Copy-Item -Path "$($SourceFolder)\$($language)\*.adml" -Destination "$($PolicyStore)\$($language)" -Force
        }
    }
}

function Get-FSLogixOnline {
<#
    .SYNOPSIS
    Returns latest Version and Uri for FSLogix
#>

    try {
        $ProgressPreference = 'SilentlyContinue'
        $url = "https://aka.ms/fslogix_download"
        # grab content without redirecting to the download
        $web = Invoke-WebRequest -Uri $url -UseBasicParsing -MaximumRedirection 0 -ErrorAction Ignore
        $url = $web.Headers.Location
        # grab content without redirecting to the download
        $web = Invoke-WebRequest -Uri $url -UseBasicParsing -MaximumRedirection 0 -ErrorAction Ignore
        # grab uri
        $URI = $web.Headers.Location
        # grab version
        $Version = ($URI.Split("/")[-1] | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch {
        Throw $_
    }
}

function Get-MicrosoftOfficeAdmxOnline {
<#
    .SYNOPSIS
    Returns latest Version and Uri for the Office Admx files (both x64 and x86)
#>

    $id = "49030"
    $urlversion = "https://www.microsoft.com/en-us/download/details.aspx?id=$($id)"
    $urldownload = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=$($id)"
    try {
        $ProgressPreference = 'SilentlyContinue'
        # load page for version scrape
        $web = Invoke-WebRequest -UseBasicParsing -Uri $urlversion -ErrorAction SilentlyContinue
        $str = ($web.ToString() -split "[`r`n]" | Select-String "Version:").ToString()
        # grab version
        $Version = ($str | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        # load page for uri scrape
        $web = Invoke-WebRequest -UseBasicParsing -Uri $urldownload -ErrorAction SilentlyContinue -MaximumRedirection 0
        # grab x64 version
        $hrefx64 = $web.Links | Where-Object { $_.outerHTML -like "*click here to download manually*" -and $_.href -like "*.exe" -and $_.href -like "*x64*" } | Select-Object -First 1
        # grab x86 version
        $hrefx86 = $web.Links | Where-Object { $_.outerHTML -like "*click here to download manually*" -and $_.href -like "*.exe" -and $_.href -like "*x86*" } | Select-Object -First 1

        # return evergreen object
        return @( @{ Version = $Version; URI = $hrefx64.href; Architecture = "x64" }, @{ Version = $Version; URI = $hrefx86.href; Architecture = "x86" })
    }
    catch {
        Throw $_
    }
}

function Get-Windows10AdmxOnline {
<#
    .SYNOPSIS
    Returns latest Version and Uri for the Windows 10 Admx files

    .PARAMETER DownloadId
    Id returned from Get-WindowsAdmxDownloadId
    #>
    <#
    .SYNOPSIS
    Returns latest Version and Uri for the Windows 10 Admx files

    .PARAMETER DownloadId
    Id returned from Get-WindowsAdmxDownloadId
#>

    param(
        [string]$DownloadId
    )

    $urlversion = "https://www.microsoft.com/en-us/download/details.aspx?id=$($DownloadId)"
    $urldownload = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=$($DownloadId)"
    try {
        $ProgressPreference = 'SilentlyContinue'
        # load page for version scrape
        $web = Invoke-WebRequest -UseBasicParsing -Uri $urlversion -ErrorAction SilentlyContinue
        $str = ($web.ToString() -split "[`r`n]" | Select-String "Version:").ToString()
        # grab version
        $Version = "$($DownloadId).$(($str | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString())"
        # load page for uri scrape
        $web = Invoke-WebRequest -UseBasicParsing -Uri $urldownload -ErrorAction SilentlyContinue -MaximumRedirection 0
        $href = $web.Links | Where-Object { $_.outerHTML -like "*click here to download manually*" -and $_.href -like "*.msi" } | Select-Object -First 1

        # return evergreen object
        return @{ Version = $Version; URI = $href.href }
    }
    catch {
        Throw $_
    }
}

function Get-OneDriveOnline {
<#
    .SYNOPSIS
    Returns latest Version and Uri for OneDrive
#>

    try {
        $ProgressPreference = 'SilentlyContinue'
        $url = "https://go.microsoft.com/fwlink/p/?LinkID=844652"
        # grab content without redirecting to the download
        $web = Invoke-WebRequest -Uri $url -UseBasicParsing -MaximumRedirection 0 -ErrorAction Ignore
        # grab uri
        $URI = $web.Headers.Location
        # grab version
        $Version = ($URI.Split("/")[-2] | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch {
        Throw $_
    }
}

function Get-MicrosoftEdgePolicyOnline {
<#
    .SYNOPSIS
    Returns latest Version and Uri for the Microsoft Edge Admx files
#>

    try {
        $ProgressPreference = 'SilentlyContinue'
        $url = "https://edgeupdates.microsoft.com/api/products?view=enterprise"
        # grab json containing product info
        $json = Invoke-WebRequest -Uri $url -UseBasicParsing -MaximumRedirection 0 -ErrorAction Ignore | ConvertFrom-Json
        # filter out the newest release
        $release = ($json | Where-Object { $_.Product -like "Policy" }).Releases | Sort-Object ProductVersion -Descending | Select-Object -First 1
        # grab version
        $Version = $release.ProductVersion
        # grab uri
        $URI = ($release.Artifacts | Where-Object { $_.ArtifactName -like "zip" }).Location

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch {
        Throw $_
    }

}

function Get-GoogleChromeAdmxOnline {
<#
    .SYNOPSIS
    Returns latest Version and Uri for the Google Chrome Admx files
#>

    try {
        $ProgressPreference = 'SilentlyContinue'
        $URI = "https://dl.google.com/dl/edgedl/chrome/policy/policy_templates.zip"
        # download the file
        Invoke-WebRequest -Uri $URI -OutFile "$($env:TEMP)\policy_templates.zip"
        # extract the file
        Expand-Archive -Path "$($env:TEMP)\policy_templates.zip" -DestinationPath "$($env:TEMP)\chromeadmx" -Force
        # open the version file
        $versionfile = (Get-Content -Path "$($env:TEMP)\chromeadmx\VERSION").Split('=')
        $Version = "$($versionfile[1]).$($versionfile[3]).$($versionfile[5]).$($versionfile[7])"

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch {
        Throw $_
    }
}

function Get-AdobeAcrobatReaderDCAdmxOnline {
<#
    .SYNOPSIS
    Returns latest Version and Uri for the Adobe AcrobatReaderDC Admx files
#>

    try {
        $ProgressPreference = 'SilentlyContinue'
        $file = "ReaderADMTemplate.zip"
        $url = "ftp://ftp.adobe.com/pub/adobe/reader/win/AcrobatDC/misc/"

        # grab ftp response from $url
        Write-Verbose "FTP $($url)"
        $listRequest = [Net.WebRequest]::Create($url)
        $listRequest.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectoryDetails
        $lines = New-Object System.Collections.ArrayList

        # process response
        $listResponse = $listRequest.GetResponse()
        $listStream = $listResponse.GetResponseStream()
        $listReader = New-Object System.IO.StreamReader($listStream)
        while (!$listReader.EndOfStream)
        {
            $line = $listReader.ReadLine()
            if ($line.Contains($file)) { $lines.Add($line) | Out-Null }
        }
        $listReader.Dispose()
        $listStream.Dispose()
        $listResponse.Dispose()

        Write-Verbose "received $($line.Length) characters response"
        
        # parse response to get Version
        $tokens = $lines[0].Split(" ", 9, [StringSplitOptions]::RemoveEmptyEntries)
        $Version = Get-Date -Date "$($tokens[6])/$($tokens[5])/$($tokens[7])" -Format "yy.M.d"

        # return evergreen object
        return @{ Version = $Version; URI = "$($url)$($file)" }
    }
    catch {
        Throw $_
    }
}

function Get-CitrixWorkspaceAppAdmxOnline {
<#
    .SYNOPSIS
    Returns latest Version and Uri for Citrix Workspace App ADMX files
#>

    try {
        $ProgressPreference = 'SilentlyContinue'
        $url = "https://www.citrix.com/downloads/workspace-app/windows/workspace-app-for-windows-latest.html"
        # grab content
        $web = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Ignore
        # find line with ADMX download
        $str = ($web.Content -split "`r`n" | Select-String -Pattern "_ADMX_")[0].ToString().Trim()
        # extract url from ADMX download string
        $URI = "https:$(((Select-String '(\/\/)([^\s,]+)(?=")' -Input $str).Matches.Value))"
        # grab version
        $Version = ($URI.Split("/")[-1] | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch {
        Throw $_
    }
}

function Get-MozillaFirefoxAdmxOnline {
<#
    .SYNOPSIS
    Returns latest Version and Uri for Mozilla Firefox ADMX files
#>

    try {
        $ProgressPreference = 'SilentlyContinue'
        # define github repo
        $repo = "mozilla/policy-templates"
        # grab latest release properties
        $latest = (Invoke-WebRequest -Uri "https://api.github.com/repos/$($repo)/releases" -UseBasicParsing | ConvertFrom-Json)[0]

        # grab version
        $Version = ($latest.tag_name | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        # grab uri
        $URI = $latest.assets.browser_download_url

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch {
        Throw $_
    }
}

function Get-ZoomDesktopClientAdmxOnline {
<#
    .SYNOPSIS
    Returns latest Version and Uri for Zoom Desktop Client ADMX files
#>

    try {
        $ProgressPreference = 'SilentlyContinue'
        $url = "https://support.zoom.us/hc/en-us/articles/360039100051"
        # grab content
        $web = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Ignore
        # find ADMX download
        $URI = (($web.Links | Where-Object {$_.href -like "*msi-templates*.zip"})[-1]).href
        # grab version
        $Version = ($URI.Split("/")[-1] | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch {
        Throw $_
    }
}

function Get-BIS-FAdmxOnline {
<#
    .SYNOPSIS
    Returns latest Version and Uri for BIS-F ADMX files
#>

    try {
        $ProgressPreference = 'SilentlyContinue'
        # define github repo
        $repo = "EUCweb/BIS-F"
        # grab latest release properties
        $latest = (Invoke-WebRequest -Uri "https://api.github.com/repos/$($repo)/releases" -UseBasicParsing | ConvertFrom-Json)[0]

        # grab version
        $Version = ($latest.tag_name | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        # grab uri
        $URI = $latest.zipball_url

        # return evergreen object
        return @{ Version = $Version; URI = $URI }
    }
    catch {
        Throw $_
    }
}

function Get-MDOPAdmxOnline {
<#
    .SYNOPSIS
    Returns latest Version and Uri for the Desktop Optimization Pack Admx files (both x64 and x86)
#>

    $id = "55531"
    $urlversion = "https://www.microsoft.com/en-us/download/details.aspx?id=$($id)"
    $urldownload = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=$($id)"
    try {
        $ProgressPreference = 'SilentlyContinue'
        # load page for version scrape
        $web = Invoke-WebRequest -UseBasicParsing -Uri $urlversion -ErrorAction SilentlyContinue
        $str = ($web.ToString() -split "[`r`n]" | Select-String "Version:").ToString()
        # grab version
        $Version = ($str | Select-String -Pattern "(\d+(\.\d+){1,4})" -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }).ToString()
        # load page for uri scrape
        $web = Invoke-WebRequest -UseBasicParsing -Uri $urldownload -ErrorAction SilentlyContinue -MaximumRedirection 0
        # grab download url
        $href = $web.Links | Where-Object { $_.outerHTML -like "*click here to download manually*" }

        # return evergreen object
        return @{ Version = $Version; URI = $href.href }
    }
    catch {
        Throw $_
    }
}

function Get-FSLogixAdmx {
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
        [string]$PolicyStore = $null
    )

    $evergreen = Get-FSLogixOnline
    $productname = "FSLogix"
    $productfolder = ""; if ($UseProductFolders) { $productfolder = "\$($productname)" }

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($evergreen.Version) for '$($productname)'"

        # download and process
        $outfile = "$($WorkingDirectory)\downloads\$($evergreen.URI.Split("/")[-1])"
        try {
            # download
            $ProgressPreference = 'SilentlyContinue'
            Write-Verbose "Downloading '$($evergreen.URI)' to '$($outfile)'"
            Invoke-WebRequest -Uri $evergreen.URI -UseBasicParsing -OutFile $outfile

            # extract
            Write-Verbose "Extracting '$($outfile)' to '$($env:TEMP)\fslogix'"
            Expand-Archive -Path $outfile -DestinationPath "$($env:TEMP)\fslogix" -Force

            # copy
            $sourceadmx = "$($env:TEMP)\fslogix"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            if (-not (Test-Path -Path "$($targetadmx)\en-US")) { $null = (mkdir -Path "$($targetadmx)\en-US" -Force) }

            Write-Verbose "Copying Admx files from '$($sourceadmx)' to '$($targetadmx)'"
            Copy-Item -Path "$($sourceadmx)\*.admx" -Destination "$($targetadmx)" -Force
            Copy-Item -Path "$($sourceadmx)\*.adml" -Destination "$($targetadmx)\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from '$($sourceadmx)' to '$($PolicyStore)'"
                Copy-Item -Path "$($sourceadmx)\*.admx" -Destination "$($PolicyStore)" -Force
                if (-not (Test-Path -Path "$($PolicyStore)\en-US")) { $null = (mkdir -Path "$($PolicyStore)\en-US" -Force) }
                Copy-Item -Path "$($sourceadmx)\*.adml" -Destination "$($PolicyStore)\en-US" -Force
            }

            # cleanup
            Remove-Item -Path "$($env:TEMP)\fslogix" -Recurse -Force

            return $evergreen
        }
        catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-MicrosoftOfficeAdmx {
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
        [string]$Architecture = "x64"
    )

    $evergreen = Get-MicrosoftOfficeAdmxOnline | Where-Object { $_.Architecture -like $Architecture }
    $productname = "Microsoft Office $($Architecture)"
    $productfolder = ""; if ($UseProductFolders) { $productfolder = "\$($productname)" }

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($evergreen.Version) for '$($productname)'"

        # download and process
        $outfile = "$($WorkingDirectory)\downloads\$($evergreen.URI.Split("/")[-1])"
        try {
            # download
            $ProgressPreference = 'SilentlyContinue'
            Write-Verbose "Downloading '$($evergreen.URI)' to '$($outfile)'"
            Invoke-WebRequest -Uri $evergreen.URI -UseBasicParsing -OutFile $outfile

            # extract
            Write-Verbose "Extracting '$($outfile)' to '$($env:TEMP)\office'"
            $null = Start-Process -FilePath $outfile -ArgumentList "/quiet /norestart /extract:`"$($env:TEMP)\office`"" -PassThru -Wait

            # copy
            $sourceadmx = "$($env:TEMP)\office\admx"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName $productname

            # cleanup
            Remove-Item -Path "$($env:TEMP)\office" -Recurse -Force

            return $evergreen
        }
        catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-Windows10Admx {
<#
    .SYNOPSIS
    Process Windows 10 Admx files

    .PARAMETER Version
    Current Version present

    .PARAMETER PolicyStore
    Destination for the Admx files

    .PARAMETER WindowsVersion
    Official WindowsVersion format
#>

    param(
        [string]$Version,
        [string]$PolicyStore = $null,
        [string]$WindowsVersion
    )
    
    $id = Get-Windows10AdmxDownloadId -WindowsVersion $WindowsVersion
    $evergreen = Get-Windows10AdmxOnline -DownloadId $id
    $productname = "Microsoft Windows 10 $($WindowsVersion)"
    $productfolder = ""; if ($UseProductFolders) { $productfolder = "\$($productname)" }

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($evergreen.Version) for '$($productname)'"

        # download and process
        $outfile = "$($WorkingDirectory)\downloads\$($evergreen.URI.Split("/")[-1])"
        try {
            # download
            $ProgressPreference = 'SilentlyContinue'
            Write-Verbose "Downloading '$($evergreen.URI)' to '$($outfile)'"
            Invoke-WebRequest -Uri $evergreen.URI -UseBasicParsing -OutFile $outfile

            # install
            Write-Verbose "Installing downloaded Windows 10 Admx installer"
            $null = Start-Process -FilePath "MsiExec.exe" -WorkingDirectory "$($WorkingDirectory)\downloads" -ArgumentList "/qn /norestart /I`"$($outfile.split('\')[-1])`"" -PassThru -Wait

            # find installation path
            Write-Verbose "Grabbing installation path for Windows 10 Admx installer"
            $installfolder = Get-ChildItem -Path "C:\Program Files (x86)\Microsoft Group Policy"
            Write-Verbose "Found '$($installfolder.Name)'"

            # find uninstall info
            Write-Verbose "Grabbing uninstallation info from registry for Windows 10 Admx installer"
            $uninstall = Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName -like "*(.admx)*" }
            Write-Verbose "Found '$($uninstall.DisplayName)'"

            # copy
            $sourceadmx = "C:\Program Files (x86)\Microsoft Group Policy\$($installfolder.Name)\PolicyDefinitions"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName $productname

            # uninstall
            Write-Verbose "Uninstalling Windows 10 Admx installer"
            $null = Start-Process -FilePath "MsiExec.exe" -ArgumentList "/qn /norestart /X$($uninstall.PSChildName)" -PassThru -Wait

            return $evergreen
        }
        catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-OneDriveAdmx {
<#
    .SYNOPSIS
    Process OneDrive Admx files

    .PARAMETER Version
    Current Version present

    .PARAMETER PolicyStore
    Destination for the Admx files
#>

    param(
        [string]$Version,
        [string]$PolicyStore = $null
    )

    $evergreen = Get-OneDriveOnline
    $productname = "Microsoft OneDrive"
    $productfolder = ""; if ($UseProductFolders) { $productfolder = "\$($productname)" }

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($evergreen.Version) for '$($productname)'"

        # download and process
        $outfile = "$($WorkingDirectory)\downloads\$($evergreen.URI.Split("/")[-1])"
        try {
            # download
            $ProgressPreference = 'SilentlyContinue'
            Write-Verbose "Downloading '$($evergreen.URI)' to '$($outfile)'"
            Invoke-WebRequest -Uri $evergreen.URI -UseBasicParsing -OutFile $outfile

            # install
            Write-Verbose "Installing downloaded OneDrive installer"
            $null = Start-Process -FilePath $outfile -ArgumentList "/allusers" -PassThru
            # wait for setup to complete
            while (Get-Process -Name "OneDriveSetup" -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 10 }
            # onedrive starts automatically after setup. kill!
            Stop-Process -Name "OneDrive" -Force

            # find uninstall info
            Write-Verbose "Grabbing uninstallation info from registry for OneDrive installer"
            $uninstall = Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\OneDriveSetup.exe"
            Write-Verbose "Found '$($uninstall.DisplayName)'"

            # find installation path
            Write-Verbose "Grabbing installation path for OneDrive installer"
            $installfolder = $uninstall.DisplayIcon.Substring(0, $uninstall.DisplayIcon.IndexOf("\OneDriveSetup.exe"))
            Write-Verbose "Found '$($installfolder)'"

            # copy
            $sourceadmx = "$($installfolder)\adm"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            if (-not (Test-Path -Path "$($targetadmx)")) { $null = (mkdir -Path "$($targetadmx)" -Force) }

            Write-Verbose "Copying Admx files from '$($sourceadmx)' to '$($targetadmx)'"
            Copy-Item -Path "$($sourceadmx)\*.admx" -Destination "$($targetadmx)" -Force
            foreach ($language in $Languages) {
                if (-not (Test-Path -Path "$($sourceadmx)\$($language)") -and -not (Test-Path -Path "$($sourceadmx)\$($language.Substring(0,2))")) {
                    if ($language -notlike "en-us") { Write-Warning "Language '$($language)' not found for '$($productname)'. Processing 'en-US' instead." }
                    if (-not (Test-Path -Path "$($targetadmx)\en-US")) { $null = (mkdir -Path "$($targetadmx)\en-US" -Force) }
                    Copy-Item -Path "$($sourceadmx)\*.adml" -Destination "$($targetadmx)\en-US" -Force    
                } else {
                    $sourcelanguage = $language; if (-not (Test-Path -Path "$($sourceadmx)\$($language)")) { $sourcelanguage = $language.Substring(0,2) }
                    if (-not (Test-Path -Path "$($targetadmx)\$($language)")) { $null = (mkdir -Path "$($targetadmx)\$($language)" -Force) }
                    Copy-Item -Path "$($sourceadmx)\$($sourcelanguage)\*.adml" -Destination "$($targetadmx)\$($language)" -Force
                }
            }

            # uninstall
            Write-Verbose "Uninstalling OneDrive installer"
            $null = Start-Process -FilePath "$($installfolder)\OneDriveSetup.exe" -ArgumentList "/uninstall /allusers" -PassThru -Wait

            return $evergreen
        }
        catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-MicrosoftEdgeAdmx {
<#
    .SYNOPSIS
    Process Microsoft Edge (Chromium) Admx files

    .PARAMETER Version
    Current Version present

    .PARAMETER PolicyStore
    Destination for the Admx files
#>

    param(
        [string]$Version,
        [string]$PolicyStore = $null
    )

    $evergreen = Get-MicrosoftEdgePolicyOnline
    $productname = "Microsoft Edge (Chromium)"
    $productfolder = ""; if ($UseProductFolders) { $productfolder = "\$($productname)" }

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($evergreen.Version) for '$($productname)'"

        # download and process
        $outfile = "$($WorkingDirectory)\downloads\$($evergreen.URI.Split("/")[-1])"
        try {
            # download
            $ProgressPreference = 'SilentlyContinue'
            Write-Verbose "Downloading '$($evergreen.URI)' to '$($outfile)'"
            Invoke-WebRequest -Uri $evergreen.URI -UseBasicParsing -OutFile $outfile

            # extract
            Write-Verbose "Extracting '$($outfile)' to '$($env:TEMP)\microsoftedgepolicy'"
            Expand-Archive -Path $outfile -DestinationPath "$($env:TEMP)\microsoftedgepolicy" -Force

            # copy
            $sourceadmx = "$($env:TEMP)\microsoftedgepolicy\windows\admx"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName $productname

            # cleanup
            Remove-Item -Path "$($env:TEMP)\microsoftedgepolicy" -Recurse -Force

            return $evergreen
        }
        catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-GoogleChromeAdmx {
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
        [string]$PolicyStore = $null
    )

    $evergreen = Get-GoogleChromeAdmxOnline
    $productname = "Google Chrome"
    $productfolder = ""; if ($UseProductFolders) { $productfolder = "\$($productname)" }

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($evergreen.Version) for '$($productname)'"

        # download and process
        $outfile = "$($WorkingDirectory)\downloads\googlechromeadmx.zip"
        try {
            # download
            $ProgressPreference = 'SilentlyContinue'
            Write-Verbose "Downloading '$($evergreen.URI)' to '$($outfile)'"
            Invoke-WebRequest -Uri $evergreen.URI -UseBasicParsing -OutFile $outfile

            # extract
            Write-Verbose "Extracting '$($outfile)' to '$($env:TEMP)\chromeadmx'"
            Expand-Archive -Path $outfile -DestinationPath "$($env:TEMP)\chromeadmx" -Force

            # copy
            $sourceadmx = "$($env:TEMP)\chromeadmx\windows\admx"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName $productname

            # cleanup
            Remove-Item -Path "$($env:TEMP)\chromeadmx" -Recurse -Force

            # chrome update admx is a seperate download
            $url = "https://dl.google.com/dl/update2/enterprise/googleupdateadmx.zip"

            # download
            $outfile = "$($WorkingDirectory)\downloads\googlechromeupdateadmx.zip"
            Write-Verbose "Downloading '$($url)' to '$($outfile)'"
            Invoke-WebRequest -Uri $url -UseBasicParsing -OutFile $outfile

            # extract
            Write-Verbose "Extracting '$($outfile)' to '$($env:TEMP)\chromeupdateadmx'"
            Expand-Archive -Path $outfile -DestinationPath "$($env:TEMP)\chromeupdateadmx" -Force

            # copy
            $sourceadmx = "$($env:TEMP)\chromeupdateadmx\GoogleUpdateAdmx"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName $productname -Quiet

            # cleanup
            Remove-Item -Path "$($env:TEMP)\chromeupdateadmx" -Recurse -Force

            return $evergreen
        }
        catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-AdobeAcrobatReaderDCAdmx {
<#
    .SYNOPSIS
    Process Adobe AcrobatReaderDC Admx files

    .PARAMETER Version
    Current Version present

    .PARAMETER PolicyStore
    Destination for the Admx files
#>

    param(
        [string]$Version,
        [string]$PolicyStore = $null
    )

    $evergreen = Get-AdobeAcrobatReaderDCAdmxOnline
    $productname = "Adobe Acrobat Reader DC"
    $productfolder = ""; if ($UseProductFolders) { $productfolder = "\$($productname)" }

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($evergreen.Version) for '$($productname)'"

        # download and process
        $outfile = "$($WorkingDirectory)\downloads\$($evergreen.URI.Split("/")[-1])"
        try {
            # download
            $ProgressPreference = 'SilentlyContinue'
            Write-Verbose "Downloading '$($evergreen.URI)' to '$($outfile)'"
            Invoke-WebRequest -Uri $evergreen.URI -UseBasicParsing -OutFile $outfile

            # extract
            Write-Verbose "Extracting '$($outfile)' to '$($env:TEMP)\acrobatreaderdc'"
            Expand-Archive -Path $outfile -DestinationPath "$($env:TEMP)\acrobatreaderdc" -Force

            # copy
            $sourceadmx = "$($env:TEMP)\acrobatreaderdc"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName $productname

            # cleanup
            Remove-Item -Path "$($env:TEMP)\acrobatreaderdc" -Recurse -Force

            return $evergreen
        }
        catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-CitrixWorkspaceAppAdmx {
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
        [string]$PolicyStore = $null
    )

    $evergreen = Get-CitrixWorkspaceAppAdmxOnline
    $productname = "Citrix Workspace App"
    $productfolder = ""; if ($UseProductFolders) { $productfolder = "\$($productname)" }

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($evergreen.Version) for '$($productname)'"

        # download and process
        $outfile = "$($WorkingDirectory)\downloads\$($evergreen.URI.Split("/")[-1].Split("?")[0])"
        try {
            # download
            $ProgressPreference = 'SilentlyContinue'
            Write-Verbose "Downloading '$($evergreen.URI)' to '$($outfile)'"
            Invoke-WebRequest -Uri $evergreen.URI -UseBasicParsing -OutFile $outfile

            # extract
            Write-Verbose "Extracting '$($outfile)' to '$($env:TEMP)\citrixworkspaceapp'"
            Expand-Archive -Path $outfile -DestinationPath "$($env:TEMP)\citrixworkspaceapp" -Force

            # copy
            $sourceadmx = "$($env:TEMP)\citrixworkspaceapp\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1].Split("?")[0]))\Configuration"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName $productname

            # cleanup
            Remove-Item -Path "$($env:TEMP)\citrixworkspaceapp" -Recurse -Force

            return $evergreen
        }
        catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-MozillaFirefoxAdmx {
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
        [string]$PolicyStore = $null
    )

    $evergreen = Get-MozillaFirefoxAdmxOnline
    $productname = "Mozilla Firefox"
    $productfolder = ""; if ($UseProductFolders) { $productfolder = "\$($productname)" }

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($evergreen.Version) for '$($productname)'"

        # download and process
        $outfile = "$($WorkingDirectory)\downloads\$($evergreen.URI.Split("/")[-1])"
        try {
            # download
            $ProgressPreference = 'SilentlyContinue'
            Write-Verbose "Downloading '$($evergreen.URI)' to '$($outfile)'"
            Invoke-WebRequest -Uri $evergreen.URI -UseBasicParsing -OutFile $outfile

            # extract
            Write-Verbose "Extracting '$($outfile)' to '$($env:TEMP)\firefoxadmx'"
            Expand-Archive -Path $outfile -DestinationPath "$($env:TEMP)\firefoxadmx" -Force

            # copy
            $sourceadmx = "$($env:TEMP)\firefoxadmx\windows"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName $productname

            # cleanup
            Remove-Item -Path "$($env:TEMP)\firefoxadmx" -Recurse -Force

            return $evergreen
        }
        catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-ZoomDesktopClientAdmx {
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
        [string]$PolicyStore = $null
    )

    $evergreen = Get-ZoomDesktopClientAdmxOnline
    $productname = "Zoom Desktop Client"
    $productfolder = ""; if ($UseProductFolders) { $productfolder = "\$($productname)" }

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($evergreen.Version) for '$($productname)'"

        # download and process
        $outfile = "$($WorkingDirectory)\downloads\$($evergreen.URI.Split("/")[-1])"
        try {
            # download
            $ProgressPreference = 'SilentlyContinue'
            Write-Verbose "Downloading '$($evergreen.URI)' to '$($outfile)'"
            Invoke-WebRequest -Uri $evergreen.URI -UseBasicParsing -OutFile $outfile

            # extract
            Write-Verbose "Extracting '$($outfile)' to '$($env:TEMP)\zoomclientadmx'"
            Expand-Archive -Path $outfile -DestinationPath "$($env:TEMP)\zoomclientadmx" -Force

            # copy
            $sourceadmx = "$($env:TEMP)\zoomclientadmx\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1]))"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName $productname

            # cleanup
            Remove-Item -Path "$($env:TEMP)\zoomclientadmx" -Recurse -Force

            return $evergreen
        }
        catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-BIS-FAdmx {
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

    $evergreen = Get-BIS-FAdmxOnline
    $productname = "BIS-F"
    $productfolder = ""; if ($UseProductFolders) { $productfolder = "\$($productname)" }

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($evergreen.Version) for '$($productname)'"

        # download and process
        $outfile = "$($WorkingDirectory)\downloads\bis-f.$($evergreen.Version).zip"
        try {
            # download
            $ProgressPreference = 'SilentlyContinue'
            Write-Verbose "Downloading '$($evergreen.URI)' to '$($outfile)'"
            Invoke-WebRequest -Uri $evergreen.URI -UseBasicParsing -OutFile $outfile

            # extract
            Write-Verbose "Extracting '$($outfile)' to '$($env:TEMP)\bisfadmx'"
            Expand-Archive -Path $outfile -DestinationPath "$($env:TEMP)\bisfadmx" -Force

            # find extraction folder
            Write-Verbose "Finding extraction folder"
            $folder = (Get-ChildItem -Path "$($env:TEMP)\bisfadmx" | Sort-Object LastWriteTime -Descending)[0].Name

            # copy
            $sourceadmx = "$($env:TEMP)\bisfadmx\$($folder)\admx"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName $productname

            # cleanup
            Remove-Item -Path "$($env:TEMP)\bisfadmx" -Recurse -Force

            return $evergreen
        }
        catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}

function Get-MDOPAdmx {
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
        [string]$PolicyStore = $null
    )

    $evergreen = Get-MDOPAdmxOnline
    $productname = "Microsoft Desktop Optimization Pack"
    $productfolder = ""; if ($UseProductFolders) { $productfolder = "\$($productname)" }

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
        Write-Verbose "Found new version $($evergreen.Version) for '$($productname)'"

        # download and process
        $outfile = "$($WorkingDirectory)\downloads\$($evergreen.URI.Split("/")[-1])"
        try {
            # download
            $ProgressPreference = 'SilentlyContinue'
            Write-Verbose "Downloading '$($evergreen.URI)' to '$($outfile)'"
            Invoke-WebRequest -Uri $evergreen.URI -UseBasicParsing -OutFile $outfile

            # extract
            Write-Verbose "Extracting '$($outfile)' to '$($env:TEMP)\mdopadmx'"
            $null = (mkdir -Path "$($env:TEMP)\mdopadmx" -Force)
            $null = (expand "$($outfile)" -F:* "$($env:TEMP)\mdopadmx")

            # find app-v folder
            Write-Verbose "Finding App-V folder"
            $appvfolder = (Get-ChildItem -Path "$($env:TEMP)\mdopadmx" -Filter "App-V*" | Sort-Object Name -Descending)[0].Name

            Write-Verbose "Finding MBAM folder"
            $mbamfolder = (Get-ChildItem -Path "$($env:TEMP)\mdopadmx" -Filter "MBAM*" | Sort-Object Name -Descending)[0].Name

            Write-Verbose "Finding UE-V folder"
            $uevfolder = (Get-ChildItem -Path "$($env:TEMP)\mdopadmx" -Filter "UE-V*" | Sort-Object Name -Descending)[0].Name

            # copy
            $sourceadmx = "$($env:TEMP)\mdopadmx\$($appvfolder)"
            $targetadmx = "$($WorkingDirectory)\admx$($productfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName "$($productname) - App-V"
            $sourceadmx = "$($env:TEMP)\mdopadmx\$($mbamfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName "$($productname) - MBAM"
            $sourceadmx = "$($env:TEMP)\mdopadmx\$($uevfolder)"
            Copy-Admx -SourceFolder $sourceadmx -TargetFolder $targetadmx -PolicyStore $PolicyStore -ProductName "$($productname) - UE-V"

            # cleanup
            Remove-Item -Path "$($env:TEMP)\mdopadmx" -Recurse -Force

            return $evergreen
        }
        catch {
            Throw $_
        }
    } else {
        # version already processed
        return $null
    }
}
#endregion

Write-Verbose "`nProcessing Admx files for Windows 10 $($WindowsVersion)"
$admx = Get-Windows10Admx -Version $admxversions.Windows.Version -PolicyStore $PolicyStore -WindowsVersion $WindowsVersion
if ($admx) { if ($admxversions.Windows) { $admxversions.Windows = $admx } else { $admxversions += @{ Windows = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "`nProcessing Admx files for Microsoft Edge (Chromium)"
$admx = Get-MicrosoftEdgeAdmx -Version $admxversions.Edge.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.Edge) { $admxversions.Edge = $admx } else { $admxversions += @{ Edge = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "`nProcessing Admx files for Microsoft OneDrive"
$admx = Get-OneDriveAdmx -Version $admxversions.OneDrive.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.OneDrive) { $admxversions.OneDrive = $admx } else { $admxversions += @{ OneDrive = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "`nProcessing Admx files for Microsoft Office"
$admx = Get-MicrosoftOfficeAdmx -Version $admxversions.Office.Version -PolicyStore $PolicyStore -Architecture "x64"
if ($admx) { if ($admxversions.Office) { $admxversions.Office = $admx } else { $admxversions += @{ Office = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "`nProcessing Admx files for FSLogix"
$admx = Get-FSLogixAdmx -Version $admxversions.FSLogix.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.FSLogix) { $admxversions.FSLogix = $admx } else { $admxversions += @{ FSLogix = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "`nProcessing Admx files for Adobe AcrobatReader DC"
$admx = Get-AdobeAcrobatReaderDCAdmx -Version $admxversions.AdobeAcrobatReaderDC.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.AdobeAcrobatReaderDC) { $admxversions.AdobeAcrobatReaderDC = $admx } else { $admxversions += @{ AdobeAcrobatReaderDC = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "`nProcessing Admx files for BIS-F"
$admx = Get-BIS-FAdmx -Version $admxversions.BISF.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.BISF) { $admxversions.BISF = $admx } else { $admxversions += @{ BISF = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "`nProcessing Admx files for Citrix Workspace App"
$admx = Get-CitrixWorkspaceAppAdmx -Version $admxversions.CitrixWorkspaceApp.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.CitrixWorkspaceApp) { $admxversions.CitrixWorkspaceApp = $admx } else { $admxversions += @{ CitrixWorkspaceApp = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "`nProcessing Admx files for Google Chrome"
$admx = Get-GoogleChromeAdmx -Version $admxversions.GoogleChrome.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.GoogleChrome) { $admxversions.GoogleChrome = $admx } else { $admxversions += @{ GoogleChrome = @{ Version = $admx.Version; URI = $admx.URI } } } }

#Write-Verbose "`nProcessing Admx files for Microsoft Desktop Optimization Pack"
#$admx = Get-MDOPAdmx -Version $admxversions.MDOP.Version -PolicyStore $PolicyStore
#if ($admx) { if ($admxversions.MDOP) { $admxversions.MDOP = $admx } else { $admxversions += @{ MDOP = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "`nProcessing Admx files for Mozilla Firefox"
$admx = Get-MozillaFirefoxAdmx -Version $admxversions.MozillaFirefox.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.MozillaFirefox) { $admxversions.MozillaFirefox = $admx } else { $admxversions += @{ MozillaFirefox = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "`nProcessing Admx files for Zoom Desktop Client"
$admx = Get-ZoomDesktopClientAdmx -Version $admxversions.ZoomDesktopClient.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.ZoomDesktopClient) { $admxversions.ZoomDesktopClient = $admx } else { $admxversions += @{ ZoomDesktopClient = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "`nSaving Admx versions to '$($WorkingDirectory)\admxversions.xml'"
$admxversions | Export-Clixml -Path "$($WorkingDirectory)\admxversions.xml" -Force
