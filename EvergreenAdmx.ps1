<#
    .SYNOPSIS
    Script to download latest Admx files for several products.

    .DESCRIPTION
    Script to download latest Admx files for several products.
    Optionally copy the latest Admx files to a folder of your chosing, for example a Policy Store.

    .PARAMETER WorkingDirectory
    Optionally provide a Working Directory for the script.
    The script will store Admx files in a subdirectory called "admx".
    The script will store downloaded files in a subdirectory called "downloads".
    If omitted the script will treat the script's folder as the working directory.

    .PARAMETER PolicyStore
    Optionally provide a Policy Store location to copy the Admx files to after processing.

    .PARAMETER WindowsVersion
    The Windows 10 version to get the Admx files for.


    .EXAMPLE
    .\EvergreenAdmx.ps1 -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions" -WindowsVersion "20H2"

    .LINK
    https://msfreaks.wordpress.com

#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$WorkingDirectory = $null,
    [Parameter(Mandatory = $false)]
    [string]$PolicyStore = $null,
    [Parameter(Mandatory = $false)][ValidateSet("1903", "1909", "2004", "20H2")]
    [string]$WindowsVersion = "20H2"
)

#region init
$admxversions = $null
if (-not $WorkingDirectory) { $WorkingDirectory = $PSScriptRoot }
if (Test-Path -Path "$($WorkingDirectory)\admxversions.xml") { $admxversions = Import-Clixml -Path "$($WorkingDirectory)\admxversions.xml" }
if (-not (Test-Path -Path "$($WorkingDirectory)\admx\en-US")) { $null = mkdir -Path "$($WorkingDirectory)\admx\en-US" -Force }
if (-not (Test-Path -Path "$($WorkingDirectory)\downloads")) { $null = mkdir -Path "$($WorkingDirectory)\downloads" -Force }
if ($PolicyStore -and -not $PolicyStore.EndsWith("\")) { $policypath += "\" }

Write-Verbose "WorkingDirectory:`t$($WorkingDirectory)"
Write-Verbose "Admx path:`t$($WorkingDirectory)\admx"
Write-Verbose "Download path:`t$($WorkingDirectory)\downloads"
Write-Verbose "PolicyStore:`t$($PolicyStore)"
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
        $str = ($web.ToString() -split "[`r`n]" | select-string "Version:").ToString()
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
        $str = ($web.ToString() -split "[`r`n]" | select-string "Version:").ToString()
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
    Returns latest Version and Uri for Citrix Workspace App ADMX files
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

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
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
            Write-Verbose "Copying Admx files from '$($env:TEMP)\fslogix\' to '$($WorkingDirectory)\admx'"
            Copy-Item -Path "$($env:TEMP)\fslogix\*.admx" -Destination "$($WorkingDirectory)\admx" -Force
            Copy-Item -Path "$($env:TEMP)\fslogix\*.adml" -Destination "$($WorkingDirectory)\admx\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from '$($env:TEMP)\fslogix\' to '$($PolicyStore)'"
                Copy-Item -Path "$($env:TEMP)\fslogix\*.admx" -Destination "$($PolicyStore)" -Force
                Copy-Item -Path "$($env:TEMP)\fslogix\*.adml" -Destination "$($PolicyStore)\en-US" -Force
            }

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

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
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
            Write-Verbose "Copying Admx files from '$($env:TEMP)\office\admx\' to '$($WorkingDirectory)\admx'"
            Copy-Item -Path "$($env:TEMP)\office\admx\*.admx" -Destination "$($WorkingDirectory)\admx" -Force
            Copy-Item -Path "$($env:TEMP)\office\admx\en-us\*.adml" -Destination "$($WorkingDirectory)\admx\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from '$($env:TEMP)\office\admx\' to '$($PolicyStore)'"
                Copy-Item -Path "$($env:TEMP)\office\admx\*.admx" -Destination "$($PolicyStore)" -Force
                Copy-Item -Path "$($env:TEMP)\office\admx\en-us\*.adml" -Destination "$($PolicyStore)\en-US" -Force
            }

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

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
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
            Write-Verbose "Copying Admx files from 'C:\Program Files (x86)\Microsoft Group Policy\$($installfolder.Name)\PolicyDefinitions\' to '$($WorkingDirectory)\admx'"
            Copy-Item -Path "C:\Program Files (x86)\Microsoft Group Policy\$($installfolder.Name)\PolicyDefinitions\*.admx" -Destination "$($WorkingDirectory)\admx" -Force
            Copy-Item -Path "C:\Program Files (x86)\Microsoft Group Policy\$($installfolder.Name)\PolicyDefinitions\en-us\*.adml" -Destination "$($WorkingDirectory)\admx\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from 'C:\Program Files (x86)\Microsoft Group Policy\$($installfolder.Name)\PolicyDefinitions\' to '$($PolicyStore)'"
                Copy-Item -Path "C:\Program Files (x86)\Microsoft Group Policy\$($installfolder.Name)\PolicyDefinitions\*.admx" -Destination "$($PolicyStore)" -Force
                Copy-Item -Path "C:\Program Files (x86)\Microsoft Group Policy\$($installfolder.Name)\PolicyDefinitions\en-us\*.adml" -Destination "$($PolicyStore)\en-US" -Force
            }

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

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
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
            $installfolder = $uninstall.DisplayIcon.Substring(0, $uninstall.DisplayIcon.IndexOf("OneDriveSetup.exe"))
            Write-Verbose "Found '$($installfolder.Name)'"

            # copy
            Write-Verbose "Copying Admx files from '$($installfolder)adm\' to '$($WorkingDirectory)\admx'"
            Copy-Item -Path "$($installfolder)adm\*.admx" -Destination "$($WorkingDirectory)\admx" -Force
            Copy-Item -Path "$($installfolder)adm\*.adml" -Destination "$($WorkingDirectory)\admx\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from '$($installfolder)adm\' to '$($PolicyStore)'"
                Copy-Item -Path "$($installfolder)adm\*.admx" -Destination "$($PolicyStore)" -Force
                Copy-Item -Path "$($installfolder)adm\*.adml" -Destination "$($PolicyStore)\en-US" -Force
            }

            # uninstall
            Write-Verbose "Uninstalling OneDrive installer"
            $null = Start-Process -FilePath "$($installfolder)OneDriveSetup.exe" -ArgumentList "/uninstall /allusers" -PassThru -Wait

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

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
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
            Write-Verbose "Copying Admx files from '$($env:TEMP)\microsoftedgepolicy\windows\admx\' to '$($WorkingDirectory)\admx'"
            Copy-Item -Path "$($env:TEMP)\microsoftedgepolicy\windows\admx\*.admx" -Destination "$($WorkingDirectory)\admx" -Force
            Copy-Item -Path "$($env:TEMP)\microsoftedgepolicy\windows\admx\en-us\*.adml" -Destination "$($WorkingDirectory)\admx\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from '$($env:TEMP)\microsoftedgepolicy\windows\admx\' to '$($PolicyStore)'"
                Copy-Item -Path "$($env:TEMP)\microsoftedgepolicy\windows\admx\*.admx" -Destination "$($PolicyStore)" -Force
                Copy-Item -Path "$($env:TEMP)\microsoftedgepolicy\windows\admx\en-us\*.adml" -Destination "$($PolicyStore)\en-US" -Force
            }

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

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
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
            Write-Verbose "Copying Admx files from '$($env:TEMP)\chromeadmx\windows\admx\' to '$($WorkingDirectory)\admx'"
            Copy-Item -Path "$($env:TEMP)\chromeadmx\windows\admx\*.admx" -Destination "$($WorkingDirectory)\admx" -Force
            Copy-Item -Path "$($env:TEMP)\chromeadmx\windows\admx\en-us\*.adml" -Destination "$($WorkingDirectory)\admx\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from '$($env:TEMP)\chromeadmx\windows\admx\' to '$($PolicyStore)'"
                Copy-Item -Path "$($env:TEMP)\chromeadmx\windows\admx\*.admx" -Destination "$($PolicyStore)" -Force
                Copy-Item -Path "$($env:TEMP)\chromeadmx\windows\admx\en-us\*.adml" -Destination "$($PolicyStore)\en-US" -Force
            }

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
            Write-Verbose "Copying Admx files from '$($env:TEMP)\chromeupdateadmx\GoogleUpdateAdmx\' to '$($WorkingDirectory)\admx'"
            Copy-Item -Path "$($env:TEMP)\chromeupdateadmx\GoogleUpdateAdmx\*.admx" -Destination "$($WorkingDirectory)\admx" -Force
            Copy-Item -Path "$($env:TEMP)\chromeupdateadmx\GoogleUpdateAdmx\en-us\*.adml" -Destination "$($WorkingDirectory)\admx\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from '$($env:TEMP)\chromeupdateadmx\GoogleUpdateAdmx\' to '$($PolicyStore)'"
                Copy-Item -Path "$($env:TEMP)\chromeupdateadmx\GoogleUpdateAdmx\*.admx" -Destination "$($PolicyStore)" -Force
                Copy-Item -Path "$($env:TEMP)\chromeupdateadmx\GoogleUpdateAdmx\en-us\*.adml" -Destination "$($PolicyStore)\en-US" -Force
            }

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

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
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
            Write-Verbose "Copying Admx files from '$($env:TEMP)\acrobatreaderdc\' to '$($WorkingDirectory)\admx'"
            Copy-Item -Path "$($env:TEMP)\acrobatreaderdc\*.admx" -Destination "$($WorkingDirectory)\admx" -Force
            Copy-Item -Path "$($env:TEMP)\acrobatreaderdc\*.adml" -Destination "$($WorkingDirectory)\admx\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from '$($env:TEMP)\acrobatreaderdc\' to '$($PolicyStore)'"
                Copy-Item -Path "$($env:TEMP)\acrobatreaderdc\*.admx" -Destination "$($PolicyStore)" -Force
                Copy-Item -Path "$($env:TEMP)\acrobatreaderdc\*.adml" -Destination "$($PolicyStore)\en-US" -Force
            }

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

    $evergreen = Get-CitrixWorkspaceAppAdmxOnline

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
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
            Write-Verbose "Copying Admx files from '$($env:TEMP)\citrixworkspaceapp\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1].Split("?")[0]))\Configuration\' to '$($WorkingDirectory)\admx'"
            Copy-Item -Path "$($env:TEMP)\citrixworkspaceapp\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1].Split("?")[0]))\Configuration\*.admx" -Destination "$($WorkingDirectory)\admx" -Force
            Copy-Item -Path "$($env:TEMP)\citrixworkspaceapp\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1].Split("?")[0]))\Configuration\en-us\*.adml" -Destination "$($WorkingDirectory)\admx\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from '$($env:TEMP)\citrixworkspaceapp\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1].Split("?")[0]))\Configuration\' to '$($PolicyStore)'"
                Copy-Item -Path "$($env:TEMP)\citrixworkspaceapp\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1].Split("?")[0]))\Configuration\*.admx" -Destination "$($PolicyStore)" -Force
                Copy-Item -Path "$($env:TEMP)\citrixworkspaceapp\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1].Split("?")[0]))\Configuration\en-us\*.adml" -Destination "$($PolicyStore)\en-US" -Force
            }

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

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
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
            Write-Verbose "Copying Admx files from '$($env:TEMP)\firefoxadmx\windows\' to '$($WorkingDirectory)\admx'"
            Copy-Item -Path "$($env:TEMP)\firefoxadmx\windows\*.admx" -Destination "$($WorkingDirectory)\admx" -Force
            Copy-Item -Path "$($env:TEMP)\firefoxadmx\windows\en-us\*.adml" -Destination "$($WorkingDirectory)\admx\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from '$($env:TEMP)\firefoxadmx\windows\' to '$($PolicyStore)'"
                Copy-Item -Path "$($env:TEMP)\firefoxadmx\windows\*.admx" -Destination "$($PolicyStore)" -Force
                Copy-Item -Path "$($env:TEMP)\firefoxadmx\windows\en-us\*.adml" -Destination "$($PolicyStore)\en-US" -Force
            }

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

    # see if this is a newer version
    if (-not $Version -or [version]$evergreen.Version -gt [version]$Version) {
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
            Write-Verbose "Copying Admx files from '$($env:TEMP)\zoomclientadmx\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1]))\' to '$($WorkingDirectory)\admx'"
            Copy-Item -Path "$($env:TEMP)\zoomclientadmx\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1]))\*.admx" -Destination "$($WorkingDirectory)\admx" -Force
            Copy-Item -Path "$($env:TEMP)\zoomclientadmx\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1]))\en-us\*.adml" -Destination "$($WorkingDirectory)\admx\en-US" -Force
            if ($PolicyStore) {
                Write-Verbose "Copying Admx files from '$($env:TEMP)\zoomclientadmx\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1]))\' to '$($PolicyStore)'"
                Copy-Item -Path "$($env:TEMP)\zoomclientadmx\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1]))\*.admx" -Destination "$($PolicyStore)" -Force
                Copy-Item -Path "$($env:TEMP)\zoomclientadmx\$([io.path]::GetFileNameWithoutExtension($evergreen.URI.Split("/")[-1]))\en-us\*.adml" -Destination "$($PolicyStore)\en-US" -Force
            }

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

Write-Verbose "Processing Admx files for Windows 10 $($WindowsVersion)"
$admx = Get-Windows10Admx -Version $admxversions.Windows.Version -PolicyStore $PolicyStore -WindowsVersion $WindowsVersion
if ($admx) { if ($admxversions.Windows) { $admxversions.Windows = $admx } else { $admxversions += @{ Windows = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "Processing Admx files for Microsoft Edge (Chromium)"
$admx = Get-MicrosoftEdgeAdmx -Version $admxversions.Edge.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.Edge) { $admxversions.Edge = $admx } else { $admxversions += @{ Edge = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "Processing Admx files for Microsoft OneDrive"
$admx = Get-OneDriveAdmx -Version $admxversions.OneDrive.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.OneDrive) { $admxversions.OneDrive = $admx } else { $admxversions += @{ OneDrive = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "Processing Admx files for Microsoft Office"
$admx = Get-MicrosoftOfficeAdmx -Version $admxversions.Office.Version -PolicyStore $PolicyStore -Architecture "x64"
if ($admx) { if ($admxversions.Office) { $admxversions.Office = $admx } else { $admxversions += @{ Office = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "Processing Admx files for FSLogix"
$admx = Get-FSLogixAdmx -Version $admxversions.FSLogix.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.FSLogix) { $admxversions.FSLogix = $admx } else { $admxversions += @{ FSLogix = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "Processing Admx files for Google Chrome"
$admx = Get-GoogleChromeAdmx -Version $admxversions.GoogleChrome.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.GoogleChrome) { $admxversions.GoogleChrome = $admx } else { $admxversions += @{ GoogleChrome = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "Processing Admx files for Adobe AcrobatReader DC"
$admx = Get-AdobeAcrobatReaderDCAdmx -Version $admxversions.AdobeAcrobatReaderDC.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.AdobeAcrobatReaderDC) { $admxversions.AdobeAcrobatReaderDC = $admx } else { $admxversions += @{ AdobeAcrobatReaderDC = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "Processing Admx files for Citrix Workspace App"
$admx = Get-CitrixWorkspaceAppAdmx -Version $admxversions.CitrixWorkspaceApp.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.CitrixWorkspaceApp) { $admxversions.CitrixWorkspaceApp = $admx } else { $admxversions += @{ CitrixWorkspaceApp = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "Processing Admx files for Mozilla Firefox"
$admx = Get-MozillaFirefoxAdmx -Version $admxversions.MozillaFirefox.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.MozillaFirefox) { $admxversions.MozillaFirefox = $admx } else { $admxversions += @{ MozillaFirefox = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "Processing Admx files for Zoom Desktop Client"
$admx = Get-ZoomDesktopClientAdmx -Version $admxversions.ZoomDesktopClient.Version -PolicyStore $PolicyStore
if ($admx) { if ($admxversions.ZoomDesktopClient) { $admxversions.ZoomDesktopClient = $admx } else { $admxversions += @{ ZoomDesktopClient = @{ Version = $admx.Version; URI = $admx.URI } } } }

Write-Verbose "Saving Admx versions to '$($WorkingDirectory)\admxversions.xml'"
$admxversions | Export-Clixml -Path "$($WorkingDirectory)\admxversions.xml" -Force
