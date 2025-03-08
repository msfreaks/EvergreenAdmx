# EvergreenAdmx

[![Release][github-release-badge]][github-release]
[![Codacy][code-quality-badge]][code-quality]
[![Twitter][twitter-follow-badge]][twitter-follow]

After deploying several Azure Virtual Desktop environments I decided I no longer wanted to manually download the Admx files I needed, and I wanted a way to keep them up-to-date.

This script solves both problems.

- Automatically checks for newer versions of ADMX files and processes them when found
- Optionally copies the new ADMX files to your Policy Store or a custom location

Named as an homage to the [Evergreen module](https://github.com/aaronparker/Evergreen) by Aaron Parker [@stealthpuppy](https://twitter.com/stealthpuppy).

## How to use

Quick start:

- Download the script to a location of your choosing (for example: C:\Scripts\EvergreenAdmx)
- Run or schedule the script

You can also install the script from the PowerShell Gallery [EvergreenAdmx][poshgallery-evergreenadmx] :

```powershell
Install-Script -Name EvergreenAdmx
```

### Examples

Download policy definitions files to Test directory

```powershell
.\EvergreenAdmx.ps1 -WorkingDirectory .\Test
```

Update local group policies policy store on your DC

```powershell
.\EvergreenAdmx.ps1 -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions"
```

Update local group policies

```powershell
.\EvergreenAdmx.ps1 -PolicyStore "C:\Windows\PolicyDefinitions"
```


```powershell
.\EvergreenAdmx.ps1 -WorkingDirectory .\Test -Include @('Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps', 'Microsoft FSLogix') -UseProductFolders
```

If you wish to run a scheduled task on a daily basis.
You can import the sample xml file in Task Scheduler provided with this script.

`Breaking change starting from 2503.1`

Valid entries are "Custom Policy Store", "Windows 10", "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft 365 Apps", "Microsoft FSLogix", "Adobe Acrobat", "Adobe Reader", "BIS-F", "Citrix Workspace App", "Google Chrome", "Microsoft Desktop Optimization Pack", "Mozilla Firefox", "Zoom", "Zoom VDI", "Microsoft AVD", "Microsoft Winget", "Brave Browser".

`Breaking change starting from 2402.1`

Valid entries are "Custom Policy Store", "Windows 10", "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft 365 Apps", "Microsoft FSLogix", "Adobe Acrobat", "Adobe Reader", "BIS-F", "Citrix Workspace App", "Google Chrome", "Microsoft Desktop Optimization Pack", "Mozilla Firefox", "Zoom Desktop Client", "Azure Virtual Desktop", "Microsoft Winget".

`Breaking change starting from 2301.1`

Valid entries are "Custom Policy Store", "Windows 10", "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft 365 Apps", "Microsoft FSLogix", "Adobe Acrobat", "Adobe Reader", "BIS-F", "Citrix Workspace App", "Google Chrome", "Microsoft Desktop Optimization Pack", "Mozilla Firefox", "Zoom Desktop Client".

`Breaking change starting from 2112.1`

Windows 11 was added to the script, and that means the -WindowsVersion parameter was substituted with -Windows10Version and -Windows11Version.
These MUST be used and for the same Windows version as any 'Include' options you provide.

From this version the 'Includes' will default to "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft 365 Apps".

`Breaking change starting from 2101.2`

This script no longer processes all the products by default. There's no need to comment out any products you don't need anymore.

In 2101.2 the parameter 'Include' was introduced which is an array you can use to specify all products that need to be processed. This parameter is required for the script to be able to run.
Valid entries are "Custom Policy Store", "Windows 10" or "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft 365 Apps", "Microsoft FSLogix", "Adobe Acrobat", "Adobe Reader", "BIS-F", "Citrix Workspace App", "Google Chrome", "Microsoft Desktop Optimization Pack", "Mozilla Firefox", "Zoom Desktop Client", "Azure Virtual Desktop".

By default, if you don't use this parameter, only "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft 365 Apps" is processed.

```powershell
SYNOPSIS
    Script to automatically download latest Admx files for several products.


SYNTAX
    N:\Intune\Scripts\EvergreenAdmx\EvergreenAdmx.ps1 [[-WindowsVersion] <String>] [-Windows10FeatureVersion <String>] [-Windows11FeatureVersion <String>] [-WorkingDirectory <String>] [-PolicyStore <String>] [-Languages <String[]>]
    [-UseProductFolders] [-AddAdmxPath <String>] [-Include <String[]>] [-PreferLocalOneDrive] [<CommonParameters>]


DESCRIPTION
    Script to automatically download latest Admx files for several products.
    Optionally copies the latest Admx files to a folder of your choosing, for example a Policy Store.


PARAMETERS
    -WindowsVersion <String>
        Specifies Windows major version. Supports 10, 11 or 2025.
        Default is 11.

        Required?                    false
        Position?                    1
        Default value                11
        Accept pipeline input?       false
        Aliases
        Accept wildcard characters?  false

    -Windows10FeatureVersion <String>
        Specifies Windows 10 feature version to get the Admx files for. This parameter is used when 'Windows 10' is included.
        Valid values are: 1903, 1909, 2004, 20H2, 21H1, 21H2, 22H2.
        Defaults to 22H2.

        Note: Windows 11 23H2 policy definitions now supports Windows 10.

        Required?                    false
        Position?                    named
        Default value                22H2
        Accept pipeline input?       false
        Aliases
        Accept wildcard characters?  false

    -Windows11FeatureVersion <String>
        Specifies Windows 11 feature version to get the Admx files for. This parameter is used when 'Windows 11' is included.
        Valid values are: 21H2, 22H2, 23H2, 24H2.
        Defaults to 24H2.

        Required?                    false
        Position?                    named
        Default value                24H2
        Accept pipeline input?       false
        Aliases
        Accept wildcard characters?  false

    -WorkingDirectory <String>
        Specifies a Working Directory for the script.
        Admx files will be stored in a subdirectory called "admx".
        Downloaded files will be stored in a subdirectory called "downloads".
        Defaults to current script location.

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Aliases
        Accept wildcard characters?  false

    -PolicyStore <String>
        Specifies a Policy Store location to copy the Admx files to after processing.

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Aliases
        Accept wildcard characters?  false

    -Languages <String[]>
        Specifies an array of languages to process. Entries must be in 'xy-XY' format.
        Defaults to 'en-US'.

        Required?                    false
        Position?                    named
        Default value                @('en-US')
        Accept pipeline input?       false
        Aliases
        Accept wildcard characters?  false

    -UseProductFolders [<SwitchParameter>]
        Admx files are copied to their respective product folders in a subfolder of 'Admx' in the WorkingDirectory.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Aliases
        Accept wildcard characters?  false

    -AddAdmxPath <String>
        Specifies a location for custom policy files. Can be UNC format or local folder.
        Find .admx files in this location, and at least one language folder holding the .adml file(s).
        Versioning will be done based on the newest file found recursively in this location (any .admx or .adml).
        Note that if any file has changed the script will process all files found in location.

        Required?                    false
        Position?                    named
        Default value
        Accept pipeline input?       false
        Aliases
        Accept wildcard characters?  false

    -Include <String[]>
        Array containing Admx products to include when checking for updates.
        Valid values are: "Windows 10", "Windows 11", "Windows 2025", "Microsoft Edge", "Microsoft OneDrive", "Microsoft 365 Apps", "Microsoft FSLogix", "Adobe Acrobat", "Adobe Reader", "BIS-F", "Citrix Workspace App", "Google Chrome", "Microsoft
        Desktop Optimization Pack", "Mozilla Firefox", "Zoom", "Zoom VDI", "Microsoft AVD", "Microsoft Winget", "Brave Browser".
        Defaults to "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft 365 Apps".

        Required?                    false
        Position?                    named
        Default value                @('Windows 11', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps')
        Accept pipeline input?       false
        Aliases
        Accept wildcard characters?  false

    -PreferLocalOneDrive [<SwitchParameter>]
        Microsoft OneDrive Admx files are only available after installing OneDrive.
        If this script is running on a machine that has OneDrive installed locally, use this switch to prevent automatically uninstalling OneDrive.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Aliases
        Accept wildcard characters?  false

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https://go.microsoft.com/fwlink/?LinkID=113216).

INPUTS

OUTPUTS

    -------------------------- EXAMPLE 1 --------------------------

    PS > .\EvergreenAdmx.ps1

    Downloads the latest admx files for Windows 11, Microsoft Edge, Microsoft OneDrive, and Microsoft 365 Apps to the current folder.




    -------------------------- EXAMPLE 2 --------------------------

    PS > .\EvergreenAdmx.ps1 -WorkingDirectory "C:\Temp\EvergreenAdmx" -Include @('Windows 11', 'Microsoft Edge', 'Microsoft OneDrive', 'Microsoft 365 Apps', 'Microsoft FSLogix')

    Downloads the latest admx files for the specified products to C:\Temp\EvergreenAdmx folder.




    -------------------------- EXAMPLE 3 --------------------------

    PS > .\EvergreenAdmx.ps1 -WindowsVersion 2025 -Include "Windows 2025"

    Downloads the latest admx files for Windows 2025 to the current folder.




    -------------------------- EXAMPLE 4 --------------------------

    PS > .\EvergreenAdmx.ps1 -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions" -Languages @("en-US", "nl-NL") -UseProductFolders

    Downloads the default set of admx files, stores them in product folders for both English and Dutch languages, and copies them to the specified Policy store.
```

## Admx files

Also see [Change Log][change-log] for a list of supported products.

Now supports

- Adobe Acrobat (Continuous Track)
- Adobe Reader (Continuous Track)
- Base Image Script Framework (BIS-F)
- Brave Browser
- Citrix Workspace App
- Custom Policy Store
- Google Chrome
- Microsoft AVD
- Microsoft Desktop Optimization Pack
- Microsoft Edge (Chromium)
- Microsoft FSLogix
- Microsoft 365 Apps
- Microsoft OneDrive (local installation or Evergreen)
- Microsoft Windows 10 (1903/1909/2004/20H2/21H1/21H2/22H2)
- Microsoft Windows 11 (21H2/22H2/23H2/24H2)
- Microsoft Windows Server 2025 (Nov 2024)
- Microsoft Winget
- Mozilla Firefox
- Zoom
- Zoom VDI

## Notes

I have not tested this script on Windows Core.
Some of the Admx files can only be obtained by installing the package that was downloaded.
For instance, the Windows 10 and Windows 11 Admx files are in an msi file, the OneDrive Admx files are in the installation folder after installing OneDrive.
If you are going to use the script to download Windows 10 or Windows 11 Admx files, you will need to remove any installs of the Windows 10 or Windows 11 Admx msi, or the script will fail.
So this is what the script does for these packages: installing the package, copying the Admx files, uninstalling the package.

Thank you [Jonathan Pitre](https://github.com/JonathanPitre) for keeping me sharp, providing fixes and improvements!

[github-release-badge]: https://img.shields.io/github/v/release/msfreaks/EvergreenAdmx.svg?style=flat-square
[github-release]: https://github.com/msfreaks/EvergreenAdmx/releases/latest
[code-quality-badge]: https://app.codacy.com/project/badge/Grade/c0efab02b66442399bb16b0493cdfbef?style=flat-square
[code-quality]: https://www.codacy.com/gh/msfreaks/EvergreenAdmx/dashboard?utm_source=github.com&amp;utm_medium=referral&amp;utm_content=msfreaks/EvergreenAdmx&amp;utm_campaign=Badge_Grade
[twitter-follow-badge]: https://img.shields.io/twitter/follow/menschab?style=flat-square
[twitter-follow]: https://twitter.com/menschab?ref_src=twsrc%5Etfw
[change-log]: https://github.com/msfreaks/EvergreenAdmx/blob/main/CHANGELOG.md
[poshgallery-evergreenadmx]: https://www.powershellgallery.com/packages/EvergreenAdmx/
