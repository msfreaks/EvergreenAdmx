# EvergreenAdmx

[![Release][github-release-badge]][github-release]
[![Codacy][code-quality-badge]][code-quality]
[![Twitter][twitter-follow-badge]][twitter-follow]

After deploying several Azure Virtual Desktop environments I decided I no longer wanted to manually download the Admx files I needed, and I wanted a way to keep them up-to-date.

This script solves both problems.
*  Checks for newer versions of the Admx files that are present and processes the new version if found
*  Optionally copies the new Admx files to the Policy Store or Definition folder, or a folder of your choosing

The name I chose for this script is an ode to the Evergreen module (https://github.com/aaronparker/Evergreen) by Aaron Parker (@stealthpuppy).

## How to use

Quick start:
*  Download the script to a location of your choosing (for example: C:\Scripts\EvergreenAdmx)
*  Run or schedule the script

You can also install the script from the PowerShell Gallery ([EvergreenAdmx][poshgallery-evergreenadmx]):
```powershell
Install-Script -Name EvergreenAdmx
```

I have scheduled the script to run daily:

```powershell
EvergreenAdmx.ps1 -Windows11Version "23H2" -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions"
```

The above execution will keep the central Policy Store up-to-date on a daily basis.

A sample .xml file that you can import in Task Scheduler is provided with this script.

`Breaking change starting from 2301.1`

Valid entries are "Custom Policy Store", "Windows 10", "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft Office", "FSLogix", "Adobe Acrobat", "Adobe Reader", "BIS-F", "Citrix Workspace App", "Google Chrome", "Microsoft Desktop Optimization Pack", "Mozilla Firefox", "Zoom Desktop Client".

`Breaking change starting from 2112.1`

Windows 11 was added to the script, and that means the -WindowsVersion parameter was substituted with -Windows10Version and -Windows11Version.
These MUST be used and for the same Windows version as any 'Include' options you provide.

From this version the 'Includes' will default to "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft Office".

`Breaking change starting from 2101.2`

This script no longer processes all the products by default. There's no need to comment out any products you don't need anymore.

In 2101.2 the parameter 'Include' was introduced which is an array you can use to specify all products that need to be processed. This parameter is required for the script to be able to run.
Valid entries are "Custom Policy Store", "Windows 10", "Microsoft Edge", "Microsoft OneDrive", "Microsoft Office", "FSLogix", "Adobe AcrobatReader DC", "BIS-F", "Citrix Workspace App", "Google Chrome", "Microsoft Desktop Optimization Pack", "Mozilla Firefox", "Zoom Desktop Client".

By default, if you don't use this parameter, only "Windows 10", "Microsoft Edge", "Microsoft OneDrive", "Microsoft Office" is processed.

```
SYNTAX
    C:\gits\github\EvergreenAdmx\EvergreenAdmx.ps1 [[-Windows10Version] <String>] [[-Windows11Version]
    <String>] [[-WorkingDirectory] <String>] [[-PolicyStore] <String>] [[-Languages] <String[]>]
    [-UseProductFolders] [[-CustomPolicyStore] <String>] [[-Include] <String[]>] [-PreferLocalOneDrive]
    [<CommonParameters>]

DESCRIPTION
    Script to automatically download latest Admx files for several products.
    Optionally copies the latest Admx files to a folder of your choosing, for example a Policy Store.


PARAMETERS
    -Windows10Version <String>
       The Windows 10 version to get the Admx files for. This value will be ignored if 'Windows 10' is
       not specified with -Include parameter.
       If the -Include parameter contains 'Windows 10', the latest Windows 10 version will be used.
       Defaults to "Windows11Version" if omitted.

       Note: Windows 11 23H2 policy definitions now supports Windows 10.

        Required?                    false
        Position?                    1
        Default value                22H2
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -Windows11Version <String>
       The Windows 11 version to get the Admx files for. This value will be ignored if 'Windows 10' is
       not specified with -Include parameter.
       If omitted, defaults to latest version available.

        Required?                    false
        Position?                    1
        Default value                23H2
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -WorkingDirectory <String>
        Optionally provide a Working Directory for the script.
        The script will store Admx files in a subdirectory called "admx".
        The script will store downloaded files in a subdirectory called "downloads".
        If omitted the script will treat the script's folder as the working directory.

        Required?                    false
        Position?                    2
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -PolicyStore <String>
        Optionally provide a Policy Store location to copy the Admx files to after processing.

        Required?                    false
        Position?                    3
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -Languages <String[]>
        Optionally provide an array of languages to process. Entries must be in 'xy-XY' format.
        If omitted the script will process 'en-US'.

        Required?                    false
        Position?                    4
        Default value                @("en-US")
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -UseProductFolders [<SwitchParameter>]
        When specified the extracted Admx files are copied to their respective product folders in a subfolder of 'Admx' in the WorkingDirectory.

        Required?                    false
        Position?                    named
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -CustomPolicyStore <String>
        When specified processes a location for custom policy files. Can be UNC format or local folder.
        The script will expect to find .admx files in this location, and at least one language folder holding the .adml file(s).
        Versioning will be done based on the newest file found recursively in this location (any .admx or .adml).
        Note that if any file has changed the script will process all files found in location.

        Required?                    false
        Position?                    5
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -Include <String[]>
        Array containing Admx products to include when checking for updates.
        Defaults to "Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft Office" if omitted.

        Required?                    false
        Position?                    6
        Default value                @("Windows 11", "Microsoft Edge", "Microsoft OneDrive", "Microsoft Office")
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -PreferLocalOneDrive [<SwitchParameter>]
        Microsoft OneDrive Admx files are only available after installing OneDrive.
        If this script is running on a machine that has OneDrive installed locally, use this switch to prevent automatically uninstalling OneDrive.

        Required?                    false
        Position?                    named
        Default value                false
        Accept pipeline input?       false
        Accept wildcard characters?  false

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\>.\EvergreenAdmx.ps1 -Windows11Version "23H2" -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions" -Languages @("en-US", "nl-NL") -UseProductFolders

    Will process the default set of products, storing results in product folders, for both English United States as Dutch languages, and copies the files to the Policy store.

```

## Admx files

Also see [Change Log][change-log] for a list of supported products.

Now supports
*  Custom Policy Store
*  Adobe Acrobat
*  Adobe Reader
*  Base Image Script Framework (BIS-F)
*  Citrix Workspace App
*  FSLogix
*  Google Chrome
*  Microsoft Desktop Optimization Pack
*  Microsoft Edge (Chromium)
*  Microsoft Office
*  Microsoft OneDrive (installed or Evergreen)
*  Microsoft Windows 10 (1903/1909/2004/20H2/21H1/21H2/22H2)
*  Microsoft Windows 11 (21H2/22H2/23H2)
*  Mozilla Firefox
*  Zoom Desktop Client

## Notes

I have not tested this script on Windows Core.
Some of the Admx files can only be obtained by installing the package that was downloaded.
For instance, the Windows 10 and Windows 11 Admx files are in an msi file, the OneDrive Admx files are in the installation folder after installing OneDrive.
If you are going to use the script to download Windows 10 or Windows 11 Admx files, you will need to remove any installs of the Windows 10 or Windows 11 Admx msi, or the script will fail.
So this is what the script does for these packages: installing the package, copying the Admx files, uninstalling the package.

Thank you Jonathan Pitre (@PitreJonathan) for keeping me sharp, providing fixes and improvements!

[github-release-badge]: https://img.shields.io/github/v/release/msfreaks/EvergreenAdmx.svg?style=flat-square
[github-release]: https://github.com/msfreaks/EvergreenAdmx/releases/latest
[code-quality-badge]: https://app.codacy.com/project/badge/Grade/c0efab02b66442399bb16b0493cdfbef?style=flat-square
[code-quality]: https://www.codacy.com/gh/msfreaks/EvergreenAdmx/dashboard?utm_source=github.com&amp;utm_medium=referral&amp;utm_content=msfreaks/EvergreenAdmx&amp;utm_campaign=Badge_Grade
[twitter-follow-badge]: https://img.shields.io/twitter/follow/menschab?style=flat-square
[twitter-follow]: https://twitter.com/menschab?ref_src=twsrc%5Etfw
[change-log]: https://github.com/msfreaks/EvergreenAdmx/blob/main/CHANGELOG.md
[poshgallery-evergreenadmx]: https://www.powershellgallery.com/packages/EvergreenAdmx/
