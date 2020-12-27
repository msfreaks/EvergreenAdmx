# About

[![License][license-badge]][license]
[![Release][github-release-badge]][github-release]

After deploying several Windows Virtual Desktop environments I decided I no longer wanted to manually download the Admx files I needed, and I wanted a way to keep them up-to-date.

This script solves both problems.
* Checks for newer versions of the Admx files that are present and processes the new version if found
* Optionally copies the new Admx files to the Policy Store or Definition folder, or a folder of your chosing

The name I chose for this script is an ode to the Evergreen module (https://github.com/aaronparker/Evergreen) by Aaron Parker (@stealthpuppy).

# How to use

Quick start:
* Download the script to a location of your chosing (for example: C:\Scripts\EvergreenAdmx)
* Run or schedule the script

I have scheduled the script to run daily:

`
EvergreenAdmx.ps1 -WindowsVersion "20H2" -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions"
`
The above execution will keep the central Policy Store up-to-date on a daily basis.

A sample .xml file that you can import in Task Scheduler is provided with this script.

# Admx files

Also see [Change Log](https://github.com/msfreaks/EvergreenAdmx/blob/main/CHANGELOG.md) for a list of supported products.

Now supports
* Adobe Acrobat Reader DC
* Citrix Workspace App
* FSLogix
* Google Chrome
* Microsoft Edge (Chromium)
* Microsoft Office
* Microsoft OneDrive
* Microsoft Windows 10 (1903/1909/2004/20H2)
* Mozilla Firefox
* Zoom Desktop Client

# Notes

I have not tested this script on Windows Core.
Some of the Admx files can only be obtained by installing the package that was downloaded. For instance, the Windows 10 Admx files are in an msi file, the OneDrive Admx files are in the installation folder after installing OneDrive.
So this is what the script does for these packages: installing the package, copying the Admx files, uninstalling the package.

This script processes all the products by default. Simply comment out any products you don't need and the script will skip those.


[github-release-badge]: https://img.shields.io/github/release/msfreaks/EvergreenAdmx.svg?style=flat-square
[github-release]: https://github.com/msfreaks/EvergreenAdmx/releases/latest
[license-badge]: https://img.shields.io/github/license/msfreaks/EvergreenAdmx.svg?style=flat-square
[license]: https://github.com/msfreaks/EvergreenAdmx/blob/master/LICENSE

