# Change Log

## To do

- Add logging options (yep, since the beginning)
- Add notification options (yep, also since the beginning)
- Detect user domain automatically (Get code from PSADT)
- Add ADMX versioning automatically, very useful for Intune

## 2607.1

- Added `-CleanPolicyStore` to remove known obsolete or conflicting Admx/Adml from `-PolicyStore` after processing (WinStoreUI, legacy Office `*12*`–`*15*`, Adobe Classic 2017/2020, `ctxprofile*`, non-policy junk files/folders)
- Added `-CleanPolicyStoreOnly` to run Policy Store cleanup without downloading Admx files
- Creates the Central Policy Store and language folders when `-PolicyStore` is set and the path is missing
- Added Schannel ADMX support via [Crosse/SchannelGroupPolicy](https://github.com/Crosse/SchannelGroupPolicy) (`-Include 'Schannel'`)
- Script now supports `-WhatIf` for Policy Store create/cleanup operations
- Added sample thin orchestrator at `samples/Update-PolicyDefinitions.ps1`
- Added Pester coverage for obsolete file patterns, `Initialize-PolicyStore`, and `Clear-ObsoleteAdmx` (including `-WhatIf`)
- Added `-CreateScheduledTask` to create or update a weekly SYSTEM scheduled task (`EvergreenAdmx`, Sunday at 01:00) via `Register-ScheduledTask`; bound parameters are forwarded and the script exits after registration
- Extracted `New-EvergreenAdmxTaskArgumentList` for scheduled-task argument building (testable without registering a task)
- Added Pester unit / Integration / Nightly suites under `tests/` and GitHub Actions workflows (`ci.yml`, `release-smoke.yml`, `nightly.yml`)
- Added markdownlint job and expanded PSScriptAnalyzer coverage (main script + `tests/`) in `ci.yml`, with shared `PSScriptAnalyzerSettings.psd1`
- Removed `EvergreenAdmx.xml` sample Task Scheduler export in favor of `-CreateScheduledTask`
- Replaced `Invoke-WebRequest -OutFile` downloads with HttpClient-based `Invoke-FileDownload` (TLS 1.3 preferred, TLS 1.2 fallback) for faster, streamed downloads on PowerShell 5.1 [#39](https://github.com/msfreaks/EvergreenAdmx/issues/39)
- Added Microsoft PowerToys ADMX [#75](https://github.com/msfreaks/EvergreenAdmx/issues/75)
- Added Windows Terminal ADMX [#76](https://github.com/msfreaks/EvergreenAdmx/issues/76)
- Added Mozilla Thunderbird ADMX [#77](https://github.com/msfreaks/EvergreenAdmx/issues/77)
- Added Dropbox ADMX [#79](https://github.com/msfreaks/EvergreenAdmx/issues/79)
- Added Foxit PDF (Reader + Editor) ADMX [#80](https://github.com/msfreaks/EvergreenAdmx/issues/80)
- Added LibreOffice / Collabora Office ADMX [#81](https://github.com/msfreaks/EvergreenAdmx/issues/81)
- Added HP Anyware (PCoIP) ADMX via Standard Agent extract (requires winget + 7-Zip) [#83](https://github.com/msfreaks/EvergreenAdmx/issues/83)
- Tracked Omnissa Horizon GPO Bundle (Customer Connect login wall) [#78](https://github.com/msfreaks/EvergreenAdmx/issues/78)
- Tracked Cisco Secure Client (no official ADMX published) [#82](https://github.com/msfreaks/EvergreenAdmx/issues/82)
- Tracked Lenovo Commercial Vantage / Dock Manager ADMX (no stable public URL yet) [#84](https://github.com/msfreaks/EvergreenAdmx/issues/84)
- Added Windows 11 **25H2** ADMX support (download id `108542`, V3.0); default Windows 11 feature version is now `25H2` [#70](https://github.com/msfreaks/EvergreenAdmx/issues/70)
- Fixed first-run crash when `AdmxVersions.xml` is missing by initializing `$AdmxVersions = @{}` [#70](https://github.com/msfreaks/EvergreenAdmx/issues/70)
- Fixed OneDrive install detection `if`/`elseif` chain [#71](https://github.com/msfreaks/EvergreenAdmx/issues/71) / [#72](https://github.com/msfreaks/EvergreenAdmx/pull/72)
- Relaxed language format validation to allow `es` and `es-419` [#52](https://github.com/msfreaks/EvergreenAdmx/issues/52)
- Added Microsoft Notepad ADMX [#61](https://github.com/msfreaks/EvergreenAdmx/issues/61)
- Added Microsoft Clipchamp ADMX [#60](https://github.com/msfreaks/EvergreenAdmx/issues/60)
- Added Microsoft Visual Studio ADMX [#32](https://github.com/msfreaks/EvergreenAdmx/issues/32)
- Added Microsoft VS Code ADMX via zip extract (no install) [#33](https://github.com/msfreaks/EvergreenAdmx/issues/33)
- Added Slack ADMX [#57](https://github.com/msfreaks/EvergreenAdmx/issues/57)
- Added TeamViewer ADMX [#69](https://github.com/msfreaks/EvergreenAdmx/issues/69)
- Added community Adobe DC ADMX (separate from official Adobe Acrobat) [#68](https://github.com/msfreaks/EvergreenAdmx/issues/68)
- Added Security-ADMX [#62](https://github.com/msfreaks/EvergreenAdmx/issues/62)
- Added Dell Command Update ADMX [#50](https://github.com/msfreaks/EvergreenAdmx/issues/50)
- Added Winget-AutoUpdate ADMX [#49](https://github.com/msfreaks/EvergreenAdmx/issues/49)
- Added Winget-AutoUpdate-Intune ADMX [#35](https://github.com/msfreaks/EvergreenAdmx/issues/35)
- Added PSAppDeployToolkit ADMX [#53](https://github.com/msfreaks/EvergreenAdmx/issues/53)
- Added Devolutions Remote Desktop Manager ADMX via Bin zip extract (no install) [#34](https://github.com/msfreaks/EvergreenAdmx/issues/34)
- Added 1Password ADMX [#74](https://github.com/msfreaks/EvergreenAdmx/issues/74)
- Expanded default `-Include` set to also download Microsoft Clipchamp, Microsoft Notepad, Microsoft Winget, and Windows Terminal
- Changed Dell Command Update ADMX source from the stale GitHub repo to the latest `Dell.CommandUpdate.Universal` installer via `winget download`, extracting ADMX from the EXE `Templates` folder (7-Zip preferred; Dell `/passthrough` extract fallback)
- Removed Adobe Acrobat/Reader Classic 2017 and Classic 2020 tracks (EOL); Continuous track only (`breaking change`)
- Removed Microsoft Desktop Optimization Pack ADMX support (extended support ended April 14, 2026) (`breaking change`)
- Removed end-of-life Windows 10 feature versions `1903`–`21H1`; Windows 10 now supports `21H2` and `22H2` only (`breaking change`)
- Removed end-of-life Windows 11 feature versions `21H2` and `22H2`; Windows 11 now supports `23H2`, `24H2`, and `25H2` only (`breaking change`)

## 2503.1

- Added script parameter **WindowsVersion** that supports value **10**, **11**, **2022** and **2025** (`breaking change`)
- Replaced script parameter **Windows10Version** and **Windows11Version** by **WindowsFeatureVersion** (`breaking change`)
- Renamed product **Microsoft Office** to **Microsoft 365 Apps** (`breaking change`)
- Renamed product **Azure Virtual Desktop** to **Microsoft AVD** (`breaking change`)
- Renamed product **FSLogix** to **Microsoft FSlogix** (`breaking change`)
- Renamed product **Zoom Desktop Client** to **Zoom** (`breaking change`)
- Improved script parameters validation
- Improved verbose logging
- Fixed issues with PreferLocalOneDrive parameter
- Added admx for Windows Server 2025
- Added admx for Windows Server 2022
- Added admx for Zoom VDI
- Renamed functions to get url and latest version of policy definitions files from Get-$ProductAdmxOnline to Get-EvergreenAdmx%Product%
- Renamed functions to download policy definitions files from Get-$ProductAdmx to Invoke-EvergreenAdmx%Product%
- Added new Update-AdmxVersion function
- Added admx for Brave Browser [#48](https://github.com/msfreaks/EvergreenAdmx/issues/48) Thanks [Tom Plant](https://github.com/pl4nty)!
- Fixed scrapping for OneDrive [#55](https://github.com/msfreaks/EvergreenAdmx/pull/55) Thanks [Tom Plant](https://github.com/pl4nty)!
- Fix Windows 11, OneDrive and 365 apps admx downloads [#51](https://github.com/msfreaks/EvergreenAdmx/issues/51) Thanks [Tom Plant](https://github.com/pl4nty)!
- Fixed errors reported by PSScriptAnalyzer rules
- Improved code formatting based on powershell and markdown best practices

## 2411.1

- Added admx for Microsoft Winget [#40](https://github.com/msfreaks/EvergreenAdmx/issues/40)

## 2411.0

- Changed Get-WindowsAdmxOnline version return
- Added Admx for Microsoft Windows 11 (24H2) [#45](https://github.com/msfreaks/EvergreenAdmx/issues/45)

## 2402.1

- Fixed Get-WindowsAdmxOnline version return
- Improved function Get-WindowsAdmxOnline, added default parameters
- Improved function Get-WindowsAdmxDownloadId, added default parameters
- Improved Get-WindowsAdmx speed by switching to MSI extraction [#41](https://github.com/msfreaks/EvergreenAdmx/issues/41)
- Improved Adobe admx downloads with https URLs [#37](https://github.com/msfreaks/EvergreenAdmx/issues/37)
- Added admx for Microsoft Windows 11 (23H2) [#38](https://github.com/msfreaks/EvergreenAdmx/issues/38)
- Added back WindowsVersion parameter as an alias for Windows11Version
- Added admx for Azure Virtual Desktop [#17](https://github.com/msfreaks/EvergreenAdmx/issues/17)
- Added new functions [Get-Link](https://github.com/DanGough/Nevergreen/blob/main/Nevergreen/Private/Get-Link.ps1), [Get-Version](https://github.com/DanGough/Nevergreen/blob/main/Nevergreen/Private/Get-Version.ps1) and [Resolve-Uri](https://github.com/DanGough/Nevergreen/blob/main/Nevergreen/Private/Resolve-Uri.ps1). Thanks [Dan Gough](https://github.com/DanGough)!
- Added new function [Invoke-Download](https://github.com/DanGough/PsDownload/blob/main/PsDownload/Public/Invoke-Download.ps1) to improve download speed and get last modified date. Thanks [Dan Gough](https://github.com/DanGough)!
- Replaced Get-RedirectUrl function by Resolve-Uri Thanks [Dan Gough](https://github.com/DanGough)!
- Fixed and improved Zoom Desktop Client version and url detection. Now works on PowerShell 5.1 as well! Thanks [Dan Gough](https://github.com/DanGough)!
- Fixed Zoom Desktop Client admx copy to policy definitions
- Fixed and improved Get-MicrosoftOfficeAdmxOnline version detection
- Fixed and improved Get-MDOPAdmxOnline version detection
- Improved Get-FSLogixOnline
- Fixed and improved Get-OneDriveOnline, now use [EvergreenApi](https://stealthpuppy.com/evergreen/invoke) method for version and url detection
- Microsoft OneDrive now install silently
- Fixed typo
- Updated help

## 2301.2

- Fixed typo in New-Item command (thanks for pointing it out, riebest!)

## 2301.1

- Added Admx for Microsoft Windows 10 (22H2)
- Added Admx for Adobe Reader and Adobe Acrobat
- Fixed Microsoft Edge (Chromium) Admx download
- Fixed Zoom Desktop Client Admx downloads (hardcoded version)
- Replaced mkdir command by native posh one
- Cleanup code

## 2209.1

- Fixed bug for Microsoft OneDrive, thanks Jonathan!
- Fixed bug for Citrix Workspace App
- Added admx for Microsoft Windows 11 (22H2)
- Added cleanup logic for Google Chrome ADMX version checking, thanks for noticing Jonathan!
- Cleanup on code, thanks Jonathan!

## 2207.1

- Fixed bug where Citrix Workspace App would fail, thanks to [Jonathan Pitre](https://github.com/JonathanPitre)!

## 2206.1

- Added requirement for elevation to prevent running unelevated
- Added parameter -UseDefaultCredentials to all Invoke-WebRequest commands, except for Adobe since that is ftp :(
- Fixed bug where Amd64 OneDrive would mess everything up
- Removed support for Zoom Desktop Client (403 error when checking for new version)

## 2112.2

- Fixed bug where Windows 10 would throw a WindowsVersion variable error

## 2112.1

`Breaking change introduced!` (See [README][read-me] for more info)

- Added parameter 'Windows10Version' (`breaking change`)
- Added parameter 'Windows11Version' (`breaking change`)
- Removed parameter 'WindowsVersion' (`breaking change`)
- Added Admx for Microsoft Windows 10 (21H2)
- Added Admx for Microsoft Windows 11 (21H2)
- Fixed bug where copying the ADMX files for Citrix Workspace App would fail
- Fixed bug where terminating a running OneDrive process would fail if a process was not found

## 2111.1

- Internal version

## 2109.2

- Fixed bug where uninstall information for OneDrive would throw an error

## 2109.1

- Fixed bug where downloading the ADMX files for Citrix Workspace App would fail
- Fixed bug where downloading the ADMX files for BIS-F would fail to copy the .adml file
- Fixed bug where downloading the ADMX files for Microsoft OneDrive would fail

## 2107.1

- Fixed bug where Get-FSLogixOnline would fail using code provided by severud (thanks!)
- Typo corrected for Windows 20 (21H1)

## 2106.2

- Fixed bug where script was unable to get the version for Citrix Workspace App Admx

## 2106.1

- Added Admx for Microsoft Windows 10 (21H1)

## 2101.2

`Breaking change introduced!` (See [README][read-me] for more info)

- Added parameter 'CustomPolicyLocation'
- Added 'CustomPolicyLocation' logic
- Added parameter 'Include' (`breaking change`)
- Added 'Include' logic
- Added parameter 'PreferLocalOneDrive'
- Added 'PreferLocalOneDrive' logic

## 2101.1

- Internal version

## 2012.6

- Added parameter 'UseProductFolders'
- Added parameter 'Languages'
- Added 'UseProductFolders' logic
- Added 'Languages' logic
- Fixed bug where extracting of .cab files would fail
- Updated verbose output

## 2012.5

- Added Admx for Base Image Script Framework (BIS-F)
- Added Admx for Microsoft Desktop Optimization Pack (disabled by default)
- Updated cleanup logic

## 2012.4

- Added Admx for Adobe AcrobatReader DC
- Added Admx for Citrix Workspace App
- Added Admx for Mozilla Firefox
- Added Admx for Zoom Desktop Client

## 2012.3

- Added Admx for Google Chrome
- Fixed bug where hash for json output would throw an error. Switched to Xml output.

## 2012.2

- Added .xml file for Scheduled Task creation

## 2012.1

- Added Admx for Microsoft Windows 10 (1903/1909/2004/20H2)
- Added Admx for Microsoft Edge (Chromium)
- Added Admx for Microsoft OneDrive
- Added Admx for Microsoft Office
- Added Admx for FSLogix

[read-me]: https://github.com/msfreaks/EvergreenAdmx/blob/main/README.md
