# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project uses a `YYMM.patch` release versioning scheme.

## [2608.0] 2026-08-01

### Added

- Added Specops Authentication Client ADMX (on-prem + Entra ID templates)
- Added WSL ADMX from [microsoft/WSL intune](https://github.com/microsoft/WSL/tree/master/intune) folder
- Added Lenovo Dock Manager ADMX via policy_setup.exe (partially addresses #84; Commercial Vantage still tracked)
- Added PDF-XChange ADMX
- Added RealVNC Connect ADMX (Server + Viewer)
- Added ABBYY FineReader PDF ADMX
- Added Admin By Request ADMX
- Added GoTo ADMX
- Added `-StampAdmxRevision` to stamp ADMX/ADML `revision` (and ADMX `minRequiredRevision` when `1.0`) from the product release Version as Major.Minor, so Intune shows a meaningful template version; files already above `1.0` are left unchanged
- `-CleanPolicyStore` / `-CleanPolicyStoreOnly` now also remove `Microsoft-Windows-Geolocation-WLPAdm.admx` / `.adml` (superseded by `LocationProviderAdm`) and copy missing language ADMLs from `en-US` when available

## [2607.1] - 2026-07-29

### Added

- Added GitHub Actions workflow to publish `EvergreenAdmx.ps1` to the PowerShell Gallery on release (`publish-psgallery.yml`)
- Added tab-completion for `-Include` names and aliases via ArgumentCompleter
- Added project logo (`assets/logo.png`) and show it in README.md

### Changed

- `-Include` accepts short aliases (ProductKey / compact forms / former names), e.g. `BISF` for `BIS-F`, `Edge` for `Microsoft Edge`
- Replaced brittle `-Include` ValidateSet with catalog-based resolution and clearer errors that list only current products (removed products such as MDOP get a dedicated message)
- `-CleanPolicyStore` / `-CleanPolicyStoreOnly` now remove obsolete `CitrixBase.admx` / `.adml` ([CTX696338](https://support.citrix.com/external/article/CTX696338/citrix-workspace-app-admx-download-does.html))
- Nightly full-matrix always includes BIS-F (removed `EVERGREENADMX_INCLUDE_BISF` opt-in; previously excluded because GitHub zipballs often returned HTTP 403)
- Improved structure and clarity of README.md
- Enhanced formatting and organization of CHANGELOG.md
- Updated `tests/README.md` for the current unit / Integration / Nightly layout

### Fixed

- `Invoke-FileDownload` now sends a default `User-Agent` (`EvergreenAdmx`) so GitHub API / zipball downloads no longer fail with HTTP 403 (notably BIS-F and other GitHub-hosted ADMX sources)
- Aligned scheduled-task and Policy Store helper names with the unit suite: `New-EvergreenAdmxTaskArgumentList` and `Get-EvergreenAdmxObsoleteFilePattern`

## [2607.0] - 2026-07-29

### Added

- Added `-CleanPolicyStore` to remove known obsolete or conflicting Admx/Adml from `-PolicyStore` after processing (WinStoreUI, legacy Office `*12*`–`*15*`, Adobe Classic 2017/2020, `ctxprofile*`, non-policy junk files/folders)
- Added `-CleanPolicyStoreOnly` to run Policy Store cleanup without downloading Admx files
- Added Schannel ADMX support via [Crosse/SchannelGroupPolicy](https://github.com/Crosse/SchannelGroupPolicy) (`-Include 'Schannel'`)
- Added Pester coverage for obsolete file patterns, `Initialize-PolicyStore`, and `Clear-ObsoleteAdmx` (including `-WhatIf`)
- Added `-CreateScheduledTask` to create or update a weekly SYSTEM scheduled task (`EvergreenAdmx`, Sunday at 01:00) via `Register-ScheduledTask`; bound parameters are forwarded and the script exits after registration
- Added Pester unit / Integration / Nightly suites under `tests/` and GitHub Actions workflows (`ci.yml`, `release-smoke.yml`, `nightly.yml`)
- Added markdownlint job and expanded PSScriptAnalyzer coverage (main script + `tests/`) in `ci.yml`, with shared `PSScriptAnalyzerSettings.psd1`
- Added Microsoft PowerToys ADMX [#75](https://github.com/msfreaks/EvergreenAdmx/issues/75)
- Added Windows Terminal ADMX [#76](https://github.com/msfreaks/EvergreenAdmx/issues/76)
- Added Mozilla Thunderbird ADMX [#77](https://github.com/msfreaks/EvergreenAdmx/issues/77)
- Added Dropbox ADMX [#79](https://github.com/msfreaks/EvergreenAdmx/issues/79)
- Added Foxit PDF (Reader + Editor) ADMX [#80](https://github.com/msfreaks/EvergreenAdmx/issues/80)
- Added LibreOffice / Collabora Office ADMX [#81](https://github.com/msfreaks/EvergreenAdmx/issues/81)
- Added HP Anyware (PCoIP) ADMX via Standard Agent extract (requires winget + 7-Zip) [#83](https://github.com/msfreaks/EvergreenAdmx/issues/83)
- Added Windows 11 **25H2** ADMX support (download id `108542`, V3.0); default Windows 11 feature version is now `25H2` [#70](https://github.com/msfreaks/EvergreenAdmx/issues/70)
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

### Changed

- Creates the Central Policy Store and language folders when `-PolicyStore` is set and the path is missing
- Script now supports `-WhatIf` for Policy Store create/cleanup operations
- Extracted `New-EvergreenAdmxTaskArgumentList` for scheduled-task argument building (testable without registering a task)
- Replaced `Invoke-WebRequest -OutFile` downloads with HttpClient-based `Invoke-FileDownload` (TLS 1.3 preferred, TLS 1.2 fallback) for faster, streamed downloads on PowerShell 5.1 [#39](https://github.com/msfreaks/EvergreenAdmx/issues/39)
- Tracked Omnissa Horizon GPO Bundle (Customer Connect login wall) [#78](https://github.com/msfreaks/EvergreenAdmx/issues/78)
- Tracked Cisco Secure Client (no official ADMX published) [#82](https://github.com/msfreaks/EvergreenAdmx/issues/82)
- Tracked Lenovo Commercial Vantage / Dock Manager ADMX (no stable public URL yet) [#84](https://github.com/msfreaks/EvergreenAdmx/issues/84)
- Relaxed language format validation to allow `es` and `es-419` [#52](https://github.com/msfreaks/EvergreenAdmx/issues/52)
- Expanded default `-Include` set to also download Microsoft Clipchamp, Microsoft Notepad, Microsoft Winget, and Windows Terminal
- Changed Dell Command Update ADMX source from the stale GitHub repo to the latest `Dell.CommandUpdate.Universal` installer via `winget download`, extracting ADMX from the EXE `Templates` folder (7-Zip preferred; Dell `/passthrough` extract fallback)

### Removed

- Removed `EvergreenAdmx.xml` sample Task Scheduler export in favor of `-CreateScheduledTask`
- Removed Adobe Acrobat/Reader Classic 2017 and Classic 2020 tracks (EOL); Continuous track only
- Removed Microsoft Desktop Optimization Pack ADMX support (extended support ended April 14, 2026)
- Removed end-of-life Windows 10 feature versions `1903`–`21H1`; Windows 10 now supports `21H2` and `22H2` only
- Removed end-of-life Windows 11 feature versions `21H2` and `22H2`; Windows 11 now supports `23H2`, `24H2`, and `25H2` only

### Fixed

- Fixed first-run crash when `AdmxVersions.xml` is missing by initializing `$AdmxVersions = @{}` [#70](https://github.com/msfreaks/EvergreenAdmx/issues/70)
- Fixed OneDrive install detection `if`/`elseif` chain [#71](https://github.com/msfreaks/EvergreenAdmx/issues/71) / [#72](https://github.com/msfreaks/EvergreenAdmx/pull/72)

## [2503.1] - 2025-03-11

### Added

- Added script parameter **WindowsVersion** that supports value **10**, **11**, **2022** and **2025**
- Added admx for Windows Server 2025
- Added admx for Windows Server 2022
- Added admx for Zoom VDI
- Added new Update-AdmxVersion function
- Added admx for Brave Browser [#48](https://github.com/msfreaks/EvergreenAdmx/issues/48) Thanks [Tom Plant](https://github.com/pl4nty)!

### Changed

- Replaced script parameter **Windows10Version** and **Windows11Version** by **WindowsFeatureVersion**
- Renamed product **Microsoft Office** to **Microsoft 365 Apps**
- Renamed product **Azure Virtual Desktop** to **Microsoft AVD**
- Renamed product **FSLogix** to **Microsoft FSlogix**
- Renamed product **Zoom Desktop Client** to **Zoom**
- Improved script parameters validation
- Improved verbose logging
- Renamed functions to get url and latest version of policy definitions files from Get-$ProductAdmxOnline to Get-EvergreenAdmx%Product%
- Renamed functions to download policy definitions files from Get-$ProductAdmx to Invoke-EvergreenAdmx%Product%
- Improved code formatting based on powershell and markdown best practices

### Fixed

- Fixed issues with PreferLocalOneDrive parameter
- Fixed scrapping for OneDrive [#55](https://github.com/msfreaks/EvergreenAdmx/pull/55) Thanks [Tom Plant](https://github.com/pl4nty)!
- Fix Windows 11, OneDrive and 365 apps admx downloads [#51](https://github.com/msfreaks/EvergreenAdmx/issues/51) Thanks [Tom Plant](https://github.com/pl4nty)!
- Fixed errors reported by PSScriptAnalyzer rules

## [2411.1] - 2024-11-15

### Added

- Added admx for Microsoft Winget [#40](https://github.com/msfreaks/EvergreenAdmx/issues/40)

## [2411.0] - 2024-11-15

### Added

- Added Admx for Microsoft Windows 11 (24H2) [#45](https://github.com/msfreaks/EvergreenAdmx/issues/45)

### Changed

- Changed Get-WindowsAdmxOnline version return

## [2402.1] - 2024-02-27

### Added

- Added admx for Microsoft Windows 11 (23H2) [#38](https://github.com/msfreaks/EvergreenAdmx/issues/38)
- Added back WindowsVersion parameter as an alias for Windows11Version
- Added admx for Azure Virtual Desktop [#17](https://github.com/msfreaks/EvergreenAdmx/issues/17)
- Added new functions [Get-Link](https://github.com/DanGough/Nevergreen/blob/main/Nevergreen/Private/Get-Link.ps1), [Get-Version](https://github.com/DanGough/Nevergreen/blob/main/Nevergreen/Private/Get-Version.ps1) and [Resolve-Uri](https://github.com/DanGough/Nevergreen/blob/main/Nevergreen/Private/Resolve-Uri.ps1). Thanks [Dan Gough](https://github.com/DanGough)!
- Added new function [Invoke-Download](https://github.com/DanGough/PsDownload/blob/main/PsDownload/Public/Invoke-Download.ps1) to improve download speed and get last modified date. Thanks [Dan Gough](https://github.com/DanGough)!

### Changed

- Improved function Get-WindowsAdmxOnline, added default parameters
- Improved function Get-WindowsAdmxDownloadId, added default parameters
- Improved Get-WindowsAdmx speed by switching to MSI extraction [#41](https://github.com/msfreaks/EvergreenAdmx/issues/41)
- Improved Adobe admx downloads with https URLs [#37](https://github.com/msfreaks/EvergreenAdmx/issues/37)
- Replaced Get-RedirectUrl function by Resolve-Uri Thanks [Dan Gough](https://github.com/DanGough)!
- Improved Get-FSLogixOnline
- Microsoft OneDrive now install silently
- Updated help

### Fixed

- Fixed Get-WindowsAdmxOnline version return
- Fixed and improved Zoom Desktop Client version and url detection. Now works on PowerShell 5.1 as well! Thanks [Dan Gough](https://github.com/DanGough)!
- Fixed Zoom Desktop Client admx copy to policy definitions
- Fixed and improved Get-MicrosoftOfficeAdmxOnline version detection
- Fixed and improved Get-MDOPAdmxOnline version detection
- Fixed and improved Get-OneDriveOnline, now use [EvergreenApi](https://stealthpuppy.com/evergreen/invoke) method for version and url detection
- Fixed typo

## [2301.2] - 2023-01-11

### Fixed

- Fixed typo in New-Item command (thanks for pointing it out, riebest!)

## [2301.1]

### Added

- Added Admx for Microsoft Windows 10 (22H2)
- Added Admx for Adobe Reader and Adobe Acrobat

### Changed

- Replaced mkdir command by native posh one
- Cleanup code

### Fixed

- Fixed Microsoft Edge (Chromium) Admx download
- Fixed Zoom Desktop Client Admx downloads (hardcoded version)

## [2209.1]

### Added

- Added admx for Microsoft Windows 11 (22H2)
- Added cleanup logic for Google Chrome ADMX version checking, thanks for noticing Jonathan!

### Changed

- Cleanup on code, thanks Jonathan!

### Fixed

- Fixed bug for Microsoft OneDrive, thanks Jonathan!
- Fixed bug for Citrix Workspace App

## [2207.1]

### Fixed

- Fixed bug where Citrix Workspace App would fail, thanks to [Jonathan Pitre](https://github.com/JonathanPitre)!

## [2206.1]

### Added

- Added requirement for elevation to prevent running unelevated
- Added parameter -UseDefaultCredentials to all Invoke-WebRequest commands, except for Adobe since that is ftp :(

### Removed

- Removed support for Zoom Desktop Client (403 error when checking for new version)

### Fixed

- Fixed bug where Amd64 OneDrive would mess everything up

## [2112.2]

### Fixed

- Fixed bug where Windows 10 would throw a WindowsVersion variable error

## [2112.1]

### Added

- Added parameter 'Windows10Version'
- Added parameter 'Windows11Version'
- Added Admx for Microsoft Windows 10 (21H2)
- Added Admx for Microsoft Windows 11 (21H2)

### Removed

- Removed parameter 'WindowsVersion'

### Fixed

- Fixed bug where copying the ADMX files for Citrix Workspace App would fail
- Fixed bug where terminating a running OneDrive process would fail if a process was not found

## [2111.1]

### Changed

- Internal version

## [2109.2]

### Fixed

- Fixed bug where uninstall information for OneDrive would throw an error

## [2109.1]

### Fixed

- Fixed bug where downloading the ADMX files for Citrix Workspace App would fail
- Fixed bug where downloading the ADMX files for BIS-F would fail to copy the .adml file
- Fixed bug where downloading the ADMX files for Microsoft OneDrive would fail

## [2107.1]

### Changed

- Typo corrected for Windows 20 (21H1)

### Fixed

- Fixed bug where Get-FSLogixOnline would fail using code provided by severud (thanks!)

## [2106.2]

### Fixed

- Fixed bug where script was unable to get the version for Citrix Workspace App Admx

## [2106.1]

### Added

- Added Admx for Microsoft Windows 10 (21H1)

## [2101.2]

### Added

- Added parameter 'CustomPolicyLocation'
- Added 'CustomPolicyLocation' logic
- Added parameter 'Include'
- Added 'Include' logic
- Added parameter 'PreferLocalOneDrive'
- Added 'PreferLocalOneDrive' logic

## [2101.1]

### Changed

- Internal version

## [2012.6]

### Added

- Added parameter 'UseProductFolders'
- Added parameter 'Languages'
- Added 'UseProductFolders' logic
- Added 'Languages' logic

### Changed

- Updated verbose output

### Fixed

- Fixed bug where extracting of .cab files would fail

## [2012.5]

### Added

- Added Admx for Base Image Script Framework (BIS-F)
- Added Admx for Microsoft Desktop Optimization Pack (disabled by default)

### Changed

- Updated cleanup logic

## [2012.4]

### Added

- Added Admx for Adobe AcrobatReader DC
- Added Admx for Citrix Workspace App
- Added Admx for Mozilla Firefox
- Added Admx for Zoom Desktop Client

## [2012.3]

### Added

- Added Admx for Google Chrome

### Fixed

- Fixed bug where hash for json output would throw an error. Switched to Xml output.

## [2012.2]

### Added

- Added .xml file for Scheduled Task creation

## [2012.1]

### Added

- Added Admx for Microsoft Windows 10 (1903/1909/2004/20H2)
- Added Admx for Microsoft Edge (Chromium)
- Added Admx for Microsoft OneDrive
- Added Admx for Microsoft Office
- Added Admx for FSLogix

[2607.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2607.0...2607.1
[2607.0]: https://github.com/msfreaks/EvergreenAdmx/compare/2503.1...2607.0
[2503.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2411.1...2503.1
[2411.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2411.0...2411.1
[2411.0]: https://github.com/msfreaks/EvergreenAdmx/compare/2402.1...2411.0
[2402.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2301.2...2402.1
[2301.2]: https://github.com/msfreaks/EvergreenAdmx/compare/2301.1...2301.2
[2301.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2209.1...2301.1
[2209.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2207.1...2209.1
[2207.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2206.1...2207.1
[2206.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2112.2...2206.1
[2112.2]: https://github.com/msfreaks/EvergreenAdmx/compare/2112.1...2112.2
[2112.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2111.1...2112.1
[2111.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2109.2...2111.1
[2109.2]: https://github.com/msfreaks/EvergreenAdmx/compare/2109.1...2109.2
[2109.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2107.1...2109.1
[2107.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2106.2...2107.1
[2106.2]: https://github.com/msfreaks/EvergreenAdmx/compare/2106.1...2106.2
[2106.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2101.2...2106.1
[2101.2]: https://github.com/msfreaks/EvergreenAdmx/compare/2101.1...2101.2
[2101.1]: https://github.com/msfreaks/EvergreenAdmx/compare/2012.6...2101.1
[2012.6]: https://github.com/msfreaks/EvergreenAdmx/compare/2012.5...2012.6
[2012.5]: https://github.com/msfreaks/EvergreenAdmx/compare/2012.4...2012.5
[2012.4]: https://github.com/msfreaks/EvergreenAdmx/compare/2012.3...2012.4
[2012.3]: https://github.com/msfreaks/EvergreenAdmx/compare/2012.2...2012.3
[2012.2]: https://github.com/msfreaks/EvergreenAdmx/compare/2012.1...2012.2
[2012.1]: https://github.com/msfreaks/EvergreenAdmx/releases/tag/2012.1
