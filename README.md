<h1 align="center">
  <img src="assets/logo.png" alt="EvergreenAdmx" />
</h1>

[![Release][github-release-badge]][github-release]
[![PowerShell Gallery][psgallery-downloads-badge]][poshgallery-evergreenadmx]
[![Codacy][code-quality-badge]][code-quality]
[![X][x-follow-badge]][x-follow]

Automatically download and sync ADMX/ADML templates so your Group Policy store stays current.

After deploying several Azure Virtual Desktop environments I decided I no longer wanted to manually download the Admx files I needed, and I wanted a way to keep them up-to-date.

This script solves both problems:

- Automatically checks for newer versions of ADMX files and processes them when found
- Optionally copies the new ADMX files to your Policy Store or a custom location

Named as an homage to the [Evergreen module](https://github.com/EUCPilots/evergreen-module) by Aaron Parker ([@stealthpuppy](https://x.com/stealthpuppy)).

## Contents

- [Requirements](#requirements)
- [Install](#install)
- [Quick start](#quick-start)
- [Examples](#examples)
- [Parameters](#parameters)
- [Breaking changes](#breaking-changes)
- [Notes](#notes)
- [Roadmap](#roadmap)
- [Credits](#credits)
- [License](#license)

<a id="requirements"></a>

## 📋 Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Administrator rights when writing to a Central Policy Store or creating a scheduled task
- `winget` for Dell Command Update and HP Anyware
- 7-Zip for HP Anyware (and preferred for Dell Command Update extraction)

<a id="install"></a>

## 🚀 Install

### Download from GitHub

```powershell
md C:\Scripts\EvergreenAdmx -Force >$null
iwr https://raw.githubusercontent.com/msfreaks/EvergreenAdmx/main/EvergreenAdmx.ps1 -OutFile C:\Scripts\EvergreenAdmx\EvergreenAdmx.ps1
```

### Install from the Gallery

```powershell
Install-Script EvergreenAdmx -Force
```

With [PSResourceGet](https://www.powershellgallery.com/packages/Microsoft.PowerShell.PSResourceGet/) (PowerShell 7.4+):

```powershell
Install-PSResource EvergreenAdmx -Reinstall
```

<a id="quick-start"></a>

## ⚡ Quick start

Defaults (Windows 11 25H2 plus Edge, OneDrive, 365 Apps, Clipchamp, Notepad, Winget, and Windows Terminal) into the current folder:

```powershell
.\EvergreenAdmx.ps1
```

Working folder:

```powershell
.\EvergreenAdmx.ps1 -WorkingDirectory C:\Temp\EvergreenAdmx
```

Local PolicyDefinitions:

```powershell
.\EvergreenAdmx.ps1 -PolicyStore C:\Windows\PolicyDefinitions
```

<a id="examples"></a>

## 💡 Examples

Windows Server 2025 instead of Windows 11:

```powershell
.\EvergreenAdmx.ps1 -WindowsVersion 2025
```

Windows 10 LTSC 2021 (`21H2`):

```powershell
.\EvergreenAdmx.ps1 -WindowsVersion 10 -WindowsFeatureVersion 21H2
```

Selected products, grouped into product folders:

```powershell
.\EvergreenAdmx.ps1 -WorkingDirectory C:\Temp\EvergreenAdmx `
  -Include 'Windows 11','Microsoft Edge','Microsoft FSLogix','Zoom' `
  -UseProductFolders
```

Central Store (multiple languages, product folders):

```powershell
.\EvergreenAdmx.ps1 `
  -PolicyStore C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions `
  -Languages en-US,nl-NL `
  -UseProductFolders
```

Prefer a locally installed OneDrive build:

```powershell
.\EvergreenAdmx.ps1 -Include 'Microsoft OneDrive' -PreferLocalOneDrive
```

Weekly SYSTEM task for the Central Store (exits after registration):

```powershell
.\EvergreenAdmx.ps1 `
  -PolicyStore C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions `
  -Languages en-US,nl-NL `
  -UseProductFolders `
  -CreateScheduledTask
```

Central Store with cleanup of known obsolete Admx files:

```powershell
.\EvergreenAdmx.ps1 `
  -PolicyStore \\contoso.com\SYSVOL\contoso.com\Policies\PolicyDefinitions `
  -Include 'Windows 11','Microsoft Edge','Schannel' `
  -CleanPolicyStore
```

Preview obsolete-file cleanup only (no download or delete):

```powershell
.\EvergreenAdmx.ps1 `
  -PolicyStore \\contoso.com\SYSVOL\contoso.com\Policies\PolicyDefinitions `
  -CleanPolicyStoreOnly `
  -WhatIf
```

Thin orchestrator sample: [`samples/Update-PolicyDefinitions.ps1`](samples/Update-PolicyDefinitions.ps1).

<a id="parameters"></a>

## ⚙ Parameters

Pass parameters with a leading `-` (PowerShell syntax), for example `-WorkingDirectory` or `-Include`.

### -WindowsVersion

Specifies Windows major version. Supports `10`, `11`, `2022`, or `2025`. Default is `11`.

### -WindowsFeatureVersion

Specifies the Windows 10 or Windows 11 feature version to get the Admx files for.

- Windows 10: `21H2`, `22H2` (default `22H2`)
- Windows 11: `23H2`, `24H2`, `25H2` (default `25H2`)

Ignored when `-WindowsVersion` is `2022` or `2025`.

Current Windows 11 ADMX templates (`23H2` / `24H2` / `25H2`) can also manage Windows 10 clients; some settings apply only to newer OS versions. Windows 10 `21H2` / `22H2` remain available for LTSC and ESU — see [Notes](#notes).

### -WorkingDirectory

Specifies a working directory for the script.

- Admx files are stored in a subdirectory called `admx`
- Downloaded files are stored in a subdirectory called `downloads`

Defaults to the current script location.

### -PolicyStore

Specifies a Policy Store location to copy the Admx files to after processing. When this path does not exist, the script creates the Policy Store directory and language subfolders from `-Languages`.

### -Languages

Specifies an array of languages to process. Entries must be in `xy-XY` format (also accepts short forms such as `es` and region tags such as `es-419`). Defaults to `en-US`.

### -UseProductFolders

When set, Admx files are copied to their respective product folders under `admx` in the WorkingDirectory.

### -CustomPolicyStore

Specifies a location for custom policy files (UNC or local folder).

- Finds `.admx` files in this location and at least one language folder with `.adml` file(s)
- Versioning is based on the newest file found recursively (any `.admx` or `.adml`)
- If any file has changed, the script processes all files found in the location

Use this for one-off templates (for example Microsoft Defender for Endpoint zips you host yourself) that are not first-class `-Include` products.

### -Include

Specifies which Admx products to download and process. Use product names from the list below, or short aliases (for example `Edge` for `Microsoft Edge`, `BISF` for `BIS-F`). Tab completion is supported.

When `-Include` is **not** specified, the script downloads this default set (the Windows product matches `-WindowsVersion`):

| `-WindowsVersion` | Windows product | Also included by default |
| ----------------- | --------------- | ------------------------ |
| `11` (default) | `Windows 11` | Edge, OneDrive, 365 Apps, Clipchamp, Notepad, Winget, Windows Terminal |
| `10` | `Windows 10` | same shared set |
| `2022` | `Windows 2022` | same shared set |
| `2025` | `Windows 2025` | same shared set |

Shared defaults: `Microsoft Edge`, `Microsoft OneDrive`, `Microsoft 365 Apps`, `Microsoft Clipchamp`, `Microsoft Notepad`, `Microsoft Winget`, `Windows Terminal`.

`-Include` replaces the default set with exactly the products you specify.

<details>
<summary>Supported <code>-Include</code> products</summary>

- `1Password`
- `Adobe DC` (community combined template)
- `Adobe Acrobat` (Continuous track)
- `Adobe Reader` (Continuous track)
- `BIS-F` (Base Image Script Framework)
- `Brave Browser`
- `Citrix Workspace app`
- `Custom Policy Store` (local / UNC path you provide)
- `Dell Command Update` (latest Universal installer via winget)
- `Devolutions Remote Desktop Manager`
- `Dropbox`
- `Foxit PDF` (Reader + Editor)
- `Google Chrome`
- `HP Anyware` (PCoIP ADMX from Standard Agent; requires 7-Zip)
- `LibreOffice` (Collabora Office / LibreOffice GPO templates)
- `Microsoft 365 Apps`
- `Microsoft AVD`
- `Microsoft Clipchamp`
- `Microsoft Edge`
- `Microsoft FSLogix`
- `Microsoft Notepad`
- `Microsoft OneDrive` (local installation or download from Evergreen)
- `Microsoft PowerToys`
- `Microsoft Visual Studio`
- `Microsoft VS Code`
- `Microsoft Winget`
- `Mozilla Firefox`
- `Mozilla Thunderbird`
- `PSAppDeployToolkit`
- `Schannel` (Crosse Schannel GPO templates)
- `Security ADMX` (Custom template for Windows hardening)
- `Slack`
- `TeamViewer`
- `Windows 10` (`21H2` / `22H2`)
- `Windows 11` (`23H2` / `24H2` / `25H2`)
- `Windows 2022` (Windows Server 2022)
- `Windows 2025` (Windows Server 2025)
- `Windows Terminal`
- `Winget-AutoUpdate`
- `Winget-AutoUpdate-Intune`
- `Zoom`
- `Zoom VDI`

</details>

### -PreferLocalOneDrive

Microsoft OneDrive Admx files are only available after installing OneDrive. If this script is running on a machine that has OneDrive installed locally, use this switch to prevent automatically uninstalling OneDrive.

### -CleanPolicyStore

After processing, removes known obsolete or conflicting files from `-PolicyStore`:

- `WinStoreUI.admx` / `.adml` (known Group Policy conflict)
- Legacy Office templates matching `*12*`–`*15*`
- Adobe Acrobat/Reader Classic 2017 and 2020 templates
- Citrix Profile Management `ctxprofile*` templates
- `CitrixBase.admx` / `.adml` (no longer required; see [CTX696338](https://support.citrix.com/external/article/CTX696338/citrix-workspace-app-admx-download-does.html))
- Non-`.admx`/`.adml` files at the store root and non-language extract folders

Requires `-PolicyStore`. Supports `-WhatIf`.

### -CleanPolicyStoreOnly

Skips Admx downloads and only runs the Policy Store cleanup described above. Requires `-PolicyStore`. Supports `-WhatIf`.

### -CreateScheduledTask

Creates or updates a Windows Scheduled Task named `EvergreenAdmx` that runs this script weekly (Sunday at 01:00) as `SYSTEM` with highest privileges. Compatible with Windows Server 2022 and 2025 via `Register-ScheduledTask`.

Other parameters bound on the same command line are forwarded to the task action. The script exits after registering the task and does not download Admx files in that run. Change day/time later in Task Scheduler.

<a id="breaking-changes"></a>

## ⚠ Breaking changes

Highlights in **2607.0** (full history and earlier breaking changes in the [Changelog](CHANGELOG.md)):

- Removed Adobe Acrobat/Reader Classic 2017 and Classic 2020 tracks (EOL); Continuous track only
- Removed Microsoft Desktop Optimization Pack ADMX support (extended support ended April 14, 2026)
- Removed end-of-life Windows 10 feature versions `1903`–`21H1`; Windows 10 now supports `21H2` and `22H2` only
- Removed end-of-life Windows 11 feature versions `21H2` and `22H2`; Windows 11 now supports `23H2`, `24H2`, and `25H2` only

<a id="notes"></a>

## 📝 Notes

- This script has not been tested on Windows Core
- Some Admx files can only be obtained by installing the downloaded package (for example Windows 10/11 Admx MSI packages, and OneDrive after install)
- If you use the script to download Windows 10 or Windows 11 Admx files, remove any existing installs of those Admx MSI packages first, or the script will fail
- For those packages the script installs the package, copies the Admx files, then uninstalls the package
- Dell Command Update ADMX requires `winget`. Extraction uses 7-Zip when available; otherwise Dell's silent `/passthrough` extract is used
- HP Anyware ADMX requires `winget` (to resolve the current version) and 7-Zip to extract `PCoIP.admx` from the Standard Agent installer
- Windows 10 `22H2` is the last GA release (mainstream support ended October 14, 2025); organizations on [Extended Security Updates (ESU)](https://learn.microsoft.com/en-us/windows/whats-new/extended-security-updates) still need matching ADMX templates
- Windows 10 `21H2` is the base for **Windows 10 Enterprise LTSC 2021** (supported until January 2027, longer for IoT Enterprise LTSC)

<a id="roadmap"></a>

## 🗺️ Roadmap

- Add logging options
- Add notification options
- Detect user domain automatically
- Add ADMX versioning automatically (useful for Intune)

<a id="credits"></a>

## 🙏 Credits

Thank you [Jonathan Pitre](https://github.com/JonathanPitre) for keeping me sharp, providing fixes and improvements!

Thank you [Dan Gough](https://github.com/DanGough) for the `Get-Link`, `Get-Version`, and `Resolve-Uri` helper functions from [Nevergreen](https://github.com/DanGough/Nevergreen).

<a id="license"></a>

## 📄 License

This project is licensed under the [MIT License](LICENSE).

[github-release-badge]: https://img.shields.io/github/v/release/msfreaks/EvergreenAdmx.svg?style=flat-square
[github-release]: https://github.com/msfreaks/EvergreenAdmx/releases/latest
[psgallery-downloads-badge]: https://img.shields.io/powershellgallery/dt/EvergreenAdmx.svg?style=flat-square
[code-quality-badge]: https://app.codacy.com/project/badge/Grade/c0efab02b66442399bb16b0493cdfbef?style=flat-square
[code-quality]: https://www.codacy.com/gh/msfreaks/EvergreenAdmx/dashboard?utm_source=github.com&utm_medium=referral&utm_content=msfreaks/EvergreenAdmx&utm_campaign=Badge_Grade
[x-follow-badge]: https://img.shields.io/twitter/follow/menschab?style=flat-square&logo=x
[x-follow]: https://x.com/menschab
[poshgallery-evergreenadmx]: https://www.powershellgallery.com/packages/EvergreenAdmx/
