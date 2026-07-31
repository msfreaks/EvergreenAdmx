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

## 📚 Contents

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

## ⚙️ Parameters

Pass parameters with a leading `-` (PowerShell syntax), for example `-WorkingDirectory` or `-Include`.

### -CleanPolicyStore

After processing, removes known obsolete or conflicting files from `-PolicyStore` and repairs missing ADMLs:

- `WinStoreUI.admx` / `.adml` (replaced by `WindowsStore.admx`; namespace conflict if both present)
- `Microsoft-Windows-Geolocation-WLPAdm.admx` / `.adml` (replaced by `LocationProviderAdm.admx`)
- Legacy Office templates matching `*12*`–`*15*`
- Adobe Acrobat/Reader Classic 2017 and 2020 templates
- Citrix Profile Management `ctxprofile*` templates
- `CitrixBase.admx` / `.adml` (no longer required; see [CTX696338](https://support.citrix.com/external/article/CTX696338/citrix-workspace-app-admx-download-does.html))
- Non-`.admx`/`.adml` files at the store root and non-language extract folders
- Missing language `.adml` files: copies from `en-US` into requested `-Languages` folders when available; warns if an `.admx` has no ADML at all

Requires `-PolicyStore`. Supports `-WhatIf`.

### -CleanPolicyStoreOnly

Skips Admx downloads and only runs the Policy Store cleanup described above. Requires `-PolicyStore`. Supports `-WhatIf`.

### -CreateScheduledTask

Creates or updates a Windows Scheduled Task named `EvergreenAdmx` that runs this script weekly (Sunday at 01:00) as `SYSTEM` with highest privileges. Compatible with Windows Server 2022 and 2025 via `Register-ScheduledTask`.

Other parameters bound on the same command line are forwarded to the task action. The script exits after registering the task and does not download Admx files in that run. Change day/time later in Task Scheduler.

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

- [`1Password`][ref-1password]
- [`ABBYY FineReader PDF`][ref-abbyy] (FineReader 16)
- [`Admin By Request`][ref-admin-by-request]
- [`Adobe Acrobat`][ref-adobe-acrobat] (Continuous track)
- [`Adobe DC`][ref-adobe-dc] (community combined template)
- [`Adobe Reader`][ref-adobe-reader] (Continuous track)
- [`BIS-F`][ref-bisf] (Base Image Script Framework)
- [`Brave Browser`][ref-brave]
- [`Citrix Workspace app`][ref-citrix]
- [`Custom Policy Store`][ref-custom-policy-store] (local / UNC path you provide)
- [`Dell Command Update`][ref-dell-command-update] (latest Universal installer via winget)
- [`Devolutions Remote Desktop Manager`][ref-devolutions]
- [`Dropbox`][ref-dropbox]
- [`Foxit PDF`][ref-foxit] (Reader + Editor)
- [`Google Chrome`][ref-chrome]
- [`GoTo`][ref-goto] (GoTo app / GoTo Connect)
- [`HP Anyware`][ref-hp-anyware] (PCoIP ADMX from Standard Agent; requires 7-Zip)
- [`Lenovo Dock Manager`][ref-lenovo-dock] (policy_setup.exe Group Policy templates)
- [`LibreOffice`][ref-libreoffice] (Collabora Office / LibreOffice GPO templates)
- [`Microsoft 365 Apps`][ref-365-apps]
- [`Microsoft AVD`][ref-avd]
- [`Microsoft Clipchamp`][ref-clipchamp]
- [`Microsoft Edge`][ref-edge]
- [`Microsoft FSLogix`][ref-fslogix]
- [`Microsoft Notepad`][ref-notepad]
- [`Microsoft OneDrive`][ref-onedrive] (local installation or download from Evergreen)
- [`Microsoft PowerToys`][ref-powertoys]
- [`Microsoft Visual Studio`][ref-visual-studio]
- [`Microsoft VS Code`][ref-vscode]
- [`Microsoft Winget`][ref-winget]
- [`Mozilla Firefox`][ref-firefox]
- [`Mozilla Thunderbird`][ref-thunderbird]
- [`PDF-XChange`][ref-pdf-xchange] (Editor, Tools, Driver, Updater, Vault)
- [`PSAppDeployToolkit`][ref-psadt]
- [`RealVNC Connect`][ref-realvnc] (Server + Viewer)
- [`Schannel`][ref-schannel] (Crosse Schannel GPO templates)
- [`Security ADMX`][ref-security-admx] (Custom template for Windows hardening)
- [`Slack`][ref-slack]
- [`Specops Authentication Client`][ref-specops] (on-prem + Entra ID)
- [`TeamViewer`][ref-teamviewer]
- [`Windows 10`][ref-win10-22h2] ([`21H2`][ref-win10-21h2] / [`22H2`][ref-win10-22h2])
- [`Windows 11`][ref-win11-25h2] ([`23H2`][ref-win11-23h2] / [`24H2`][ref-win11-24h2] / [`25H2`][ref-win11-25h2])
- [`Windows 2022`][ref-winserver-2022] (Windows Server 2022)
- [`Windows 2025`][ref-winserver-2025] (Windows Server 2025)
- [`Windows Terminal`][ref-windows-terminal]
- [`Winget-AutoUpdate`][ref-wau]
- [`Winget-AutoUpdate-Intune`][ref-wau-intune]
- [`WSL`][ref-wsl] (Windows Subsystem for Linux Intune ADMX)
- [`Zoom`][ref-zoom]
- [`Zoom VDI`][ref-zoom-vdi]

</details>

### -Languages

Specifies an array of languages to process. Entries must be in `xy-XY` format (also accepts short forms such as `es` and region tags such as `es-419`). Defaults to `en-US`.

### -PolicyStore

Specifies a Policy Store location to copy the Admx files to after processing. When this path does not exist, the script creates the Policy Store directory and language subfolders from `-Languages`.

### -PreferLocalOneDrive

Microsoft OneDrive Admx files are only available after installing OneDrive. If this script is running on a machine that has OneDrive installed locally, use this switch to prevent automatically uninstalling OneDrive.

### -StampAdmxRevision

When set, stamps ADMX/ADML `revision` attributes from the product release Version (GitHub release, download page, etc.) normalized to the ADMX `versionString` **Major.Minor** form (for example `143.0.3624.0` becomes `143.0`). This helps Intune Imported Administrative Templates show a meaningful Version column.

Only attributes currently set to `1.0` are updated:

- ADMX `policyDefinitions/@revision`
- ADMX `resources/@minRequiredRevision` (when also `1.0`)
- ADML `policyDefinitionResources/@revision`

Higher vendor revisions are left unchanged. Default is off so Central Policy Store copies keep vendor XML unless you opt in.

### -UseProductFolders

When set, Admx files are copied to their respective product folders under `admx` in the WorkingDirectory.

### -WindowsFeatureVersion

Specifies the Windows 10 or Windows 11 feature version to get the Admx files for.

- Windows 10: `21H2`, `22H2` (default `22H2`)
- Windows 11: `23H2`, `24H2`, `25H2` (default `25H2`)

Ignored when `-WindowsVersion` is `2022` or `2025`.

Current Windows 11 ADMX templates (`23H2` / `24H2` / `25H2`) can also manage Windows 10 clients; some settings apply only to newer OS versions. Windows 10 `21H2` / `22H2` remain available for LTSC and ESU — see [Notes](#notes).

### -WindowsVersion

Specifies Windows major version. Supports `10`, `11`, `2022`, or `2025`. Default is `11`.

### -WorkingDirectory

Specifies a working directory for the script.

- Admx files are stored in a subdirectory called `admx`
- Downloaded files are stored in a subdirectory called `downloads`

Defaults to the current script location.

<a id="breaking-changes"></a>

## ⚠️ Breaking changes

Highlights in **2607.0** (full history and earlier breaking changes in the [Changelog](CHANGELOG.md)):

- Removed Adobe Acrobat/Reader Classic 2017 and Classic 2020 tracks (EOL); Continuous track only
- Removed Microsoft Desktop Optimization Pack ADMX support (extended support ended April 14, 2026)
- Removed end-of-life Windows 10 feature versions `1903`–`21H1`; Windows 10 now supports `21H2` and `22H2` only
- Removed end-of-life Windows 11 feature versions `21H2` and `22H2`; Windows 11 now supports `23H2`, `24H2`, and `25H2` only

<a id="notes"></a>

## 📝 Notes

- This script has not been tested on Windows Core
- Use `-StampAdmxRevision` when preparing templates for Intune so the Version column reflects the product release (Major.Minor); this does not bypass Intune namespace conflicts on re-upload
- Some Admx files can only be obtained by installing the downloaded package (for example Windows 10/11 Admx MSI packages, and OneDrive after install)
- If you use the script to download Windows 10 or Windows 11 Admx files, remove any existing installs of those Admx MSI packages first, or the script will fail
- For those packages the script installs the package, copies the Admx files, then uninstalls the package
- Dell Command Update ADMX requires `winget`. Extraction uses 7-Zip when available; otherwise Dell's silent `/passthrough` extract is used
- HP Anyware ADMX requires `winget` (to resolve the current version) and 7-Zip to extract `PCoIP.admx` from the Standard Agent installer
- Windows 10 `22H2` is the last GA release (mainstream support ended October 14, 2025); organizations on [Extended Security Updates (ESU)](https://learn.microsoft.com/en-us/windows/whats-new/extended-security-updates) still need matching ADMX templates
- Windows 10 `21H2` is the base for **Windows 10 Enterprise LTSC 2021** (supported until January 2027, longer for IoT Enterprise LTSC)
- Use `-CleanPolicyStore` (or `-CleanPolicyStoreOnly`) to remove known conflicting ADMX files (for example `WinStoreUI` / `Microsoft-Windows-Geolocation-WLPAdm`) and copy missing language ADMLs from `en-US` when available
- Some Microsoft ADMX upgrades change GPO registry value types or paths (for example ErrorReporting `DefaultConsent` “unexpected type”, or SkyDrive → OneDrive registry keys). Those require rebuilding the affected GPO settings; the script does not rewrite `registry.pol`. See [Known issues for managing Group Policy clients](https://learn.microsoft.com/en-us/troubleshoot/windows-client/active-directory/known-issues-for-group-policy-clients)

<a id="roadmap"></a>

## 🗺️ Roadmap

- Add logging options
- Add notification options
- Detect user domain automatically

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
[ref-1password]: https://support.1password.com/mobile-device-management/?windows=
[ref-abbyy]: https://help.abbyy.com/en-us/finereader/16/admin_guide/gpo_domain/
[ref-admin-by-request]: https://www.adminbyrequest.com/ADMX
[ref-adobe-acrobat]: https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/misc/AcrobatADMTemplate.zip
[ref-adobe-dc]: https://github.com/systmworks/Adobe-DC-ADMX
[ref-adobe-reader]: https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/misc/ReaderADMTemplate.zip
[ref-bisf]: https://github.com/EUCweb/BIS-F
[ref-brave]: https://github.com/brave/brave-browser/wiki/Deploying-Brave-with-Group-Policy
[ref-citrix]: https://docs.citrix.com/en-us/citrix-workspace-app-for-windows/group-policy
[ref-custom-policy-store]: https://learn.microsoft.com/en-us/troubleshoot/windows-client/group-policy/create-and-manage-central-store
[ref-dell-command-update]: https://www.dell.com/support/kbdoc/en-us/000293701/how-do-i-access-amdx-and-adml-files-for-use-with-dell-command-update
[ref-devolutions]: https://docs.devolutions.net/rdm/knowledge-base/how-to-articles/apply-policies-gpos
[ref-dropbox]: https://github.com/dropbox/GPO-Templates
[ref-foxit]: https://kb.foxit.com/s/articles/360040241112-Available-GPO-templates
[ref-goto]: https://goto-desktop.goto.com/GoToAppAdministrativeTemplates.zip
[ref-hp-anyware]: https://anyware.hp.com/components/standard-agent-for-windows/26.05/documentation/administrators-guide/reference/install-gpo-template-files
[ref-lenovo-dock]: https://download.lenovo.com/consumer/options/policy_setup.exe
[ref-libreoffice]: https://github.com/CollaboraOnline/ADMX
[ref-chrome]: https://chromeenterprise.google/policies/
[ref-pdf-xchange]: https://www.pdf-xchange.com/Tracker%5FAD%5FAdministrativeTemplates.zip
[ref-powertoys]: https://github.com/microsoft/PowerToys/releases
[ref-realvnc]: https://downloads.realvnc.com/download/file/policy.files/RealVNC-server-admx-templates-Latest.zip
[ref-specops]: https://download.specopssoft.com/Release/Client/Specops.Client.AdmxTemplates.zip
[ref-thunderbird]: https://github.com/thunderbird/policy-templates
[ref-windows-terminal]: https://github.com/microsoft/terminal/releases
[ref-wsl]: https://github.com/microsoft/WSL/tree/master/intune
[ref-365-apps]: https://www.microsoft.com/en-us/download/details.aspx?id=49030
[ref-avd]: https://aka.ms/avdgpo
[ref-clipchamp]: https://www.microsoft.com/en-us/download/details.aspx?id=105674
[ref-edge]: https://learn.microsoft.com/en-us/deployedge/configure-microsoft-edge
[ref-fslogix]: https://learn.microsoft.com/en-us/fslogix/how-to-use-group-policy-templates
[ref-notepad]: https://download.microsoft.com/download/72ea16a9-4cc9-4032-945d-3a56a483d034/WindowsNotepadAdminTemplates.cab
[ref-onedrive]: https://learn.microsoft.com/en-us/sharepoint/use-group-policy
[ref-visual-studio]: https://www.microsoft.com/en-us/download/details.aspx?id=104405
[ref-vscode]: https://code.visualstudio.com/docs/setup/enterprise
[ref-win10-21h2]: https://www.microsoft.com/en-us/download/details.aspx?id=104042
[ref-win10-22h2]: https://www.microsoft.com/en-us/download/details.aspx?id=104677
[ref-win11-23h2]: https://www.microsoft.com/en-us/download/details.aspx?id=105667
[ref-win11-24h2]: https://www.microsoft.com/en-us/download/details.aspx?id=106254
[ref-win11-25h2]: https://www.microsoft.com/en-us/download/details.aspx?id=108542
[ref-winserver-2022]: https://www.microsoft.com/en-us/download/details.aspx?id=104003
[ref-winserver-2025]: https://www.microsoft.com/en-us/download/details.aspx?id=106295
[ref-winget]: https://github.com/microsoft/winget-cli/releases
[ref-firefox]: https://github.com/mozilla/policy-templates
[ref-psadt]: https://github.com/PSAppDeployToolkit/PSAppDeployToolkit
[ref-schannel]: https://github.com/Crosse/SchannelGroupPolicy
[ref-security-admx]: https://github.com/Harvester57/Security-ADMX
[ref-slack]: https://slack.com/help/articles/11906214948755-Manage-desktop-app-configurations
[ref-teamviewer]: https://github.com/systmworks/TeamViewer-ADMX
[ref-wau]: https://github.com/Romanitho/Winget-AutoUpdate
[ref-wau-intune]: https://github.com/Weatherlights/Winget-AutoUpdate-Intune
[ref-zoom]: https://support.zoom.com/hc/en/article?id=zm_kb&sysparm_article=KB0065466
[ref-zoom-vdi]: https://support.zoom.com/hc/en/article?id=zm_kb&sysparm_article=KB0064784
