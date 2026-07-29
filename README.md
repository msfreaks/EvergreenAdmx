# 🌲 EvergreenAdmx

[![Release][github-release-badge]][github-release]
[![PowerShell Gallery][psgallery-downloads-badge]][poshgallery-evergreenadmx]
[![Codacy][code-quality-badge]][code-quality]
[![X][x-follow-badge]][x-follow]

After deploying several Azure Virtual Desktop environments I decided I no
longer wanted to manually download the Admx files I needed, and I wanted a
way to keep them up-to-date.

This script solves both problems.

- Automatically checks for newer versions of ADMX files and processes them
  when found
- Optionally copies the new ADMX files to your Policy Store or a custom
  location

Named as an homage to the [Evergreen module][evergreen-module] by Aaron Parker
[@stealthpuppy][stealthpuppy].

## 🚀 How to use

Quick start:

- Download the script to a location of your choosing
  (for example: `C:\Scripts\EvergreenAdmx`)
- Run or schedule the script

You can also install the script from the PowerShell Gallery
[EvergreenAdmx][poshgallery-evergreenadmx]:

```powershell
Install-Script -Name EvergreenAdmx
```

To register a weekly scheduled task (Sunday at 01:00, runs as SYSTEM), use
`-CreateScheduledTask`. Bound parameters are forwarded to the task; the
script exits after registration and does not download Admx files in that run.

### 💡 Examples

Run with defaults (Windows 11 25H2 plus Edge, OneDrive, 365 Apps,
Clipchamp, Notepad, Winget, and Windows Terminal) into the current folder:

```powershell
.\EvergreenAdmx.ps1
```

Download the defaults into a working folder:

```powershell
.\EvergreenAdmx.ps1 -WorkingDirectory "C:\Temp\EvergreenAdmx"
```

Use Windows Server 2025 defaults instead of Windows 11:

```powershell
.\EvergreenAdmx.ps1 -WindowsVersion 2025
```

Target a specific Windows 10 feature version (LTSC 2021 / `21H2`):

```powershell
.\EvergreenAdmx.ps1 -WindowsVersion 10 -WindowsFeatureVersion 21H2
```

Download only selected products, grouped into product folders:

```powershell
.\EvergreenAdmx.ps1 -WorkingDirectory "C:\Temp\EvergreenAdmx" `
  -Include @('Windows 11', 'Microsoft Edge', 'Microsoft FSLogix', 'Zoom') `
  -UseProductFolders
```

Copy the defaults into the domain Central Store (multiple languages,
product folders):

```powershell
.\EvergreenAdmx.ps1 `
  -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions" `
  -Languages @('en-US', 'nl-NL') `
  -UseProductFolders
```

Update the local machine PolicyDefinitions folder:

```powershell
.\EvergreenAdmx.ps1 -PolicyStore "C:\Windows\PolicyDefinitions"
```

Keep a locally installed OneDrive build when processing OneDrive ADMX:

```powershell
.\EvergreenAdmx.ps1 -Include @('Microsoft OneDrive') -PreferLocalOneDrive
```

Create or update a weekly SYSTEM scheduled task that refreshes the Central
Store (exits after registration):

```powershell
.\EvergreenAdmx.ps1 `
  -PolicyStore "C:\Windows\SYSVOL\domain\Policies\PolicyDefinitions" `
  -Languages @('en-US', 'nl-NL') `
  -UseProductFolders `
  -CreateScheduledTask
```

Download products to the Central Store and remove known obsolete Admx files:

```powershell
.\EvergreenAdmx.ps1 `
  -PolicyStore "\\contoso.com\SYSVOL\contoso.com\Policies\PolicyDefinitions" `
  -Include @('Windows 11', 'Microsoft Edge', 'Schannel') `
  -CleanPolicyStore
```

Preview obsolete-file cleanup without downloading or deleting:

```powershell
.\EvergreenAdmx.ps1 `
  -PolicyStore "\\contoso.com\SYSVOL\contoso.com\Policies\PolicyDefinitions" `
  -CleanPolicyStoreOnly `
  -WhatIf
```

A thin VDI-style orchestrator example (params + optional custom zips) lives in
[`samples/Update-PolicyDefinitions.ps1`](samples/Update-PolicyDefinitions.ps1).

## ⚙️ Parameters

Pass parameters with a leading `-` (PowerShell syntax), for example
`-WorkingDirectory` or `-Include`.

### 🪟 -WindowsVersion

Specifies Windows major version. Supports `10`, `11`, `2022`, or `2025`.
Default is `11`.

### 🔢 -WindowsFeatureVersion

Specifies the Windows 10 or Windows 11 feature version to get the Admx files
for.

- Windows 10: `21H2`, `22H2` (default `22H2`)
- Windows 11: `23H2`, `24H2`, `25H2` (default `25H2`)

Ignored when `-WindowsVersion` is `2022` or `2025`.

Note: Current Windows 11 ADMX templates (`23H2` / `24H2` / `25H2`) can also
manage Windows 10 clients; some settings apply only to newer OS versions.

Why Windows 10 `21H2` and `22H2` are still offered:

- **`22H2`** — last Windows 10 general availability (GA) release. Mainstream
  support ended October 14, 2025, but organizations on
  [Extended Security Updates (ESU)][windows-10-esu] still need matching ADMX
  templates.
- **`21H2`** — base for **Windows 10 Enterprise LTSC 2021**, which remains
  supported until January 2027 (longer for IoT Enterprise LTSC). Kept so LTSC
  2021 environments can pull the correct policy definitions.

Older Windows 10 feature versions (`1903`–`21H1`) and Windows 11 `21H2` /
`22H2` were removed because they are fully out of support.

### 📁 -WorkingDirectory

Specifies a working directory for the script.

- Admx files are stored in a subdirectory called `admx`
- Downloaded files are stored in a subdirectory called `downloads`

Defaults to the current script location.

### 🏪 -PolicyStore

Specifies a Policy Store location to copy the Admx files to after processing.
When this path does not exist, the script creates the Policy Store directory
and language subfolders from `-Languages`.

### 🌐 -Languages

Specifies an array of languages to process. Entries must be in `xy-XY` format
(also accepts short forms such as `es` and region tags such as `es-419`).
Defaults to `en-US`.

### 📂 -UseProductFolders

When set, Admx files are copied to their respective product folders under
`admx` in the WorkingDirectory.

### 📦 -CustomPolicyStore

Specifies a location for custom policy files (UNC or local folder).

- Finds `.admx` files in this location and at least one language folder with
  `.adml` file(s)
- Versioning is based on the newest file found recursively
  (any `.admx` or `.adml`)
- If any file has changed, the script processes all files found in the
  location

Use this for one-off templates (for example Microsoft Defender for Endpoint
zips you host yourself) that are not first-class `-Include` products.

### ✅ -Include

Specifies which Admx products to download and process. Use exact product
names from the list below.

When `-Include` is **not** specified, the script downloads this default set
(the Windows product matches `-WindowsVersion`):

| `-WindowsVersion` | Default products |
| --- | --- |
| `11` (default) | `Windows 11`, `Microsoft Edge`, `Microsoft OneDrive`, `Microsoft 365 Apps`, `Microsoft Clipchamp`, `Microsoft Notepad`, `Microsoft Winget`, `Windows Terminal` |
| `10` | `Windows 10`, `Microsoft Edge`, `Microsoft OneDrive`, `Microsoft 365 Apps`, `Microsoft Clipchamp`, `Microsoft Notepad`, `Microsoft Winget`, `Windows Terminal` |
| `2022` | `Windows 2022`, `Microsoft Edge`, `Microsoft OneDrive`, `Microsoft 365 Apps`, `Microsoft Clipchamp`, `Microsoft Notepad`, `Microsoft Winget`, `Windows Terminal` |
| `2025` | `Windows 2025`, `Microsoft Edge`, `Microsoft OneDrive`, `Microsoft 365 Apps`, `Microsoft Clipchamp`, `Microsoft Notepad`, `Microsoft Winget`, `Windows Terminal` |

Use the `-Include` parameter to specify exactly which products you want to download and process, replacing the default set.  
Below are all valid values you can use with `-Include`. Each links to the respective ADMX download or product policy documentation.

- [`1Password`][ref-1password]
- [`Adobe Acrobat`][ref-adobe-acrobat] (Continuous track)
- [`Adobe DC`][ref-adobe-dc] (community combined template)
- [`Adobe Reader`][ref-adobe-reader] (Continuous track)
- [`BIS-F`][ref-bisf] (Base Image Script Framework)
- [`Brave Browser`][ref-brave]
- [`Citrix Workspace App`][ref-citrix]
- [`Custom Policy Store`][ref-custom-policy-store]
  (local / UNC path you provide)
- [`Dell Command Update`][ref-dell-command-update]
  (latest Universal installer via winget; ADMX from `Templates`)
- [`Devolutions Remote Desktop Manager`][ref-devolutions]
- [`Dropbox`][ref-dropbox]
- [`Foxit PDF`][ref-foxit] (Reader + Editor)
- [`Google Chrome`][ref-chrome]
- [`HP Anyware`][ref-hp-anyware]
  (PCoIP ADMX from Standard Agent; requires 7-Zip)
- [`LibreOffice`][ref-libreoffice]
  (Collabora Office / LibreOffice GPO templates)
- [`Microsoft 365 Apps`][ref-365-apps]
- [`Microsoft AVD`][ref-avd]
- [`Microsoft Clipchamp`][ref-clipchamp]
- [`Microsoft Edge`][ref-edge] (Chromium)
- [`Microsoft FSLogix`][ref-fslogix]
- [`Microsoft Notepad`][ref-notepad]
- [`Microsoft OneDrive`][ref-onedrive]
  (local installation or Evergreen)
- [`Microsoft PowerToys`][ref-powertoys]
- [`Microsoft Visual Studio`][ref-visual-studio]
- [`Microsoft VS Code`][ref-vscode]
- [`Microsoft Winget`][ref-winget]
- [`Mozilla Firefox`][ref-firefox]
- [`Mozilla Thunderbird`][ref-thunderbird]
- [`PSAppDeployToolkit`][ref-psadt]
- [`Schannel`][ref-schannel] (Crosse Schannel Group Policy templates)
- [`Security ADMX`][ref-security-admx] (Custom template for Windows hardening)
- [`Slack`][ref-slack]
- [`TeamViewer`][ref-teamviewer]
- [`Windows 10`][ref-win10-22h2]
  ([`21H2`][ref-win10-21h2] / [`22H2`][ref-win10-22h2]; see
  -WindowsFeatureVersion above for why these remain)
- [`Windows 11`][ref-win11-25h2]
  ([`23H2`][ref-win11-23h2] / [`24H2`][ref-win11-24h2] /
  [`25H2`][ref-win11-25h2])
- [`Windows 2022`][ref-winserver-2022] (Windows Server 2022)
- [`Windows 2025`][ref-winserver-2025] (Windows Server 2025)
- [`Windows Terminal`][ref-windows-terminal]
- [`Winget-AutoUpdate`][ref-wau]
- [`Winget-AutoUpdate-Intune`][ref-wau-intune]
- [`Zoom`][ref-zoom]
- [`Zoom VDI`][ref-zoom-vdi]

### 💾 -PreferLocalOneDrive

Microsoft OneDrive Admx files are only available after installing OneDrive.
If this script is running on a machine that has OneDrive installed locally,
use this switch to prevent automatically uninstalling OneDrive.

### 🧹 -CleanPolicyStore

After processing, removes known obsolete or conflicting files from
`-PolicyStore`:

- `WinStoreUI.admx` / `.adml` (known Group Policy conflict)
- Legacy Office templates matching `*12*`–`*15*`
- Adobe Acrobat/Reader Classic 2017 and 2020 templates
- Citrix Profile Management `ctxprofile*` templates
- Non-`.admx`/`.adml` files at the store root and non-language extract folders

Requires `-PolicyStore`. Supports `-WhatIf`.

### 🧹 -CleanPolicyStoreOnly

Skips Admx downloads and only runs the Policy Store cleanup described above.
Requires `-PolicyStore`. Supports `-WhatIf`.

### 🗓️ -CreateScheduledTask

Creates or updates a Windows Scheduled Task named `EvergreenAdmx` that runs
this script weekly (Sunday at 01:00) as `SYSTEM` with highest privileges.
Compatible with Windows Server 2022 and 2025 via `Register-ScheduledTask`.

Other parameters bound on the same command line are forwarded to the task
action. The script exits after registering the task and does not download
Admx files in that run.

## ⚠️ Breaking changes

See the [Change Log][change-log] for the full history. Highlights in
**2607.0**:

- Removed Adobe Acrobat/Reader Classic 2017 and Classic 2020 tracks (EOL);
  Continuous track only
- Removed Microsoft Desktop Optimization Pack ADMX support
  (extended support ended April 14, 2026)
- Removed end-of-life Windows 10 feature versions `1903`–`21H1`;
  Windows 10 now supports `21H2` and `22H2` only
- Removed end-of-life Windows 11 feature versions `21H2` and `22H2`;
  Windows 11 now supports `23H2`, `24H2`, and `25H2` only

Earlier breaking changes (product renames, `-WindowsVersion` /
`-WindowsFeatureVersion`, and the `-Include` parameter) are documented in the
[Change Log][change-log].

## 📝 Notes

- This script has not been tested on Windows Core
- Some Admx files can only be obtained by installing the downloaded package
  (for example Windows 10/11 Admx MSI packages, and OneDrive after install)
- If you use the script to download Windows 10 or Windows 11 Admx files, remove
  any existing installs of those Admx MSI packages first, or the script will
  fail
- For those packages the script installs the package, copies the Admx files,
  then uninstalls the package
- Dell Command Update ADMX requires `winget`. Extraction uses 7-Zip when
  available; otherwise Dell's silent `/passthrough` extract is used
- HP Anyware ADMX requires `winget` (to resolve the current version) and
  7-Zip to extract `PCoIP.admx` from the Standard Agent installer

## 🙏 Credits

Thank you [Jonathan Pitre][jonathan-pitre] for keeping me sharp, providing
fixes and improvements!

Thank you [Dan Gough][dan-gough] for the `Get-Link`, `Get-Version`, and
`Resolve-Uri` helper functions from [Nevergreen][nevergreen].

[github-release-badge]: https://img.shields.io/github/v/release/msfreaks/EvergreenAdmx.svg?style=flat-square
[github-release]: https://github.com/msfreaks/EvergreenAdmx/releases/latest
[psgallery-downloads-badge]: https://img.shields.io/powershellgallery/dt/EvergreenAdmx.svg?style=flat-square
[code-quality-badge]: https://app.codacy.com/project/badge/Grade/c0efab02b66442399bb16b0493cdfbef?style=flat-square
[code-quality]: https://www.codacy.com/gh/msfreaks/EvergreenAdmx/dashboard?utm_source=github.com&utm_medium=referral&utm_content=msfreaks/EvergreenAdmx&utm_campaign=Badge_Grade
[x-follow-badge]: https://img.shields.io/twitter/follow/menschab?style=flat-square&logo=x
[x-follow]: https://x.com/menschab
[change-log]: https://github.com/msfreaks/EvergreenAdmx/blob/main/CHANGELOG.md
[poshgallery-evergreenadmx]: https://www.powershellgallery.com/packages/EvergreenAdmx/
[evergreen-module]: https://github.com/EUCPilots/evergreen-module
[stealthpuppy]: https://x.com/stealthpuppy
[jonathan-pitre]: https://github.com/JonathanPitre
[dan-gough]: https://github.com/DanGough
[nevergreen]: https://github.com/DanGough/Nevergreen
[windows-10-esu]: https://learn.microsoft.com/en-us/windows/whats-new/extended-security-updates
[ref-1password]: https://support.1password.com/mobile-device-management/?windows=
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
[ref-hp-anyware]: https://anyware.hp.com/components/standard-agent-for-windows/26.05/documentation/administrators-guide/reference/install-gpo-template-files
[ref-libreoffice]: https://github.com/CollaboraOnline/ADMX
[ref-chrome]: https://chromeenterprise.google/policies/
[ref-powertoys]: https://github.com/microsoft/PowerToys/releases
[ref-thunderbird]: https://github.com/thunderbird/policy-templates
[ref-windows-terminal]: https://github.com/microsoft/terminal/releases
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
