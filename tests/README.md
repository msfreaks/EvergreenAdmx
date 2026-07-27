# EvergreenAdmx tests

Pester 5 suites for EvergreenAdmx.

| Suite | File | Tag | When |
| --- | --- | --- | --- |
| Unit | `EvergreenAdmx.Tests.ps1` | _(none)_ | Every PR / push (`ci.yml`) |
| Integration | `EvergreenAdmx.Integration.Tests.ps1` | `Integration` | Release published / manual (`release-smoke.yml`) |
| Nightly | `EvergreenAdmx.Nightly.Tests.ps1` | `Nightly` | Weekly schedule / manual (`nightly.yml`) |

## Prerequisites

- Windows PowerShell 5.1+ or PowerShell 7+
- [Pester](https://pester.dev/) 5.5+
- Integration / Nightly require an elevated session (`#Requires -RunAsAdministrator`)

```powershell
Install-Module Pester -MinimumVersion 5.5.0 -Scope CurrentUser -Force -SkipPublisherCheck
```

## Local runs

Unit only (no admin, no downloads):

```powershell
$config = New-PesterConfiguration
$config.Run.Path = '.\tests\EvergreenAdmx.Tests.ps1'
$config.Output.Verbosity = 'Detailed'
Invoke-Pester -Configuration $config
```

Integration smoke (elevated):

```powershell
$config = New-PesterConfiguration
$config.Run.Path = '.\tests\EvergreenAdmx.Integration.Tests.ps1'
$config.Filter.Tag = @('Integration')
Invoke-Pester -Configuration $config
```

Full matrix (elevated, long-running):

```powershell
$config = New-PesterConfiguration
$config.Run.Path = '.\tests\EvergreenAdmx.Nightly.Tests.ps1'
$config.Filter.Tag = @('Nightly')
Invoke-Pester -Configuration $config
```

Set `EVERGREENADMX_INCLUDE_BISF=1` to include BIS-F in the nightly matrix (often 403 without GitHub auth).
