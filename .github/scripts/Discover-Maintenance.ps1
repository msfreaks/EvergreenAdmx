#Requires -Version 7.4
[CmdletBinding()]
param([switch]$DryRun)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Maintenance.ps1')
$root = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$artifactRoot = Join-Path $env:RUNNER_TEMP 'maintenance-proposals'
$changed = [System.Collections.Generic.List[string]]::new()
$failures = [System.Collections.Generic.List[string]]::new()
$dependencyCache = @{}
foreach ($kind in @('routine', 'major', 'families')) {
    try {
        & (Join-Path $PSScriptRoot 'Update-Maintenance.ps1') -Kind $kind -DryRun:$DryRun -DependencyCache $dependencyCache
        if (-not $DryRun) {
            $resultPath = Join-Path $root 'maintenance-result.json'
            $result = Get-Content $resultPath -Raw | ConvertFrom-Json
            if ($result.changes.Count) {
                $folder = Join-Path $artifactRoot $kind
                foreach ($file in @(Get-MaintenanceFileList $kind $result.files) + 'maintenance-result.json') {
                    $destination = Join-Path $folder $file
                    $null = New-Item (Split-Path $destination) -ItemType Directory -Force
                    Copy-Item -LiteralPath (Join-Path $root $file) -Destination $destination -Force
                }
                $changed.Add($kind)
            }
        }
    } catch {
        $failures.Add("${kind}: $_")
        Write-Warning "${kind}: $_"
    } finally {
        # Each category starts from the same baseline in this disposable Actions checkout.
        git -C $root restore --worktree -- .
        if ($LASTEXITCODE -ne 0) { throw 'Cannot restore the discovery baseline.' }
    }
}
$matrix = @{ kind = if ($changed.Count) { @($changed) } else { @('none') } } | ConvertTo-Json -Compress
"changed=$($changed.Count -gt 0)".ToLowerInvariant() | Out-File $env:GITHUB_OUTPUT -Append -Encoding utf8
"matrix=$matrix" | Out-File $env:GITHUB_OUTPUT -Append -Encoding utf8
if ($failures.Count) { throw ($failures -join '; ') }
