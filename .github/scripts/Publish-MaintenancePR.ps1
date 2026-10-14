#Requires -Version 7.4
[CmdletBinding()]
param([Parameter(Mandatory)][string]$ArtifactPath)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Maintenance.ps1')

function Invoke-GitHubCLI {
    param([string[]]$Arguments)
    $output = & gh @Arguments
    if ($LASTEXITCODE -ne 0) { throw "GitHub CLI failed: $($Arguments[0])" }
    return $output
}

$result = Get-Content (Join-Path $ArtifactPath 'maintenance-result.json') -Raw | ConvertFrom-Json
if (-not $result.changes.Count) { Write-Output 'No updates available.'; return }
$allowed = @(Get-MaintenanceFileList -Kind $result.kind -Files $result.files)
$repository = $env:GITHUB_REPOSITORY
$defaultBranch = Invoke-GitHubCLI @('api', "repos/$repository", '--jq', '.default_branch')
if ($env:GITHUB_REF -ne "refs/heads/$defaultBranch") { throw 'Maintenance PRs may only be published from the default branch.' }
$branch = "automation/$($result.kind)"
$pulls = @(Invoke-GitHubCLI @('pr', 'list', '--repo', $repository, '--head', $branch, '--base', $defaultBranch, '--json', 'number,author,headRefOid') | ConvertFrom-Json)
$lease = ''
if ($pulls.Count) {
    $pull = $pulls[0]
    if ($pull.author.login -ne 'app/github-actions') { throw 'Maintenance branch PR is not owned by GitHub Actions.' }
    $authors = @(Invoke-GitHubCLI @('api', '--paginate', "repos/$repository/pulls/$($pull.number)/commits", '--jq', '.[].author.login'))
    if (@($authors | Where-Object { $_ -ne 'github-actions[bot]' }).Count) { throw 'Maintenance branch contains human commits; leave it for review.' }
    $lease = $pull.headRefOid
} else {
    # Refuse to overwrite an orphaned branch or another author's branch.
    $refs = @(Invoke-GitHubCLI @('api', "repos/$repository/git/matching-refs/heads/$branch", '--jq', '.[].ref'))
    if ($refs -contains "refs/heads/$branch") { throw 'Maintenance branch exists without an open bot PR.' }
}
foreach ($file in $allowed) {
    Copy-Item -LiteralPath (Join-Path $ArtifactPath $file) -Destination $file -Force
}
git checkout -b $branch
if ($LASTEXITCODE -ne 0) { throw 'Cannot create maintenance branch.' }
git add -- $allowed
if ($LASTEXITCODE -ne 0) { throw 'Cannot stage maintenance changes.' }
$null = Invoke-GitHubCLI @('auth', 'setup-git')
if ($lease) {
    git fetch --no-tags --depth=2 origin $branch
    if ($LASTEXITCODE -ne 0) { throw 'Cannot inspect the existing maintenance branch.' }
    $existingHead = git rev-parse FETCH_HEAD
    if ($LASTEXITCODE -ne 0 -or $existingHead -ne $lease) { throw 'Maintenance branch changed concurrently; refresh refused.' }
    if (Test-MaintenanceProposalUnchanged $existingHead) {
        Write-Output 'Maintenance PR and baseline are unchanged; keeping its commit and CI results.'
        return
    }
}
git config user.name 'github-actions[bot]'
git config user.email '41898282+github-actions[bot]@users.noreply.github.com'
$title = switch ($result.kind) {
    'routine' { 'chore(deps): update PowerShell dependencies' }
    'major' { 'chore(deps): review PowerShell major updates' }
    'families' { 'feat: review newly discovered product families' }
}
git commit -m $title
if ($LASTEXITCODE -ne 0) { throw 'Cannot commit maintenance changes.' }
git push "--force-with-lease=refs/heads/${branch}:$lease" origin "HEAD:refs/heads/$branch"
if ($LASTEXITCODE -ne 0) { throw 'Maintenance branch changed concurrently; push refused.' }
$body = @"
$title

$($result.changes | ForEach-Object { "- $_" } | Out-String)
CI runs lint and unit tests under Windows PowerShell 5.1 and PowerShell 7 against the exact proposed commit.
$(if ($result.kind -eq 'routine') { 'Patch/minor dependencies merge automatically after required checks pass.' } else { 'Human review is required before merging this update.' })
"@
$bodyPath = Join-Path $env:RUNNER_TEMP 'maintenance-pr.md'
[IO.File]::WriteAllText($bodyPath, $body)
if ($pulls.Count) {
    $null = Invoke-GitHubCLI @('pr', 'edit', [string]$pull.number, '--repo', $repository, '--title', $title, '--body-file', $bodyPath)
    $number = $pull.number
} else {
    $null = Invoke-GitHubCLI @('pr', 'create', '--repo', $repository, '--base', $defaultBranch, '--head', $branch, '--title', $title, '--body-file', $bodyPath)
    $number = Invoke-GitHubCLI @('pr', 'view', $branch, '--repo', $repository, '--json', 'number', '--jq', '.number')
}
$sha = git rev-parse HEAD
$null = Invoke-GitHubCLI @('workflow', 'run', 'ci.yml', '--repo', $repository, '--ref', $branch, '-f', "expected_sha=$sha")
if ($result.kind -eq 'routine') {
    $null = Invoke-GitHubCLI @('pr', 'merge', [string]$number, '--repo', $repository, '--auto', '--squash')
}
