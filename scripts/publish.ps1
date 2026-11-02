param(
    [string]$Version,
    [Parameter(Mandatory = $true)][string]$Message,
    [string]$Python = 'python',
    [ValidateSet('git', 'api')][string]$Transport = 'git',
    [switch]$DocumentationOnly
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$projectRoot = Split-Path -Parent $PSScriptRoot
Push-Location -LiteralPath $projectRoot
try {
    function Invoke-Git {
        & git @args
        if ($LASTEXITCODE -ne 0) { throw 'Git command failed; publication stopped' }
    }
    $branch = & git branch --show-current
    if ($LASTEXITCODE -ne 0 -or $branch -ne 'main') { throw 'Publish from the main branch' }
    $remote = & git remote get-url origin
    if ($LASTEXITCODE -ne 0 -or $remote -notmatch 'github\.com[:/]') { throw 'GitHub origin is required' }
    if ($Transport -eq 'api') {
        & $Python scripts/push_api.py --check
        if ($LASTEXITCODE -ne 0) { throw 'Integrate remote main before publishing' }
    } else {
        Invoke-Git fetch origin main
        & git merge-base --is-ancestor origin/main HEAD
        if ($LASTEXITCODE -ne 0) { throw 'Integrate remote main before publishing' }
    }
    if (-not $DocumentationOnly) {
        if ($Version -notmatch '^\d+\.\d+\.\d+$') { throw 'Version must be X.Y.Z' }
        & $Python scripts/release.py check --tag "v$Version"
        if ($LASTEXITCODE -ne 0) { throw 'Version check failed' }
        if (-not (Test-Path -LiteralPath "docs/releases/v$Version.md")) { throw 'Release notes are required' }
        & git show-ref --verify --quiet "refs/tags/v$Version"
        if ($LASTEXITCODE -eq 0) { throw 'Local release tag already exists' }
        if ($Transport -eq 'git') {
            & git ls-remote --exit-code --tags origin "refs/tags/v$Version" | Out-Null
            if ($LASTEXITCODE -eq 0) { throw 'Remote release tag already exists' }
            if ($LASTEXITCODE -ne 2) { throw 'Could not check remote tags' }
        }
        & $Python verify.py "release-$Version"
        if ($LASTEXITCODE -ne 0) { throw 'Verification failed; nothing was pushed' }
    }
    Invoke-Git diff --check
    Invoke-Git add --all
    Invoke-Git diff --cached --check
    & git diff --cached --quiet
    if ($LASTEXITCODE -eq 1) { Invoke-Git commit -m $Message }
    elseif ($LASTEXITCODE -ne 0) { throw 'Cannot inspect staged changes' }
    if ($DocumentationOnly) {
        if ($Transport -eq 'api') {
            & $Python scripts/push_api.py
            if ($LASTEXITCODE -ne 0) { throw 'GitHub API publication failed' }
        } else { Invoke-Git push origin main }
    } else {
        Invoke-Git tag -a "v$Version" -m "Release v$Version"
        if ($Transport -eq 'api') {
            & $Python scripts/push_api.py --tag "v$Version"
            if ($LASTEXITCODE -ne 0) { throw 'GitHub API publication failed' }
        } else { Invoke-Git push --atomic origin main "refs/tags/v$Version" }
    }
    Write-Output 'Push completed. Check GitHub Actions for verification and release status.'
} finally {
    Pop-Location
}
