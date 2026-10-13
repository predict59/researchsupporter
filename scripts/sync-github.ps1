param(
  [Parameter(Position = 0)]
  [string]$Message = "Update price research supporter"
)

$ErrorActionPreference = "Stop"

function Invoke-Git {
  param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)

  & git @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "git $($Arguments -join ' ') failed with exit code $LASTEXITCODE"
  }
}

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $projectRoot

$branch = (& git branch --show-current).Trim()
if ($LASTEXITCODE -ne 0 -or $branch -ne "main") {
  throw "GitHub sync must be run from the main branch. Current branch: $branch"
}

Write-Host "[1/5] Verifying production build..."
& npm run build
if ($LASTEXITCODE -ne 0) {
  throw "Build failed. Nothing was committed or pushed."
}

Write-Host "[2/5] Fetching origin/main..."
Invoke-Git fetch origin main

$counts = (& git rev-list --left-right --count HEAD...origin/main).Trim() -split "\s+"
if ($LASTEXITCODE -ne 0 -or $counts.Count -ne 2) {
  throw "Could not compare local main with origin/main."
}

$ahead = [int]$counts[0]
$behind = [int]$counts[1]
if ($behind -gt 0) {
  Write-Host "[3/5] Rebasing local work onto the latest origin/main..."
  Invoke-Git pull --rebase --autostash origin main
} else {
  Write-Host "[3/5] Local main is current with origin/main."
}

Write-Host "[4/5] Creating a commit when files changed..."
Invoke-Git add --all
& git diff --cached --quiet
$hasChanges = $LASTEXITCODE -ne 0

if ($hasChanges) {
  Invoke-Git commit -m $Message
} else {
  Write-Host "No uncommitted changes found."
}

Write-Host "[5/5] Pushing main to GitHub..."
Invoke-Git push origin main

Write-Host "Sync complete. GitHub Pages deployment will run from the pushed main branch."
