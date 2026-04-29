param(
  [Parameter(Mandatory = $true)]
  [string]$CommitMessage
)

$ErrorActionPreference = "Stop"

Write-Host "Checking git status..."
git status --short

git add .
$hasChanges = -not [string]::IsNullOrWhiteSpace((git status --porcelain))
if ($hasChanges) {
  Write-Host "Creating git commit..."
  git commit -m $CommitMessage

  Write-Host "Pushing code to GitHub..."
  git push
} else {
  Write-Host "No git changes detected. Skipping git commit/push."
}

Write-Host "Pushing code to Apps Script..."
clasp push -f

Write-Host "Creating Apps Script version..."
$versionOutput = clasp version $CommitMessage
$versionLine = $versionOutput | Select-String -Pattern "Created version (\d+)"
if (-not $versionLine) {
  throw "Could not determine new Apps Script version number."
}
$versionNumber = [int]$versionLine.Matches[0].Groups[1].Value

Write-Host "Updating live Apps Script deployment(s)..."
$deploymentsOutput = clasp deployments
$liveDeploymentIds = @()
foreach ($line in $deploymentsOutput) {
  if ($line -match "^\s*-\s+([A-Za-z0-9_-]+)\s+@([^\s]+)") {
    $deploymentId = $matches[1]
    $deploymentTarget = $matches[2]
    if ($deploymentTarget -ne "HEAD") {
      $liveDeploymentIds += $deploymentId
    }
  }
}
if ($liveDeploymentIds.Count -eq 0) {
  Write-Host "No live deployment found (non-HEAD). Skipping deploy update."
} else {
  foreach ($deploymentId in $liveDeploymentIds) {
    clasp deploy -i $deploymentId -V $versionNumber -d $CommitMessage
  }
}

Write-Host "Done. GitHub + Apps Script (including deployment) are updated."
