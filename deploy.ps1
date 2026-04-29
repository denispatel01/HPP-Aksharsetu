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

Write-Host "Done. GitHub + Apps Script are updated."
