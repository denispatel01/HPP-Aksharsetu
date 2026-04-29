param(
  [Parameter(Mandatory = $true)]
  [string]$CommitMessage
)

$ErrorActionPreference = "Stop"

Write-Host "Checking git status..."
git status --short

Write-Host "Creating git commit..."
git add .
git commit -m $CommitMessage

Write-Host "Pushing code to GitHub..."
git push

Write-Host "Pushing code to Apps Script..."
clasp push -f

Write-Host "Done. GitHub + Apps Script are updated."
