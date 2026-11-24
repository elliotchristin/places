# Stamp a new build id (so iPhones load the new files instead of cached ones), commit and push.
# Usage: ./deploy.ps1 "commit message"
param([Parameter(Mandatory = $true)][string]$Message)
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
$build = 'BUILD_' + (Get-Date -Format 'yyyyMMddHHmmss')
$html = Get-Content index.html -Raw
$html = [regex]::Replace($html, 'BUILD_[0-9a-z]+', $build)
[IO.File]::WriteAllText("$PSScriptRoot\index.html", $html)
[IO.File]::WriteAllText("$PSScriptRoot\version.json", "{`"build`":`"$build`"}")
node --check app.js
node --check mapstyle.js
git add -A
git commit -q -m "$Message`n`nCo-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
git push -q
Write-Output "Deployed $build"
