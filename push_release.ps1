# Публікація гілки main і тега v1.2.0 на GitHub.
# Приклад:
#   .\push_release.ps1 -RepoUrl 'https://github.com/USER/PDF_Rename_Expert.git'
# Якщо origin уже додано:
#   .\push_release.ps1
param(
    [string] $RepoUrl = ""
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$remotes = @(git remote 2>$null)
$hasOrigin = $remotes -contains "origin"

if ($RepoUrl -and -not $hasOrigin) {
    git remote add origin $RepoUrl
    Write-Host "Додано remote: origin -> $RepoUrl"
    $hasOrigin = $true
}

if (-not $hasOrigin) {
    Write-Error @"
Немає remote 'origin'. Виконайте один із варіантів:
  git remote add origin https://github.com/USER/REPO.git
  .\push_release.ps1 -RepoUrl 'https://github.com/USER/REPO.git'
"@
    exit 1
}

Write-Host "Запуск тестів..."
python -m pytest tests -q
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "git push -u origin main"
git push -u origin main
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "git push origin v1.2.0"
git push origin v1.2.0
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "Готово. GitHub -> Releases -> створіть реліз для тега v1.2.0; за потреби додайте exe з dist\."
