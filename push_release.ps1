# Збереження гілки main на GitHub (резервна копія документації й структури).
# Тег пушиться лише якщо вказано -Tag (релізи на GitHub не ведемо).
# Приклад:
#   .\push_release.ps1
#   .\push_release.ps1 -RepoUrl 'https://github.com/USER/PDF_Rename.git'
param(
    [string] $RepoUrl = "",
    [string] $Tag = "",
    [string] $Branch = "main"
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
  .\push_release.ps1 -RepoUrl 'https://github.com/USER/REPO.git' -Tag $Tag
"@
    exit 1
}

Write-Host "Запуск тестів..."
python -m pytest tests -q
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "git push -u origin $Branch"
git push -u origin $Branch
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

if ($Tag.Trim()) {
    Write-Host "git push origin $Tag"
    git push origin $Tag
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

Write-Host @"
Готово: гілка $Branch збережена на GitHub.
  Збірка й випуск — локально (build_win7.bat, docs/ZBIRKA_I_VYPUSK.md).
"@
