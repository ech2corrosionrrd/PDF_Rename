# Публікація гілки main і тега релізу на GitHub.
# Приклад:
#   .\push_release.ps1 -Tag v1.4.0
#   .\push_release.ps1 -RepoUrl 'https://github.com/USER/PDF_Rename.git' -Tag v1.4.0
# Якщо origin уже додано:
#   .\push_release.ps1 -Tag v1.4.0
param(
    [string] $RepoUrl = "",
    [string] $Tag = "v1.4.0",
    [string] $Branch = "main"
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not $Tag.Trim()) {
    Write-Error "Порожній тег. Вкажіть -Tag, наприклад v1.4.0."
    exit 1
}

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

Write-Host "git push origin $Tag"
git push origin $Tag
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host @"
Готово.
  Тег $Tag на GitHub запустить workflow «GitHub Release» (тести + PDF_Rename_Expert.exe).
  Локальна збірка під Win7: build_win7.bat
"@
