<#
  Build Doc2Word on a Windows PC: tests -> app folder -> installer.

  Needs: Python 3.12 (python.org, "Add to PATH" ticked) and Inno Setup 6
  (https://jrsoftware.org/isdl.php). Run from the repo root:

      powershell -ExecutionPolicy Bypass -File packaging\build.ps1

  Output: dist\Doc2Word-Setup-<version>.exe
#>
$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot -Parent)

$version = (python -c "import doc2word; print(doc2word.__version__)").Trim()
if (-not (Select-String -Path packaging\version_info.txt -Pattern "'$version'" -Quiet)) {
    throw "packaging\version_info.txt doesn't match version $version - update it first."
}
Write-Host "Building Doc2Word $version" -ForegroundColor Cyan

if (-not (Test-Path .venv)) { python -m venv .venv }
.\.venv\Scripts\python -m pip install --upgrade pip
.\.venv\Scripts\python -m pip install -e ".[dev]"

Write-Host "Running tests..." -ForegroundColor Cyan
.\.venv\Scripts\python -m pytest -q
if ($LASTEXITCODE -ne 0) { throw "Tests failed" }

.\.venv\Scripts\python packaging\make_icon.py packaging\doc2word.ico

Write-Host "Packaging app..." -ForegroundColor Cyan
.\.venv\Scripts\pyinstaller packaging\doc2word.spec --noconfirm --clean
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

$iscc = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw "Inno Setup 6 not found - install it from https://jrsoftware.org/isdl.php" }

Write-Host "Building installer..." -ForegroundColor Cyan
& $iscc "/DAppVersion=$version" packaging\installer.iss
if ($LASTEXITCODE -ne 0) { throw "Inno Setup failed" }

Write-Host "Done: dist\Doc2Word-Setup-$version.exe" -ForegroundColor Green
