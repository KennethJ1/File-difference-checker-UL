<#
build_windows.ps1
PowerShell helper to build a single-folder Windows executable using PyInstaller.
Run from the project root in PowerShell (recommended to run as normal user):
  .\tools\build_windows.ps1

What it does:
 - creates a virtualenv in .venv_build
 - installs dependencies from requirements.txt
 - runs pyinstaller with difference_checker.spec
 - copies the resulting dist\difference_checker to artifacts\difference_checker-<timestamp>
#>

param(
    [string]$SpecFile = "difference_checker.spec",
    [string]$VenvDir = ".venv_build",
    [switch]$Clean
)

Set-StrictMode -Version Latest

$root = Split-Path -Path $PSScriptRoot -Parent
Push-Location $root

if ($Clean) {
    Write-Host "Cleaning previous build artifacts..."
    Remove-Item -Recurse -Force "$VenvDir" -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force "build" -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force "dist" -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force "artifacts" -ErrorAction SilentlyContinue
}

if (-not (Test-Path $VenvDir)) {
    Write-Host "Creating virtual environment in $VenvDir..."
    python -m venv $VenvDir
}

$activate = Join-Path $VenvDir "Scripts\Activate.ps1"
if (-not (Test-Path $activate)) {
    throw "Activate script not found at $activate. Ensure Python is installed and 'python -m venv' works."
}

Write-Host "Activating virtualenv..."
. $activate

Write-Host "Upgrading pip and installing build requirements..."
python -m pip install --upgrade pip setuptools wheel
python -m pip install -r requirements.txt

Write-Host "Running PyInstaller with spec: $SpecFile"
pyinstaller $SpecFile
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE"
}

# copy dist to artifacts with timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$artifactDir = "artifacts\difference_checker_$timestamp"
New-Item -ItemType Directory -Path $artifactDir -Force | Out-Null
Copy-Item -Recurse -Force "dist\difference_checker" $artifactDir

Write-Host "Build finished. Artifacts copied to: $artifactDir"

Pop-Location
