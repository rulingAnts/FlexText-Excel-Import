#requires -Version 5.0
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Build Windows exe (onefile + noconsole)
# Target: convert_interlinear_gui.py
# Output: dist/Interlinear Converter.exe

$AppName = 'Interlinear Converter'
$Entry = 'convert_interlinear_gui.py'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
Set-Location $ProjectRoot

# Ensure Python is available
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) { $python = Get-Command py -ErrorAction SilentlyContinue }
if (-not $python) { Write-Error 'Python interpreter not found (python or py). Install Python 3.x.' }

# Use a dedicated Windows venv to avoid conflicts with macOS venvs
if (-not (Test-Path '.venv-win')) { & $python.Path -m venv .venv-win }

# Resolve venv python across layouts (Scripts vs bin)
$venvPythonCandidates = @(
  '.venv-win\Scripts\python.exe',
  '.venv-win\Scripts\python',
  '.venv-win\bin\python.exe',
  '.venv-win\bin\python'
)
$venvPython = $venvPythonCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $venvPython) { Write-Error 'Could not locate venv python in .venv-win\Scripts or .venv-win\bin.' }

# Upgrade pip and install deps inside the venv
& $venvPython -m pip install --upgrade pip
# Runtime deps used by the app
& $venvPython -m pip install pyinstaller openpyxl Pillow

# Clean previous build artifacts
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue

# Build the exe
# Notes:
# - --collect-all openpyxl ensures openpyxl resources are bundled
# - --noconsole suppresses the console window
# Generate icon if missing (optional helper; safe to skip if absent)
if (-not (Test-Path 'assets/app.ico')) { & $venvPython scripts/generate_icons.py }

& $venvPython -m PyInstaller `
  --onefile `
  --noconsole `
  --name "$AppName" `
  --collect-all openpyxl `
  --icon assets/app.ico `
  "$Entry"

Write-Host "Built exe: dist\$AppName.exe"