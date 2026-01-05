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

# Optional: use local venv for reproducible builds
if (-not (Test-Path '.venv')) { & $python.Path -m venv .venv }
& .\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
# Runtime deps used by the app
python -m pip install pyinstaller openpyxl

# Clean previous build artifacts
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue

# Build the exe
# Notes:
# - --collect-all openpyxl ensures openpyxl resources are bundled
# - --noconsole suppresses the console window
pyinstaller `
  --onefile `
  --noconsole `
  --name "$AppName" `
  --collect-all openpyxl `
  "$Entry"

Write-Host "Built exe: dist\$AppName.exe"