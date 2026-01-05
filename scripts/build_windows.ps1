#requires -Version 5.0
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Build Windows exe (onefile + noconsole)
# Target: convert_interlinear_gui.py
# Output: dist/win/Interlinear Converter.exe

$AppName = 'Interlinear Converter'
$Entry = 'convert_interlinear_gui.py'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
Set-Location $ProjectRoot

# Windows-specific output directories (isolate from mac builds)
$DistDir = Join-Path $PWD 'dist\win'
$BuildDir = Join-Path $PWD 'build\win'

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

# Clean previous Windows build artifacts only
Remove-Item -Recurse -Force $BuildDir, $DistDir -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $BuildDir | Out-Null
New-Item -ItemType Directory -Force -Path $DistDir | Out-Null

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
  --distpath "$DistDir" `
  --workpath "$BuildDir" `
  --specpath "$BuildDir" `
  --icon (Join-Path $PWD 'assets\app.ico') `
  "$Entry"

Write-Host "Built exe: $DistDir\$AppName.exe"

# Create high-compression portable zip using 7-Zip if available, else fallback
$Version = (& git describe --tags --abbrev=0 2>$null)
if (-not $Version) { $Version = 'latest' }
$exePath = Join-Path $DistDir "$AppName.exe"
$zipOut  = Join-Path $DistDir ("{0}.{1}.portable.exe.zip" -f $AppName, $Version)

if (-not (Test-Path $exePath)) {
  Write-Warning "Executable not found at $exePath; skipping zip packaging."
} else {
  $sevenZipCandidates = @(
    (Join-Path $env:ProgramFiles '7-Zip\7z.exe'),
    (Join-Path ${env:ProgramFiles(x86)} '7-Zip\7z.exe'),
    (Get-Command 7z -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue)
  $sevenZip = $sevenZipCandidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
  if ($sevenZip) {
    Write-Host "Creating portable zip with 7-Zip: $zipOut"
    & $sevenZip 'a' '-tzip' '-mx=9' '-mm=Deflate64' '-mfb=258' '-mpass=15' $zipOut $exePath | Out-Null
  } else {
    Write-Host "7-Zip not found; using Compress-Archive (Optimal)"
    Compress-Archive -Path $exePath -DestinationPath $zipOut -CompressionLevel Optimal -Force
  }
  Write-Host "Created portable zip: $zipOut"
}