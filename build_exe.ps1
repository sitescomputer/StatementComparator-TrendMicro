[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$projectDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $projectDirectory

$launcher = Get-Command py -ErrorAction SilentlyContinue
if ($null -ne $launcher) {
    $pythonCommand = $launcher.Source
} else {
    $launcher = Get-Command python -ErrorAction SilentlyContinue
    if ($null -eq $launcher) {
        throw 'Python was not found. Install Python 3.11 or newer, then rerun this script.'
    }
    $pythonCommand = $launcher.Source
}

& $pythonCommand -m pip install -r requirements-dev.txt
& $pythonCommand -m unittest discover -s tests -v
& $pythonCommand -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name TrendMicroComparator `
    TrendMicroComparator.py

Write-Host "Build complete: $projectDirectory\dist\TrendMicroComparator.exe"
