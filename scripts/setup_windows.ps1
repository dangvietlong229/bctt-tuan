Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

$root = Split-Path -Parent $PSScriptRoot
$venvPath = Join-Path $root ".venv"
$venvPython = Join-Path $venvPath "Scripts\python.exe"
$requirements = Join-Path $root "requirements-windows.txt"

function Find-PythonLauncher {
    if (Get-Command py.exe -ErrorAction SilentlyContinue) {
        return [pscustomobject]@{ Executable = "py.exe"; Arguments = @("-3") }
    }
    $installedPythons = @(
        Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA "Programs\Python\Python*\python.exe") -ErrorAction SilentlyContinue |
            Sort-Object {
                $folder = Split-Path -Leaf (Split-Path -Parent $_.FullName)
                $digits = $folder -replace '^Python', ''
                if ($digits -match '^\d+$') { [int]$digits } else { 0 }
            } -Descending
    )
    if ($installedPythons.Count -gt 0) {
        return [pscustomobject]@{ Executable = $installedPythons[0].FullName; Arguments = @() }
    }
    if (Get-Command python.exe -ErrorAction SilentlyContinue) {
        return [pscustomobject]@{ Executable = "python.exe"; Arguments = @() }
    }
    throw @"
Python was not found.
Install Python 3.12 or newer from https://www.python.org/downloads/windows/
Select "Add python.exe to PATH" in the installer, then run this setup again.
"@
}

try {
    Set-Location $root
    if (-not (Test-Path -LiteralPath $requirements -PathType Leaf)) {
        throw "Requirements file not found: $requirements"
    }

    if (-not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
        $launcher = Find-PythonLauncher
        & $launcher.Executable @($launcher.Arguments) -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 12) else 1)"
        if ($LASTEXITCODE -ne 0) { throw "Python 3.12 or newer is required." }
        Write-Host "Creating the project Python environment at .venv..." -ForegroundColor Cyan
        & $launcher.Executable @($launcher.Arguments) -m venv $venvPath
        if ($LASTEXITCODE -ne 0) { throw "Could not create the .venv Python environment." }
    }

    Write-Host "Installing/updating Python packages..." -ForegroundColor Cyan
    & $venvPython -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) { throw "Could not update pip." }
    & $venvPython -m pip install -r $requirements
    if ($LASTEXITCODE -ne 0) { throw "Could not install packages from requirements-windows.txt." }

    Write-Host "Checking Excel, PowerPoint and program files..." -ForegroundColor Cyan
    & $venvPython (Join-Path $root "weekly_report.py") doctor
    if ($LASTEXITCODE -ne 0) { throw "Environment check failed." }

    Write-Host ""
    Write-Host "SETUP COMPLETE." -ForegroundColor Green
    Write-Host "Use 1_Tao_ban_nhap_Windows.bat and 2_Xuat_bao_cao_Windows.bat for weekly runs."
    exit 0
}
catch {
    Write-Host ""
    Write-Host "SETUP FAILED: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
