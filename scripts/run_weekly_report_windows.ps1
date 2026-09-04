param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("all", "finalize")]
    [string]$ReportCommand,
    [string]$ExcludeModules = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

$root = Split-Path -Parent $PSScriptRoot
$python = Join-Path $root ".venv\Scripts\python.exe"
$program = Join-Path $root "weekly_report.py"

try {
    Set-Location $root
    if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
        throw "Python environment not found. Run 0_Cai_dat_Windows.bat first."
    }
    if ($ReportCommand -eq "all" -and [string]::IsNullOrWhiteSpace($ExcludeModules)) {
        Write-Host ""
        Write-Host "CAC MODULE XU LY DU LIEU:" -ForegroundColor Cyan
        Write-Host "  2. Nganh"
        Write-Host "  3. Danh muc MBS"
        Write-Host "  4. Giao dich Nuoc ngoai"
        Write-Host "  5. Thanh khoan thi truong"
        Write-Host "  6. Dong gop cua Vingroup"
        Write-Host "  7. Nuoc ngoai ban rong"
        Write-Host "  8. Giao dich Tu doanh"
        Write-Host "  9. Tin doanh nghiep"
        Write-Host " 10. Lich su kien"
        Write-Host ""
        Write-Host "Nhan Enter de chay tat ca." -ForegroundColor Green
        $selection = Read-Host "Nhap cac module muon TAT, cach nhau boi dau phay (vi du: 3,9)"
        $selection = $selection.Trim()
        if (-not [string]::IsNullOrWhiteSpace($selection)) {
            if ($selection -notmatch '^\s*(?:[2-9]|10)(?:\s*,\s*(?:[2-9]|10))*\s*$') {
                throw "Danh sach module khong hop le. Chi dung cac so 2-10, cach nhau boi dau phay."
            }
            $ExcludeModules = (($selection -split ',') | ForEach-Object { $_.Trim() } | Select-Object -Unique) -join ','
            Write-Host "Se bo qua module: $ExcludeModules" -ForegroundColor Yellow
        }
        else {
            Write-Host "Se chay tat ca module." -ForegroundColor Green
        }
        Write-Host ""
    }

    if ($ReportCommand -eq "all" -and -not [string]::IsNullOrWhiteSpace($ExcludeModules)) {
        & $python $program $ReportCommand --exclude-modules $ExcludeModules
    }
    else {
        & $python $program $ReportCommand
    }
    exit $LASTEXITCODE
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
