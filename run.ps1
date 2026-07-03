$ErrorActionPreference = "Stop"

$projectRoot = $PSScriptRoot
$python = Join-Path $projectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $python)) {
    Write-Host "Dang tao moi truong Python..."
    python -m venv (Join-Path $projectRoot ".venv")
}

& $python -c "import openpyxl, PySide6, win32com.client" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Dang cai thu vien..."
    & $python -m pip install -e $projectRoot
}

& $python -m agribank_v3
