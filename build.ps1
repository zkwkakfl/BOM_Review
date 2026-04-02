# BOM_Review 단일 exe 빌드 (Windows)
# 사용: .\build.ps1

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

try {
    $py = (Get-Command python -ErrorAction Stop).Source
} catch {
    $py = "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe"
}
if (-not (Test-Path $py)) {
    Write-Error "python 을 찾을 수 없습니다. PATH에 Python을 추가하거나 Python312 경로를 확인하세요."
}

& $py -m pip install -q -r requirements-dev.txt
& $py -m PyInstaller --noconfirm BOM_Review.spec

Write-Host "완료: dist\BOM_Review.exe"
