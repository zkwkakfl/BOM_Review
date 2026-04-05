# 일반 빌드 (프로젝트 루트 기준). 루트 build.ps1 이 여기로 위임.
$ErrorActionPreference = "Stop"
Set-Location (Resolve-Path "$PSScriptRoot\..").Path

Write-Host "===== BOM_Review 일반 빌드 =====" -ForegroundColor Cyan

try {
    $py = (Get-Command python -ErrorAction Stop).Source
} catch {
    $py = $null
    foreach ($p in @(
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe"
    )) {
        if (Test-Path $p) { $py = $p; break }
    }
}
if (-not $py -or -not (Test-Path $py)) {
    Write-Error "python 을 찾을 수 없습니다."
}

Write-Host "Python: $py"
& $py --version

& $py -m pip install -q -r requirements-dev.txt
& $py -m PyInstaller --noconfirm BOM_Review.spec

$out = Join-Path $PWD "dist\BOM_Review.exe"
if (-not (Test-Path $out)) {
    Write-Error "빌드 실패: dist\BOM_Review.exe 가 없습니다."
}
Write-Host "완료: $out"
