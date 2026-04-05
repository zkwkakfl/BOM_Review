# BOM_Review — 정식 빌드 (PowerShell)
# 사용: .\build_formal.ps1
# 산출물: dist\BOM_Review.exe , dist\BOM_Review_v{version}.exe

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

function Get-PythonExe {
    try {
        return (Get-Command python -ErrorAction Stop).Source
    } catch {
        foreach ($p in @(
            "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
            "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe"
        )) {
            if (Test-Path $p) { return $p }
        }
        throw "python 을 찾을 수 없습니다."
    }
}

Write-Host ""
Write-Host "===== BOM_Review 정식 빌드 =====" -ForegroundColor Cyan

$py = Get-PythonExe
Write-Host "Python: $py"
& $py --version

$verLine = Get-Content -Path "bom_review\_version.py" -Raw -Encoding UTF8
$ver = $null
if ($verLine -match '__version__\s*=\s*"([^"]+)"') {
    $ver = $Matches[1]
}
if (-not $ver) {
    throw "bom_review\_version.py 에서 버전을 읽을 수 없습니다."
}
Write-Host "패키지 버전: $ver"

Write-Host "[정리] build, dist ..."
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }

Write-Host "[1/2] pip install ..."
& $py -m pip install -q -r requirements-dev.txt

Write-Host "[2/2] PyInstaller ..."
& $py -m PyInstaller --noconfirm BOM_Review.spec

$exe = Join-Path $PWD "dist\BOM_Review.exe"
if (-not (Test-Path $exe)) {
    throw "dist\BOM_Review.exe 가 생성되지 않았습니다."
}

$tag = Join-Path $PWD "dist\BOM_Review_v$ver.exe"
Copy-Item $exe $tag -Force
Write-Host ""
Write-Host "===== 정식 빌드 완료 ====="
Write-Host $exe
Write-Host $tag
