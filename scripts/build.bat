@echo off
rem 일반 빌드 — 루트의 build.bat 이 여기로 위임됨.
chcp 65001 >nul
title BOM_Review — 일반 빌드
setlocal EnableExtensions
cd /d "%~dp0\.."
set "ERR=0"

echo.
echo ===== BOM_Review 일반 빌드 =====
echo.

call "%~dp0resolve_python.bat"
if errorlevel 1 (
    echo [오류] python 을 찾을 수 없습니다.
    echo     PATH 또는 python.org 설치를 확인하세요. ^(Windows 스토어 스텁만 있으면 실패할 수 있습니다.^)
    set ERR=1
    goto :finish
)

echo [확인] 사용 Python:
"%PY%" --version
if errorlevel 1 (
    echo [오류] Python 실행 실패.
    set ERR=1
    goto :finish
)
echo.

echo [1/2] pip install -r requirements-dev.txt ...
"%PY%" -m pip install -q -r requirements-dev.txt
if errorlevel 1 (
    echo [오류] pip install 실패.
    set ERR=1
    goto :finish
)

echo [2/2] PyInstaller ^(수 분 소요^)...
"%PY%" -m PyInstaller --noconfirm BOM_Review.spec
if errorlevel 1 (
    echo [오류] PyInstaller 실패. 로그를 확인하세요.
    set ERR=1
    goto :finish
)

if not exist "dist\BOM_Review.exe" (
    echo [오류] dist\BOM_Review.exe 가 없습니다.
    set ERR=1
    goto :finish
)

echo.
echo ===== 완료 =====
echo 산출물: %CD%\dist\BOM_Review.exe

:finish
echo.
if %ERR% neq 0 echo 빌드 실패 ^(종료 코드 %ERR%^).
if /i not "%~1"=="nopause" (
    echo 아무 키나 누르면 이 창을 닫습니다...
    pause >nul
)
exit /b %ERR%
