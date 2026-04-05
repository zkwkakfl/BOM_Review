@echo off
rem ============================================================
rem  BOM_Review — 일반 빌드 (개발·빠른 재빌드)
rem  사용: 더블클릭 /  build.bat
rem  창 없이(CI):     build.bat nopause
rem  산출물:         dist\BOM_Review.exe
rem ============================================================
chcp 65001 >nul
title BOM_Review — 일반 빌드
setlocal EnableExtensions
cd /d "%~dp0"
set "ERR=0"

echo.
echo ===== BOM_Review 일반 빌드 =====
echo.

call "%~dp0_resolve_python.bat"
if errorlevel 1 (
    echo [오류] python 을 찾을 수 없습니다.
    echo     - PATH에 Python 3.12+ 를 등록하거나
    echo     - python.org 에서 설치해 주세요. ^(Windows 스토어만 있으면 실행이 안 될 수 있습니다.^)
    set ERR=1
    goto :finish
)

echo [확인] 사용 Python:
"%PY%" --version
if errorlevel 1 (
    echo [오류] Python 실행 실패. 'python' 이 스토어 플레이스홀더일 수 있습니다.
    set ERR=1
    goto :finish
)
echo.

echo [1/2] pip install -r requirements-dev.txt ...
"%PY%" -m pip install -q -r requirements-dev.txt
if errorlevel 1 (
    echo [오류] pip install 실패. 인터넷 연결과 requirements-dev.txt 를 확인하세요.
    set ERR=1
    goto :finish
)

echo [2/2] PyInstaller ^(시간이 수 분 걸릴 수 있습니다^)...
"%PY%" -m PyInstaller --noconfirm BOM_Review.spec
if errorlevel 1 (
    echo [오류] PyInstaller 실패. 위 로그를 스크롤해 원인을 확인하세요.
    set ERR=1
    goto :finish
)

if not exist "%~dp0dist\BOM_Review.exe" (
    echo [오류] dist\BOM_Review.exe 가 없습니다.
    set ERR=1
    goto :finish
)

echo.
echo ===== 완료 =====
echo 산출물: %~dp0dist\BOM_Review.exe

:finish
echo.
if %ERR% neq 0 echo 빌드 실패 ^(종료 코드 %ERR%^).
if /i not "%~1"=="nopause" (
    echo 아무 키나 누르면 이 창을 닫습니다...
    pause >nul
)
exit /b %ERR%
