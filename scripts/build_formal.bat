@echo off
rem 정식 빌드 — build/dist 삭제 후 재생성, 버전 복사본 생성. 루트 build_formal.bat 이 위임.
chcp 65001 >nul
title BOM_Review — 정식 빌드
setlocal EnableExtensions
cd /d "%~dp0\.."
set "ERR=0"

echo.
echo ===== BOM_Review 정식 빌드 =====
echo   build, dist 폴더를 지운 뒤 처음부터 다시 만듭니다.
echo.

call "%~dp0resolve_python.bat"
if errorlevel 1 (
    echo [오류] python 을 찾을 수 없습니다.
    set ERR=1
    goto :finish
)

echo [확인] 사용 Python:
"%PY%" --version
if errorlevel 1 (
    set ERR=1
    goto :finish
)

for /f "usebackq delims=" %%V in (`"%PY%" -c "exec(open('bom_review/_version.py',encoding='utf-8').read()); print(__version__)"`) do set "APPVER=%%V"
echo [확인] 패키지 버전: %APPVER%
echo.

echo [정리] rmdir build, dist ...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
echo.

echo [1/2] pip install -r requirements-dev.txt ...
"%PY%" -m pip install -q -r requirements-dev.txt
if errorlevel 1 (
    echo [오류] pip install 실패
    set ERR=1
    goto :finish
)

echo [2/2] PyInstaller ^(수 분 소요^)...
"%PY%" -m PyInstaller --noconfirm BOM_Review.spec
if errorlevel 1 (
    echo [오류] PyInstaller 실패
    set ERR=1
    goto :finish
)

if not exist "dist\BOM_Review.exe" (
    echo [오류] dist\BOM_Review.exe 가 없습니다.
    set ERR=1
    goto :finish
)

copy /y "dist\BOM_Review.exe" "dist\BOM_Review_v%APPVER%.exe" >nul
if errorlevel 1 (
    echo [경고] 버전 복사본 실패. BOM_Review.exe 만 사용하세요.
) else (
    echo [복사] BOM_Review_v%APPVER%.exe
)

echo.
echo ===== 정식 빌드 완료 =====
echo   %CD%\dist\BOM_Review.exe
echo   %CD%\dist\BOM_Review_v%APPVER%.exe

:finish
echo.
if %ERR% neq 0 echo 빌드 실패 ^(종료 코드 %ERR%^).
if /i not "%~1"=="nopause" (
    echo 아무 키나 누르면 이 창을 닫습니다...
    pause >nul
)
exit /b %ERR%
