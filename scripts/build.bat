@echo off
rem 통합 빌드
rem   build.bat              — 빠른 재빌드 (build 폴더 유지)
rem   build.bat formal       — 정식: build·dist 삭제 후 전부 재생성 + BOM_Review_v버전.exe
rem   build.bat nopause      — 끝에서 창 유지 안 함 (둘 다: build.bat formal nopause)
chcp 65001 >nul
setlocal EnableExtensions

set "FORMAL=0"
set "NOPAUSE=0"
:arg
if "%~1"=="" goto :args_done
if /i "%~1"=="formal" set "FORMAL=1"
if /i "%~1"=="nopause" set "NOPAUSE=1"
if /i not "%~1"=="formal" if /i not "%~1"=="nopause" echo [경고] 알 수 없는 인자: %~1
shift
goto :arg
:args_done

cd /d "%~dp0\.."
set "ERR=0"

if "%FORMAL%"=="1" (
    title BOM_Review — 정식 빌드
    echo.
    echo ===== BOM_Review 정식 빌드 =====
    echo   build, dist 를 지운 뒤 처음부터 다시 만듭니다.
) else (
    title BOM_Review — 일반 빌드
    echo.
    echo ===== BOM_Review 일반 빌드 =====
)
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
echo.

if "%FORMAL%"=="1" (
    for /f "usebackq delims=" %%V in (`"%PY%" -c "exec(open('bom_review/_version.py',encoding='utf-8').read()); print(__version__)"`) do set "APPVER=%%V"
    echo [확인] 패키지 버전: %APPVER%
    echo.
    echo [정리] rmdir build, dist ...
    if exist "build" rmdir /s /q "build"
    if exist "dist" rmdir /s /q "dist"
    echo.
)

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
    echo [오류] PyInstaller 실패.
    set ERR=1
    goto :finish
)

if not exist "dist\BOM_Review.exe" (
    echo [오류] dist\BOM_Review.exe 가 없습니다.
    set ERR=1
    goto :finish
)

if "%FORMAL%"=="1" (
    copy /y "dist\BOM_Review.exe" "dist\BOM_Review_v%APPVER%.exe" >nul
    if errorlevel 1 (
        echo [경고] 버전 복사본 실패.
    ) else (
        echo [복사] BOM_Review_v%APPVER%.exe
    )
    echo.
    echo ===== 정식 빌드 완료 =====
    echo   %CD%\dist\BOM_Review.exe
    echo   %CD%\dist\BOM_Review_v%APPVER%.exe
) else (
    echo.
    echo ===== 완료 =====
    echo 산출물: %CD%\dist\BOM_Review.exe
)

:finish
echo.
if %ERR% neq 0 echo 빌드 실패 ^(종료 코드 %ERR%^).
if "%NOPAUSE%"=="0" (
    echo 아무 키나 누르면 이 창을 닫습니다...
    pause >nul
)
exit /b %ERR%
