@echo off
chcp 65001 >nul
setlocal EnableExtensions
cd /d "%~dp0"

rem BOM_Review 단일 exe 빌드 (Windows)
rem 사용: build.bat 또는 더블클릭

set "PY="
where python >nul 2>&1
if %errorlevel% equ 0 (
    set "PY=python"
    goto :py_ok
)
set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if exist "%PY%" goto :py_ok

echo [오류] python 을 찾을 수 없습니다. PATH에 Python을 추가하거나 Python 3.12 경로를 확인하세요.
exit /b 1

:py_ok
echo Python: %PY%
echo.

"%PY%" -m pip install -q -r requirements-dev.txt
if errorlevel 1 (
    echo [오류] pip install 실패
    exit /b 1
)

"%PY%" -m PyInstaller --noconfirm BOM_Review.spec
if errorlevel 1 (
    echo [오류] PyInstaller 실패
    exit /b 1
)

echo.
echo 완료: dist\BOM_Review.exe
exit /b 0
