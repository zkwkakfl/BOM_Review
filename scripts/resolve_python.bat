@echo off
rem 프로젝트 루트에서 call 할 것. PY 에 Python 경로 설정, 실패 시 errorlevel 1.
set "PY="
where python >nul 2>&1
if %errorlevel% equ 0 (
    set "PY=python"
    exit /b 0
)
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    exit /b 0
)
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    exit /b 0
)
exit /b 1
