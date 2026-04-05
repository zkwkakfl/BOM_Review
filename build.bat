@echo off
rem 빌드: build.bat  ^|  정식: build.bat formal  ^|  창 안 멈춤: ... nopause
call "%~dp0scripts\build.bat" %*
exit /b %errorlevel%
