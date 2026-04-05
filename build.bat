@echo off
rem → scripts\build.bat  (루트에서 더블클릭용 래퍼)
call "%~dp0scripts\build.bat" %*
exit /b %errorlevel%
