@echo off
setlocal
cd /d "%~dp0"
title iConvert - Update
set "APPDIR=%LOCALAPPDATA%\iConvert"
if not exist "%APPDIR%" (
  echo iConvert is not installed yet. Run INSTALL.bat first.
  pause
  exit /b 1
)
echo Updating the installed iConvert with the files in this folder...
copy /y "file_converter.py" "%APPDIR%\" >nul
copy /y "converters.py"     "%APPDIR%\" >nul
copy /y "version.txt"       "%APPDIR%\" >nul
copy /y "icon.ico"          "%APPDIR%\" >nul
echo.
echo Done! Close iConvert if it is open, then reopen it.
pause
