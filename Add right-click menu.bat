@echo off
setlocal EnableDelayedExpansion
title iConvert - Add right-click menu
set "APPDIR=%LOCALAPPDATA%\iConvert"
if not exist "%APPDIR%\file_converter.py" (
  echo Please run INSTALL.bat first, then run this.
  pause
  exit /b 1
)
set "PYW="
for /f "usebackq delims=" %%P in (`py -3 -c "import sys,os;print(os.path.join(sys.base_prefix,'pythonw.exe'))" 2^>nul`) do set "PYW=%%P"
if not exist "!PYW!" for /f "usebackq delims=" %%P in (`python -c "import sys,os;print(os.path.join(sys.base_prefix,'pythonw.exe'))" 2^>nul`) do set "PYW=%%P"
if not exist "!PYW!" (
  echo Could not find Python. Run INSTALL.bat first.
  pause
  exit /b 1
)
set "KEY=HKCU\Software\Classes\*\shell\iConvert"
reg add "%KEY%" /ve /d "Convert with iConvert" /f >nul
reg add "%KEY%" /v Icon /d "%APPDIR%\icon.ico" /f >nul
reg add "%KEY%\command" /ve /d "\"%PYW%\" \"%APPDIR%\file_converter.py\" \"%%1\"" /f >nul
echo.
echo Done! Right-click any file in Explorer and choose "Convert with iConvert".
echo.
pause
