@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title iConvert - Installer

echo ============================================================
echo   iConvert - Local File Converter : INSTALLER
echo ============================================================
echo.

set "PY="
call :detect "py -3"
if not defined PY call :detect "python"
if not defined PY call :detect "python3"
if not defined PY goto :no_python

echo [1/3] Found Python: %PY%
%PY% --version
echo.

echo [2/3] Installing components (one-time; can take a few minutes)...
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt
if errorlevel 1 goto :pip_failed
echo.

echo [3/3] Installing iConvert and creating shortcuts...
set "PYW="
for /f "usebackq delims=" %%P in (`%PY% -c "import sys,os;print(os.path.join(sys.base_prefix,'pythonw.exe'))"`) do set "PYW=%%P"
if not exist "!PYW!" for /f "usebackq delims=" %%P in (`%PY% -c "import sys;print(sys.executable)"`) do set "PYW=%%P"

set "APPDIR=%LOCALAPPDATA%\iConvert"
if not exist "%APPDIR%" mkdir "%APPDIR%"
copy /y "file_converter.py" "%APPDIR%\" >nul
copy /y "converters.py"     "%APPDIR%\" >nul
copy /y "version.txt"       "%APPDIR%\" >nul
copy /y "icon.ico"          "%APPDIR%\" >nul

set "ICONV_PYW=!PYW!"
set "ICONV_APPDIR=%APPDIR%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make_shortcuts.ps1"

echo.
echo ============================================================
echo   DONE!  iConvert is installed.
echo     - Press the Windows key and type   iConvert
echo     - Or double-click the iConvert icon on your Desktop
echo ============================================================
echo.
pause
exit /b 0

:no_python
echo [X] Python 3 is not installed on this PC.
echo     The "python" command you have is only a Microsoft Store placeholder.
echo.
echo   HOW TO FIX - takes about 5 minutes:
echo     1. Open  https://www.python.org/downloads/
echo     2. Download Python 3 and run the installer.
echo     3. On the FIRST screen, TICK  "Add python.exe to PATH".
echo     4. Click  "Install Now"  and let it finish.
echo     5. Close this window, then double-click INSTALL.bat again.
echo.
pause
exit /b 1

:pip_failed
echo.
echo [X] Component install failed - check your internet connection and try again.
pause
exit /b 1

:detect
set "TMPCHK=%TEMP%\iconv_pychk.txt"
del "%TMPCHK%" >nul 2>nul
%~1 -c "open(r'%TMPCHK%','w').write('OK')" >nul 2>nul
set "_chk="
if exist "%TMPCHK%" set /p _chk=<"%TMPCHK%"
del "%TMPCHK%" >nul 2>nul
if "%_chk%"=="OK" set "PY=%~1"
exit /b
