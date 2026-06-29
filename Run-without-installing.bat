@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title iConvert

set "PY="
call :detect "py -3"
if not defined PY call :detect "python"
if not defined PY call :detect "python3"
if not defined PY goto :no_python

if not exist ".deps_installed" (
  echo Installing components, one-time, please wait...
  %PY% -m pip install --upgrade pip
  %PY% -m pip install -r requirements.txt
  if errorlevel 1 ( echo Install failed. & pause & exit /b 1 )
  echo done> ".deps_installed"
)

set "PYW="
for /f "usebackq delims=" %%P in (`%PY% -c "import sys,os;print(os.path.join(sys.base_prefix,'pythonw.exe'))"`) do set "PYW=%%P"
if not exist "!PYW!" set "PYW=%PY%"
start "" "!PYW!" "file_converter.py"
exit /b 0

:no_python
echo Python 3 is not installed. Get it from https://www.python.org/downloads/
echo On the first installer screen, TICK "Add python.exe to PATH", then run this again.
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
