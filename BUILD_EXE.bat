@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title iConvert - Build standalone EXE

set "PY="
call :det "py -3"
if not defined PY call :det "python"
if not defined PY ( echo Install Python first (see README). & pause & exit /b 1 )

echo Installing build tools and dependencies (one-time)...
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt pyinstaller
if errorlevel 1 ( echo Install failed. & pause & exit /b 1 )

echo.
echo Building the standalone app... this can take several minutes.
%PY% -m PyInstaller --noconfirm --onefile --windowed --name iConvert ^
  --icon "icon.ico" --add-data "icon.ico;." --add-data "version.txt;." ^
  --collect-all customtkinter --collect-all tkinterdnd2 --collect-all pdf2docx ^
  --hidden-import win32com --hidden-import win32com.client ^
  --hidden-import pythoncom --hidden-import win32timezone ^
  file_converter.py
if not exist "dist\iConvert.exe" ( echo Build failed. & pause & exit /b 1 )

echo.
echo ============================================================
echo   DONE!  Your standalone app is here:
echo     dist\iConvert.exe
echo   Share that single file - it runs without Python.
echo   (OCR still needs the Tesseract program installed.)
echo ============================================================
echo.
pause
exit /b 0

:det
set "TMPCHK=%TEMP%\iconv_pychk.txt"
del "%TMPCHK%" >nul 2>nul
%~1 -c "open(r'%TMPCHK%','w').write('OK')" >nul 2>nul
set "_chk="
if exist "%TMPCHK%" set /p _chk=<"%TMPCHK%"
del "%TMPCHK%" >nul 2>nul
if "%_chk%"=="OK" set "PY=%~1"
exit /b
