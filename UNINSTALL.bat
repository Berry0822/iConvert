@echo off
setlocal
title iConvert - Uninstall
echo Removing iConvert shortcuts, right-click menu and app files...
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\iConvert.lnk" 2>nul
del "%USERPROFILE%\Desktop\iConvert.lnk" 2>nul
reg delete "HKCU\Software\Classes\*\shell\iConvert" /f >nul 2>nul
if exist "%LOCALAPPDATA%\iConvert" rmdir /s /q "%LOCALAPPDATA%\iConvert"
echo Done. (Python and pip packages were left untouched.)
pause
