@echo off
setlocal
title iConvert - Uninstall
echo Removing iConvert shortcuts and app files...
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\iConvert.lnk" 2>nul
del "%USERPROFILE%\Desktop\iConvert.lnk" 2>nul
if exist "%LOCALAPPDATA%\iConvert" rmdir /s /q "%LOCALAPPDATA%\iConvert"
echo Done. (Python and pip packages were left untouched.)
pause
