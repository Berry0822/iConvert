@echo off
title iConvert - Remove right-click menu
reg delete "HKCU\Software\Classes\*\shell\iConvert" /f >nul 2>nul
echo Removed the "Convert with iConvert" right-click menu.
pause
