@echo off
chcp 65001 >nul
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Bodskarm.lnk" 2>nul
echo Autostart borttagen.
pause
