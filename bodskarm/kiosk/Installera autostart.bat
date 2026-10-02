@echo off
rem ======================================================================
rem  BODSKARM - kor EN gang pa datorn i boden.
rem   * Skarmen startar automatiskt nar datorn loggar in.
rem   * Skarmslackare/vila stangs av sa att TV:n alltid visar bilden.
rem ======================================================================
chcp 65001 >nul
setlocal

set "BAT=%~dp0Starta bodskarm.bat"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$s = (New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Startup') + '\Bodskarm.lnk');" ^
  "$s.TargetPath = '%BAT%'; $s.Arguments = '/autostart'; $s.WorkingDirectory = '%~dp0'; $s.WindowStyle = 7; $s.Save()"

powercfg /change monitor-timeout-ac 0
powercfg /change standby-timeout-ac 0
powercfg /change hibernate-timeout-ac 0

echo.
echo Klart! Bodskarmen startar nu automatiskt vid inloggning.
echo Tips: stall in automatisk inloggning och att datorn startar efter stromavbrott (BIOS).
echo.
pause
