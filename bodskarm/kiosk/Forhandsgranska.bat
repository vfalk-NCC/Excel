@echo off
rem ======================================================================
rem  BODSKARM - forhandsgranska pa din egen dator i ett vanligt fonster.
rem  Piltangenter = byt vy, mellanslag = pausa, R = las om Excel-filerna.
rem ======================================================================
chcp 65001 >nul
setlocal

set "MAPP=%~dp0.."
for %%I in ("%MAPP%") do set "MAPP=%%~fI"
set "SIDA=file:///%MAPP:\=/%/index.html"
set "PROFIL=%LOCALAPPDATA%\Bodskarm\Forhandsgranska"

set "EDGE=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if not exist "%EDGE%" set "EDGE=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"

set "FLAGGOR=--allow-file-access-from-files --user-data-dir="%PROFIL%" --no-first-run --window-size=1600,900"

if exist "%EDGE%" ( start "" "%EDGE%" --app="%SIDA%" %FLAGGOR% & exit /b 0 )
if exist "%CHROME%" ( start "" "%CHROME%" --app="%SIDA%" %FLAGGOR% & exit /b 0 )
echo Hittar varken Microsoft Edge eller Google Chrome.
pause
