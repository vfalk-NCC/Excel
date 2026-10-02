@echo off
rem ======================================================================
rem  BODSKARM - startar skarmen i helskarm (kiosklage) i Microsoft Edge.
rem  Avsluta med Alt+F4.
rem
rem  Webblasaren startas med en egen profil och tillatelse att lasa lokala
rem  filer, sa att sidan kan lasa Excel-filerna direkt fran OneDrive-mappen.
rem ======================================================================
chcp 65001 >nul
setlocal

set "MAPP=%~dp0.."
for %%I in ("%MAPP%") do set "MAPP=%%~fI"
set "SIDA=file:///%MAPP:\=/%/index.html"
set "PROFIL=%LOCALAPPDATA%\Bodskarm\Webblasare"

rem Vid autostart: vanta sa att OneDrive hinner synka efter inloggning.
if /i "%~1"=="/autostart" timeout /t 30 /nobreak >nul

set "EDGE=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if not exist "%EDGE%" set "EDGE=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"

set "FLAGGOR=--allow-file-access-from-files --user-data-dir="%PROFIL%" --no-first-run --noerrdialogs --disable-session-crashed-bubble --disable-features=Translate --overscroll-history-navigation=0"

if exist "%EDGE%" (
  start "" "%EDGE%" --kiosk "%SIDA%" --edge-kiosk-type=fullscreen %FLAGGOR%
  exit /b 0
)
if exist "%CHROME%" (
  start "" "%CHROME%" --kiosk "%SIDA%" %FLAGGOR%
  exit /b 0
)
echo Hittar varken Microsoft Edge eller Google Chrome.
pause
exit /b 1
