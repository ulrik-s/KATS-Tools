@echo off
setlocal

set "KATS_URL=%~1"
set "KATS_DESTDIR=%~2"
set "KATS_TARGETNAME=%~3"
set "KATS_TMP=%TEMP%\%KATS_TARGETNAME%.download"

if "%KATS_URL%"=="" exit /b 1
if "%KATS_DESTDIR%"=="" exit /b 2
if "%KATS_TARGETNAME%"=="" exit /b 3

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing -Uri $env:KATS_URL -OutFile $env:KATS_TMP"
if errorlevel 1 exit /b 10

:waitloop
tasklist /FI "IMAGENAME eq WINWORD.EXE" | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  timeout /t 2 /nobreak >nul
  goto waitloop
)

if not exist "%KATS_DESTDIR%" mkdir "%KATS_DESTDIR%"

copy /Y "%KATS_TMP%" "%KATS_DESTDIR%\%KATS_TARGETNAME%" >nul
del /Q "%KATS_TMP%" >nul 2>nul

powershell -NoProfile -Command ^
  "[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms');" ^
  "[System.Windows.Forms.MessageBox]::Show('KATS-Tools uppdaterad. Starta Word igen.','KATS-Tools')"

endlocal
exit /b 0

