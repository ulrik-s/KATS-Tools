@echo off
setlocal EnableExtensions

set "KATS_URL=%~1"
set "KATS_DESTDIR=%~2"
set "KATS_TARGETNAME=%~3"
set "KATS_TMP=%TEMP%\%~n3.download"

if "%KATS_URL%"=="" exit /b 1
if "%KATS_TARGETNAME%"=="" set "KATS_TARGETNAME=KATS-Tools.dotm"
if "%KATS_TMP%"=="" set "KATS_TMP=%TEMP%\KATS-Tools.download"

REM ------------------------------------------------------------
REM If StartupPath wasn't passed in, ask Word via COM as fallback.
REM Normal case should still be: VBA passes Application.StartupPath.
REM ------------------------------------------------------------
if "%KATS_DESTDIR%"=="" (
  for /f "usebackq delims=" %%I in (`powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$word = New-Object -ComObject Word.Application; try { $p = $word.StartupPath; Write-Output $p } finally { $word.Quit() }"`) do (
    set "KATS_DESTDIR=%%I"
  )
)

if "%KATS_DESTDIR%"=="" exit /b 2

REM ------------------------------------------------------------
REM Download latest dotm
REM ------------------------------------------------------------
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing -Uri $env:KATS_URL -OutFile $env:KATS_TMP"
if errorlevel 1 exit /b 10

REM ------------------------------------------------------------
REM Wait until Word has been closed
REM ------------------------------------------------------------
:waitloop
tasklist /FI "IMAGENAME eq WINWORD.EXE" | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  timeout /t 2 /nobreak >nul
  goto waitloop
)

REM ------------------------------------------------------------
REM Install new version
REM ------------------------------------------------------------
if not exist "%KATS_DESTDIR%" mkdir "%KATS_DESTDIR%"

copy /Y "%KATS_TMP%" "%KATS_DESTDIR%\%KATS_TARGETNAME%" >nul
if errorlevel 1 exit /b 20

del /Q "%KATS_TMP%" >nul 2>nul

powershell -NoProfile -Command ^
  "[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms');" ^
  "[System.Windows.Forms.MessageBox]::Show('KATS-Tools uppdaterad. Starta Word igen.','KATS-Tools')"

endlocal
exit /b 0

