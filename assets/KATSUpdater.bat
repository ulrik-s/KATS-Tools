@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "REPO_OWNER=ulrik-s"
set "REPO_NAME=KATS-Tools"
set "INSTALL_DIR=%APPDATA%\Microsoft\Word\STARTUP"
set "LOG_FILE=%TEMP%\KATSUpdater.log"

rem Relaunch from a temp copy so the installed updater can update itself
if /I not "%~1"=="--worker" (
    set "SELF=%~f0"
    set "WORKER=%TEMP%\KATSUpdater-worker-%RANDOM%%RANDOM%.bat"
    copy /Y "%SELF%" "%WORKER%" >NUL
    if errorlevel 1 (
        echo Failed to create temporary worker updater.
        exit /b 1
    )
    start "" cmd /c ""%WORKER%" --worker"
    exit /b 0
)

echo ==== KATS Updater ==== > "%LOG_FILE%"
echo Install dir: %INSTALL_DIR% >> "%LOG_FILE%"

set "CURRENT_VERSION=0.0.0"
if exist "%INSTALL_DIR%\KATS-Version.txt" (
    set /P CURRENT_VERSION=<"%INSTALL_DIR%\KATS-Version.txt"
)
if "%CURRENT_VERSION%"=="" set "CURRENT_VERSION=0.0.0"

echo Current version: %CURRENT_VERSION% >> "%LOG_FILE%"

set "TEMP_DIR=%TEMP%\KATSUpdate_%RANDOM%%RANDOM%"
set "META_FILE=%TEMP_DIR%\meta.txt"
set "ZIP_FILE=%TEMP_DIR%\KATS-Tools-windows-update.zip"
set "UNPACK_DIR=%TEMP_DIR%\payload"

mkdir "%TEMP_DIR%" >NUL 2>&1
if errorlevel 1 (
    echo Failed to create temp directory: %TEMP_DIR%
    echo Failed to create temp directory >> "%LOG_FILE%"
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$owner=$args[0]; $repo=$args[1]; $current=$args[2]; $meta=$args[3];" ^
  "function Norm([string]$v){ if([string]::IsNullOrWhiteSpace($v)){ return '0.0.0' }; if($v.StartsWith('v') -or $v.StartsWith('V')){ return $v.Substring(1) }; return $v }" ^
  "function Compare-Version([string]$a,[string]$b){ $aa=(Norm $a).Split('.'); $bb=(Norm $b).Split('.'); $len=[Math]::Max($aa.Length,$bb.Length); for($i=0;$i -lt $len;$i++){ $av=if($i -lt $aa.Length){ [int]$aa[$i] } else { 0 }; $bv=if($i -lt $bb.Length){ [int]$bb[$i] } else { 0 }; if($av -lt $bv){ return -1 }; if($av -gt $bv){ return 1 } }; return 0 }" ^
  "$release = Invoke-RestMethod -Headers @{Accept='application/vnd.github+json'; 'X-GitHub-Api-Version'='2022-11-28'} -Uri ('https://api.github.com/repos/{0}/{1}/releases/latest' -f $owner,$repo);" ^
  "$latest = Norm $release.tag_name;" ^
  "if((Compare-Version $current $latest) -ge 0){ Set-Content -LiteralPath $meta -Value 'UPTODATE' -Encoding ASCII; exit 0 }" ^
  "$asset = $release.assets | Where-Object { $_.name -eq 'KATS-Tools-windows-update.zip' } | Select-Object -First 1;" ^
  "if(-not $asset){ throw 'Release asset KATS-Tools-windows-update.zip not found.' }" ^
  "@('LATEST=' + $latest, 'URL=' + $asset.browser_download_url) | Set-Content -LiteralPath $meta -Encoding ASCII" ^
  -- "%REPO_OWNER%" "%REPO_NAME%" "%CURRENT_VERSION%" "%META_FILE%"
if errorlevel 1 goto :fail

set "FIRST_LINE="
set /P FIRST_LINE=<"%META_FILE%"
if /I "%FIRST_LINE%"=="UPTODATE" (
    echo Already up to date. >> "%LOG_FILE%"
    echo You are already running the latest version.
    goto :cleanup_ok
)

set "LATEST_VERSION="
set "DOWNLOAD_URL="
for /F "usebackq tokens=1,* delims==" %%A in ("%META_FILE%") do (
    if /I "%%A"=="LATEST" set "LATEST_VERSION=%%B"
    if /I "%%A"=="URL" set "DOWNLOAD_URL=%%B"
)

if "%DOWNLOAD_URL%"=="" (
    echo Failed to determine update download URL. >> "%LOG_FILE%"
    goto :fail
)

echo New version available: %LATEST_VERSION%
echo Downloading update package...
echo Download URL: %DOWNLOAD_URL% >> "%LOG_FILE%"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop'; Invoke-WebRequest -UseBasicParsing -Uri $args[0] -OutFile $args[1]" ^
  -- "%DOWNLOAD_URL%" "%ZIP_FILE%"
if errorlevel 1 goto :fail

echo Please close Word to continue the update.
echo Waiting for WINWORD.EXE to exit...

:wait_word
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I "WINWORD.EXE" >NUL
if not errorlevel 1 (
    timeout /t 2 /nobreak >NUL
    goto :wait_word
)

mkdir "%UNPACK_DIR%" >NUL 2>&1

echo Extracting update package...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop'; Expand-Archive -LiteralPath $args[0] -DestinationPath $args[1] -Force" ^
  -- "%ZIP_FILE%" "%UNPACK_DIR%"
if errorlevel 1 goto :fail

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo Installing files...

if exist "%UNPACK_DIR%\KATS-Tools.dotm" copy /Y "%UNPACK_DIR%\KATS-Tools.dotm" "%INSTALL_DIR%\KATS-Tools.dotm" >NUL
if errorlevel 1 goto :fail

if exist "%UNPACK_DIR%\KATS-Version.txt" copy /Y "%UNPACK_DIR%\KATS-Version.txt" "%INSTALL_DIR%\KATS-Version.txt" >NUL
if errorlevel 1 goto :fail

if exist "%UNPACK_DIR%\KATSUpdater.bat" copy /Y "%UNPACK_DIR%\KATSUpdater.bat" "%INSTALL_DIR%\KATSUpdater.bat" >NUL

echo Installed version %LATEST_VERSION%. >> "%LOG_FILE%"
echo Installed version %LATEST_VERSION%.
echo Start Word again to load the new version.
goto :cleanup_ok

:fail
echo.
echo Update failed.
echo See log: %LOG_FILE%
goto :cleanup_fail

:cleanup_ok
rmdir /S /Q "%UNPACK_DIR%" >NUL 2>&1
del /Q "%ZIP_FILE%" >NUL 2>&1
del /Q "%META_FILE%" >NUL 2>&1
rmdir /S /Q "%TEMP_DIR%" >NUL 2>&1
exit /b 0

:cleanup_fail
exit /b 1
