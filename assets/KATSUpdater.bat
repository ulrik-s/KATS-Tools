@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "REPO_OWNER=ulrik-s"
set "REPO_NAME=KATS-Tools"
set "INSTALL_DIR=%APPDATA%\Microsoft\Word\STARTUP"
set "LOG_FILE=%TEMP%\KATSUpdater.log"

if not defined TEMP set "TEMP=%LOCALAPPDATA%\Temp"
if not exist "%TEMP%" mkdir "%TEMP%" >NUL 2>&1

if /I "%~1"=="--checkonly" goto checkonly
if /I "%~1"=="--worker" goto worker

echo ==== KATS Updater ==== > "%LOG_FILE%"
echo TEMP=%TEMP% >> "%LOG_FILE%"
echo APPDATA=%APPDATA% >> "%LOG_FILE%"
echo INSTALL_DIR=%INSTALL_DIR% >> "%LOG_FILE%"

set "SELF=%~f0"
set "WORKER=%TEMP%\KATSUpdater-worker-%RANDOM%%RANDOM%.bat"

echo SELF=%SELF% >> "%LOG_FILE%"
echo WORKER=%WORKER% >> "%LOG_FILE%"

copy /Y "%SELF%" "%WORKER%" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo Failed to create temporary worker updater. >> "%LOG_FILE%"
    call :show_error "Kunde inte starta uppdateraren. Se loggen: %LOG_FILE%"
    exit /b 1
)

start "" /min cmd /c ""%WORKER%" --worker"
exit /b 0

:checkonly
call :read_current_version
call :query_github
if errorlevel 1 (
    echo FAILED^|GitHub query failed
    exit /b 1
)

if /I "%STATUS%"=="UPTODATE" (
    echo UPTODATE
    exit /b 0
)

if /I "%STATUS%"=="UPDATE" (
    echo UPDATE^|%LATEST_VERSION%
    exit /b 0
)

echo FAILED^|Unknown updater status
exit /b 1

:worker
shift

echo ==== KATS Updater ==== > "%LOG_FILE%"
echo TEMP=%TEMP% >> "%LOG_FILE%"
echo APPDATA=%APPDATA% >> "%LOG_FILE%"
echo INSTALL_DIR=%INSTALL_DIR% >> "%LOG_FILE%"

call :read_current_version
echo Current version: %CURRENT_VERSION% >> "%LOG_FILE%"

call :query_github
if errorlevel 1 goto :fail

echo STATUS=%STATUS% >> "%LOG_FILE%"
echo LATEST_VERSION=%LATEST_VERSION% >> "%LOG_FILE%"
echo DOWNLOAD_URL=%DOWNLOAD_URL% >> "%LOG_FILE%"

if /I "%STATUS%"=="UPTODATE" (
    echo Already up to date. >> "%LOG_FILE%"
    goto :cleanup_ok
)

if /I not "%STATUS%"=="UPDATE" (
    echo Invalid update status. >> "%LOG_FILE%"
    goto :fail
)

if "%LATEST_VERSION%"=="" (
    echo Failed to determine latest version. >> "%LOG_FILE%"
    goto :fail
)

if "%DOWNLOAD_URL%"=="" (
    echo Failed to determine update download URL. >> "%LOG_FILE%"
    goto :fail
)

echo New version available: %LATEST_VERSION%
echo Downloading update package...
echo Download URL: %DOWNLOAD_URL% >> "%LOG_FILE%"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop'; Invoke-WebRequest -UseBasicParsing -Uri '%DOWNLOAD_URL%' -OutFile '%ZIP_FILE%'" >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto :fail

echo Waiting for WINWORD.EXE to exit... >> "%LOG_FILE%"
set "WAIT_COUNT=0"

:wait_word
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>NUL | find /I "WINWORD.EXE" >NUL
if not errorlevel 1 (
    set /A WAIT_COUNT+=1
    if !WAIT_COUNT! EQU 1 echo Waiting for Word to close... >> "%LOG_FILE%"
    timeout /t 2 /nobreak >NUL
    goto :wait_word
)

mkdir "%UNPACK_DIR%" >NUL 2>&1

echo Extracting update package...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop'; Expand-Archive -LiteralPath '%ZIP_FILE%' -DestinationPath '%UNPACK_DIR%' -Force" >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto :fail

if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo Installing files...

if exist "%UNPACK_DIR%\KATS-Tools.dotm" copy /Y "%UNPACK_DIR%\KATS-Tools.dotm" "%INSTALL_DIR%\KATS-Tools.dotm" >NUL
if errorlevel 1 goto :fail

if exist "%UNPACK_DIR%\KATS-Version.txt" copy /Y "%UNPACK_DIR%\KATS-Version.txt" "%INSTALL_DIR%\KATS-Version.txt" >NUL
if errorlevel 1 goto :fail

if exist "%UNPACK_DIR%\KATSUpdater.bat" copy /Y "%UNPACK_DIR%\KATSUpdater.bat" "%INSTALL_DIR%\KATSUpdater.bat" >NUL

echo Installed version %LATEST_VERSION%. >> "%LOG_FILE%"
call :show_info "KATS-Tools har uppdaterats till version %LATEST_VERSION%. Starta Word igen."
goto :cleanup_ok

:read_current_version
set "CURRENT_VERSION=0.0.0"
if exist "%INSTALL_DIR%\KATS-Version.txt" (
    set /P CURRENT_VERSION=<"%INSTALL_DIR%\KATS-Version.txt"
)
if "%CURRENT_VERSION%"=="" set "CURRENT_VERSION=0.0.0"

set "TEMP_DIR=%TEMP%\KATSUpdate_%RANDOM%%RANDOM%"
set "STATUS_FILE=%TEMP_DIR%\status.txt"
set "LATEST_FILE=%TEMP_DIR%\latest.txt"
set "URL_FILE=%TEMP_DIR%\url.txt"
set "ZIP_FILE=%TEMP_DIR%\KATS-Tools-windows-update.zip"
set "UNPACK_DIR=%TEMP_DIR%\payload"

mkdir "%TEMP_DIR%" >NUL 2>&1
exit /b 0

:query_github
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$owner = '%REPO_OWNER%';" ^
  "$repo = '%REPO_NAME%';" ^
  "$current = '%CURRENT_VERSION%';" ^
  "$statusFile = '%STATUS_FILE%';" ^
  "$latestFile = '%LATEST_FILE%';" ^
  "$urlFile = '%URL_FILE%';" ^
  "function Norm([string]$v){ if([string]::IsNullOrWhiteSpace($v)){ return '0.0.0' }; if($v.StartsWith('v') -or $v.StartsWith('V')){ return $v.Substring(1) }; return $v }" ^
  "function Compare-Version([string]$a,[string]$b){ $aa=(Norm $a).Split('.'); $bb=(Norm $b).Split('.'); $len=[Math]::Max($aa.Length,$bb.Length); for($i=0;$i -lt $len;$i++){ $av=if($i -lt $aa.Length){ [int]$aa[$i] } else { 0 }; $bv=if($i -lt $bb.Length){ [int]$bb[$i] } else { 0 }; if($av -lt $bv){ return -1 }; if($av -gt $bv){ return 1 } }; return 0 }" ^
  "$url = ('https://api.github.com/repos/{0}/{1}/releases/latest' -f $owner,$repo);" ^
  "$release = Invoke-RestMethod -Headers @{Accept='application/vnd.github+json'; 'X-GitHub-Api-Version'='2022-11-28'} -Uri $url;" ^
  "$latest = Norm $release.tag_name;" ^
  "if((Compare-Version $current $latest) -ge 0){ Set-Content -LiteralPath $statusFile -Value 'UPTODATE' -Encoding ASCII; exit 0 }" ^
  "$asset = $release.assets | Where-Object { $_.name -eq 'KATS-Tools-windows-update.zip' } | Select-Object -First 1;" ^
  "if(-not $asset){ throw 'Release asset KATS-Tools-windows-update.zip not found.' }" ^
  "Set-Content -LiteralPath $statusFile -Value 'UPDATE' -Encoding ASCII;" ^
  "Set-Content -LiteralPath $latestFile -Value $latest -Encoding ASCII;" ^
  "Set-Content -LiteralPath $urlFile -Value $asset.browser_download_url -Encoding ASCII" >> "%LOG_FILE%" 2>&1
if errorlevel 1 exit /b 1

set "STATUS="
set "LATEST_VERSION="
set "DOWNLOAD_URL="

if exist "%STATUS_FILE%" set /P STATUS=<"%STATUS_FILE%"
if exist "%LATEST_FILE%" set /P LATEST_VERSION=<"%LATEST_FILE%"
if exist "%URL_FILE%" set /P DOWNLOAD_URL=<"%URL_FILE%"

exit /b 0

:show_info
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('%~1','KATS-Tools','OK','Information') | Out-Null" >NUL 2>&1
exit /b 0

:show_error
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('%~1','KATS-Tools','OK','Error') | Out-Null" >NUL 2>&1
exit /b 0

:fail
echo Update failed. >> "%LOG_FILE%"
call :show_error "Uppdateringen misslyckades. Se loggen: %LOG_FILE%"
goto :cleanup_fail

:cleanup_ok
del /Q "%STATUS_FILE%" >NUL 2>&1
del /Q "%LATEST_FILE%" >NUL 2>&1
del /Q "%URL_FILE%" >NUL 2>&1
del /Q "%ZIP_FILE%" >NUL 2>&1
rmdir /S /Q "%UNPACK_DIR%" >NUL 2>&1
rmdir /S /Q "%TEMP_DIR%" >NUL 2>&1
exit /b 0

:cleanup_fail
exit /b 1
