@echo off
setlocal EnableExtensions EnableDelayedExpansion
title ExcelTrainingMonitor Builder
color 0A

set "PROJECT=src\ExcelTrainingMonitor.csproj"
set "APP_NAME=ExcelTrainingMonitor"
set "ARTIFACTS=artifacts"
set "RELEASE_ROOT=%ARTIFACTS%\release"
set "SOURCE_ROOT=%ARTIFACTS%\source"
set "TEMP_ROOT=%ARTIFACTS%\temp"
set "PUBLISH_DIR=%TEMP_ROOT%\publish"
set "STAGE_DIR=%TEMP_ROOT%\source-stage"

if not exist version.txt (
    echo Missing version.txt.
    echo Create version.txt in the repo root with a value such as 1.0.0.
    goto failed
)

for /f "usebackq tokens=*" %%i in ("version.txt") do set "VERSION=%%i"
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "BUILDSTAMP=%%i"

set "RELEASE_NAME=%APP_NAME%_%VERSION%_%BUILDSTAMP%"
set "RELEASE_DIR=%RELEASE_ROOT%\%RELEASE_NAME%"
set "RELEASE_ZIP=%RELEASE_ROOT%\%RELEASE_NAME%.zip"
set "SOURCE_ZIP=%SOURCE_ROOT%\%APP_NAME%_source_%VERSION%_%BUILDSTAMP%.zip"
set "OLD_VERSION=%VERSION%"

echo.
echo ==========================================
echo %APP_NAME% Build System
echo ==========================================
echo.

call :check_dependencies
if errorlevel 1 goto failed

call :bump_version
if errorlevel 1 goto failed

for /f "usebackq tokens=*" %%i in ("version.txt") do set "VERSION=%%i"
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "BUILDSTAMP=%%i"

set "RELEASE_NAME=%APP_NAME%_%VERSION%_%BUILDSTAMP%"
set "RELEASE_DIR=%RELEASE_ROOT%\%RELEASE_NAME%"
set "RELEASE_ZIP=%RELEASE_ROOT%\%RELEASE_NAME%.zip"
set "SOURCE_ZIP=%SOURCE_ROOT%\%APP_NAME%_source_%VERSION%_%BUILDSTAMP%.zip"

echo.
echo [1/6] Cleaning artifact folders...
if exist "%ARTIFACTS%" rd /s /q "%ARTIFACTS%"
mkdir "%PUBLISH_DIR%" || goto failed
mkdir "%RELEASE_ROOT%" || goto failed
mkdir "%SOURCE_ROOT%" || goto failed

echo.
echo [2/6] Restoring NuGet packages...
dotnet restore "%PROJECT%"
if errorlevel 1 goto failed

echo.
echo [3/6] Publishing self-contained Windows x64 release...
dotnet publish "%PROJECT%" ^
    -c Release ^
    -r win-x64 ^
    --no-restore ^
    --self-contained true ^
    -p:PublishSingleFile=true ^
    -p:IncludeNativeLibrariesForSelfExtract=true ^
    -p:PublishTrimmed=false ^
    -p:Version=%VERSION% ^
    -o "%PUBLISH_DIR%"
if errorlevel 1 goto failed

echo.
echo [4/6] Creating release package...
mkdir "%RELEASE_DIR%" || goto failed
xcopy "%PUBLISH_DIR%" "%RELEASE_DIR%\" /E /I /Y >nul
copy version.txt "%RELEASE_DIR%\version.txt" >nul
if exist CHANGELOG.txt copy CHANGELOG.txt "%RELEASE_DIR%\CHANGELOG.txt" >nul
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path '%RELEASE_DIR%\*' -DestinationPath '%RELEASE_ZIP%' -Force"
if errorlevel 1 goto failed

call :build_installer

echo.
echo [5/6] Creating source pipeline package...
if exist "%STAGE_DIR%" rd /s /q "%STAGE_DIR%"
mkdir "%STAGE_DIR%" || goto failed
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference = 'Stop';" ^
    "$exclude = @('.git','.vs','bin','obj','artifacts','dist','publish','build','theme-analysis','tools');" ^
    "$root = (Get-Location).Path;" ^
    "$stage = Join-Path $root '%STAGE_DIR%';" ^
    "Get-ChildItem -LiteralPath $root -Recurse -File -Force | Where-Object { $relative = $_.FullName.Substring($root.Length).TrimStart('\'); $parts = $relative -split '\\'; -not ($parts | Where-Object { $exclude -contains $_ }) } | ForEach-Object { $relative = $_.FullName.Substring($root.Length).TrimStart('\'); $destination = Join-Path $stage $relative; New-Item -ItemType Directory -Force -Path (Split-Path $destination) | Out-Null; Copy-Item -LiteralPath $_.FullName -Destination $destination -Force };" ^
    "Compress-Archive -Path (Join-Path $stage '*') -DestinationPath (Join-Path $root '%SOURCE_ZIP%') -Force"
if errorlevel 1 goto failed

echo.
echo [6/6] Cleaning temporary publish files...
if exist "%TEMP_ROOT%" rd /s /q "%TEMP_ROOT%"

echo.
echo ==========================================
echo BUILD SUCCESS
echo ==========================================
echo.
echo Release folder:
echo %RELEASE_DIR%
echo.
echo Version:
echo %OLD_VERSION% -^> %VERSION%
echo.
echo Release ZIP:
echo %RELEASE_ZIP%
echo.
echo Source pipeline ZIP:
echo %SOURCE_ZIP%
echo.
echo Upload the release ZIP to a GitHub Release.
echo Upload source changes to GitHub from the repo, or use the source pipeline ZIP if you need a clean source package.
echo.
pause
exit /b 0

:bump_version
echo.
echo Auto-incrementing patch version...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference = 'Stop';" ^
    "$versionPath = 'version.txt';" ^
    "$old = (Get-Content $versionPath -Raw).Trim();" ^
    "if ($old -notmatch '^\d+\.\d+\.\d+$') { throw 'version.txt must use Major.Minor.Patch, for example 1.0.0' }" ^
    "$parts = $old.Split('.') | ForEach-Object { [int]$_ };" ^
    "$new = '{0}.{1}.{2}' -f $parts[0], $parts[1], ($parts[2] + 1);" ^
    "Set-Content -Path $versionPath -Value $new -NoNewline;" ^
    "$project = '%PROJECT%';" ^
    "$xml = Get-Content $project -Raw;" ^
    "$assembly = '{0}.{1}.{2}.{2}' -f $parts[0], $parts[1], ($parts[2] + 1);" ^
    "$xml = [regex]::Replace($xml, '<AssemblyVersion>.*?</AssemblyVersion>', '<AssemblyVersion>' + $assembly + '</AssemblyVersion>');" ^
    "$xml = [regex]::Replace($xml, '<FileVersion>.*?</FileVersion>', '<FileVersion>' + $assembly + '</FileVersion>');" ^
    "if ($xml -match '<Version>.*?</Version>') { $xml = [regex]::Replace($xml, '<Version>.*?</Version>', '<Version>' + $new + '</Version>') } else { $xml = [regex]::Replace($xml, '(<FileVersion>.*?</FileVersion>)', '$1' + [Environment]::NewLine + '    <Version>' + $new + '</Version>') };" ^
    "Set-Content -Path $project -Value $xml -NoNewline;" ^
    "Write-Host ('[OK] Version ' + $old + ' -> ' + $new)"
if errorlevel 1 exit /b 1
exit /b 0

:build_installer
echo.
echo Optional installer package...
where ISCC.exe >nul 2>nul
if errorlevel 1 (
    echo [SKIP] Inno Setup Compiler not found.
    echo        Portable release ZIP was created successfully.
    echo        To build an installer, install Inno Setup:
    echo        https://jrsoftware.org/isdl.php
    exit /b 0
)

mkdir "%ARTIFACTS%\installer" 2>nul
ISCC.exe "installer\installer.iss" /DMyAppVersion=%VERSION% /DPublishDir="%CD%\%PUBLISH_DIR%" /DInstallerOutDir="%CD%\%ARTIFACTS%\installer"
if errorlevel 1 (
    echo [WARN] Installer build failed. Portable release ZIP is still available.
    exit /b 0
)

echo [OK] Installer created in %ARTIFACTS%\installer
exit /b 0

:check_dependencies
echo Checking build dependencies...
set "MISSING=0"

where dotnet >nul 2>nul
if errorlevel 1 (
    echo [MISSING] .NET SDK 8.0 or newer
    echo           Download: https://dotnet.microsoft.com/download/dotnet/8.0
    set "MISSING=1"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$sdks = dotnet --list-sdks | ForEach-Object { if ($_ -match '^(\d+)\.') { [int]$matches[1] } }; if (($sdks | Where-Object { $_ -ge 8 } | Select-Object -First 1) -ne $null) { exit 0 } else { exit 1 }" >nul 2>nul
    if errorlevel 1 (
        echo [MISSING] .NET SDK 8.0 or newer
        echo           dotnet exists, but no compatible SDK was found.
        echo           Download: https://dotnet.microsoft.com/download/dotnet/8.0
        set "MISSING=1"
    ) else (
        echo [OK] .NET SDK 8.0 or newer
    )
)

where powershell >nul 2>nul
if errorlevel 1 (
    echo [MISSING] Windows PowerShell
    echo           Required for timestamping and ZIP creation.
    echo           Included with Windows, or install PowerShell from:
    echo           https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows
    set "MISSING=1"
) else (
    echo [OK] Windows PowerShell
)

if not exist "%PROJECT%" (
    echo [MISSING] Project file: %PROJECT%
    echo           Run this script from the repository root.
    set "MISSING=1"
) else (
    echo [OK] Project file
)

where git >nul 2>nul
if errorlevel 1 (
    echo [INFO] Git was not found on PATH.
    echo        Git is not required to build locally, but it is recommended for GitHub release tags.
    echo        Download: https://git-scm.com/download/win
) else (
    echo [OK] Git
)

if "%MISSING%"=="1" (
    echo.
    echo One or more required build dependencies are missing.
    exit /b 1
)

echo [OK] Required build dependencies found.
exit /b 0

:failed
echo.
echo ==========================================
echo BUILD FAILED
echo ==========================================
echo.
pause
exit /b 1
