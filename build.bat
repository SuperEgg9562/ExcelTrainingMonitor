@echo off
title ExcelTrainingMonitor Builder
color 0A

set PROJECT=src\ExcelTrainingMonitor.csproj

for /f %%i in (version.txt) do set VERSION=%%i

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set BUILDSTAMP=%%i

set OUTPUT=dist\ExcelTrainingMonitor_%VERSION%_%BUILDSTAMP%

echo.
echo ==========================================
echo ExcelTrainingMonitor Build System
echo ==========================================
echo.

echo [1/5] Cleaning old output folders...

if exist publish rd /s /q publish

mkdir publish

echo.
echo [2/5] Restoring packages...

dotnet restore %PROJECT%

if errorlevel 1 goto failed

echo.
echo [3/5] Publishing Single File EXE...

dotnet publish %PROJECT% ^
-c Release ^
-r win-x64 ^
--no-restore ^
--self-contained true ^
-p:PublishSingleFile=true ^
-p:IncludeNativeLibrariesForSelfExtract=true ^
-p:PublishTrimmed=false ^
-p:Version=%VERSION% ^
-o publish


if errorlevel 1 goto failed

echo.
echo [4/5] Creating Release Folder...

mkdir "%OUTPUT%"

xcopy publish "%OUTPUT%\" /E /I /Y

copy version.txt "%OUTPUT%\version.txt"
if exist CHANGELOG.txt (
    copy CHANGELOG.txt "%OUTPUT%\CHANGELOG.txt"
)

echo.
echo [5/5] Creating ZIP Package...

powershell -NoProfile -Command ^
"Compress-Archive -Path '%OUTPUT%\*' -DestinationPath '%OUTPUT%.zip' -Force"

echo.
echo ==========================================
echo BUILD SUCCESS
echo ==========================================
echo.
echo Output:
echo %OUTPUT%
echo.
pause
exit

:failed
echo.
echo ==========================================
echo BUILD FAILED
echo ==========================================
echo.
pause
exit /b 1
