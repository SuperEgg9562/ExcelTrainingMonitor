@echo off
setlocal

cd /d %~dp0
cd ..

if not exist src mkdir src
if not exist build mkdir build
if not exist dist mkdir dist
if not exist installer mkdir installer

if not exist src\ExcelTrainingMonitor.csproj (
    echo Project structure initialized.
) else (
    echo Project already exists.
)

pause