[Setup]
AppName=ExcelTrainingMonitor
AppVersion=1.0.0
DefaultDirName={pf}\ExcelTrainingMonitor
DefaultGroupName=ExcelTrainingMonitor
OutputDir=output
OutputBaseFilename=ExcelTrainingMonitorSetup
Compression=lzma
SolidCompression=yes

[Files]
Source: "..\dist\*\ExcelTrainingMonitor.exe"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\ExcelTrainingMonitor"; Filename: "{app}\ExcelTrainingMonitor.exe"
Name: "{commondesktop}\ExcelTrainingMonitor"; Filename: "{app}\ExcelTrainingMonitor.exe"