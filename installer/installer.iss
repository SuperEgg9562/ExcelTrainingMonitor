#define MyAppVersion "1.0.0"
#ifndef PublishDir
#define PublishDir "..\artifacts\temp\publish"
#endif
#ifndef InstallerOutDir
#define InstallerOutDir "..\artifacts\installer"
#endif

[Setup]
AppName=ExcelTrainingMonitor
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\ExcelTrainingMonitor
DefaultGroupName=ExcelTrainingMonitor
OutputDir={#InstallerOutDir}
OutputBaseFilename=ExcelTrainingMonitorSetup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName=ExcelTrainingMonitor
PrivilegesRequired=lowest

[Files]
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\ExcelTrainingMonitor"; Filename: "{app}\ExcelTrainingMonitor.exe"
Name: "{group}\Uninstall ExcelTrainingMonitor"; Filename: "{uninstallexe}"
Name: "{commondesktop}\ExcelTrainingMonitor"; Filename: "{app}\ExcelTrainingMonitor.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Run]
Filename: "{app}\ExcelTrainingMonitor.exe"; Description: "Launch ExcelTrainingMonitor"; Flags: nowait postinstall skipifsilent
