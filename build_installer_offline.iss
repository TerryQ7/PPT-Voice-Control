[Setup]
AppId={{4FCA2D8C-D9E3-4E2E-9B5E-6A90D3EF2A11}
AppName=PPT Voice Control Offline
AppVersion=2026.04.10
AppPublisher=TerryQ7
DefaultDirName={localappdata}\PPT Voice Control Offline
DefaultGroupName=PPT Voice Control Offline
DisableProgramGroupPage=yes
OutputDir=release
OutputBaseFilename=PPT-Voice-Control-Offline-Installer
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
SetupLogging=yes

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Files]
Source: "installer_staging\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\PPT Voice Control Offline"; Filename: "{app}\PPT-Voice-Control.exe"
Name: "{autodesktop}\PPT Voice Control Offline"; Filename: "{app}\PPT-Voice-Control.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\PPT-Voice-Control.exe"; Description: "Launch PPT Voice Control Offline"; Flags: nowait postinstall skipifsilent
