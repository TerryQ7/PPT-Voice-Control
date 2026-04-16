; PPT Voice Control — bilingual Inno Setup template.
;
; Compile from CI / local with one of:
;   ISCC.exe /DAppLang=zh /DMyAppVersion=1.0.0 installer.iss
;   ISCC.exe /DAppLang=en /DMyAppVersion=1.0.0 installer.iss
;
; The Chinese build produces a Chinese-named installer
; (PPT语音控制助手-安装程序-v<ver>.exe) and a Chinese setup wizard.
; The English build produces PPT-Voice-Control-Setup-v<ver>.exe with the
; default English wizard.
;
; Both editions can be installed side-by-side: each variant has its own
; AppId, install folder, and Start-menu entry.

#ifndef AppLang
  #define AppLang "en"
#endif
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif

#if AppLang == "zh"
  #define MyAppName        "PPT语音控制助手"
  #define MyAppPublisher   "United Imaging"
  #define MyExeName        "PPT语音控制助手.exe"
  #define MyAppId          "{B5DE3F4A-7C21-4A5F-9D8E-1F2A3B4C5D6E}"
  #define MyOutputName     "PPT语音控制助手-安装程序-v" + MyAppVersion
  #define MyTaskDesc       "在桌面创建快捷方式"
  #define MyTaskGroup      "附加任务："
  #define MyRunDesc        "立即启动 PPT 语音控制助手"
#else
  #define MyAppName        "PPT Voice Control"
  #define MyAppPublisher   "United Imaging"
  #define MyExeName        "PPT-Voice-Control.exe"
  #define MyAppId          "{4FCA2D8C-D9E3-4E2E-9B5E-6A90D3EF2A11}"
  #define MyOutputName     "PPT-Voice-Control-Setup-v" + MyAppVersion
  #define MyTaskDesc       "Create a desktop shortcut"
  #define MyTaskGroup      "Additional shortcuts:"
  #define MyRunDesc        "Launch " + MyAppName
#endif

[Setup]
AppId={{#MyAppId}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=release
OutputBaseFilename={#MyOutputName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
SetupLogging=yes
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\{#MyExeName}

[Languages]
#if AppLang == "zh"
Name: "zhs"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
#else
Name: "en"; MessagesFile: "compiler:Default.isl"
#endif

[Tasks]
Name: "desktopicon"; Description: "{#MyTaskDesc}"; GroupDescription: "{#MyTaskGroup}"

[Files]
Source: "installer_staging\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyExeName}"; Description: "{#MyRunDesc}"; Flags: nowait postinstall skipifsilent
