; Inno Setup Script for Desktop Labeler (mddsklbl)
; Build with: iscc installers\DesktopLabeler.iss

#define MyAppName "Desktop Labeler"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "0x4D44 Software"
#define MyAppExeName "mddsklbl.exe"

[Setup]
AppId={{F7D9C1E7-AB15-4EFA-8E3F-7C6E9F6D9B21}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={pf}\0x4D44 Software\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableDirPage=no
DisableProgramGroupPage=yes
OutputBaseFilename=DesktopLabeler-{#MyAppVersion}-Setup
OutputDir=.
Compression=lzma
SolidCompression=yes
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\target\release\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[Tasks]
Name: "autostart"; Description: "Run at Windows startup"; Flags: checkedonce

[Registry]
; HKCU Run entry for current user. Removed on uninstall.
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "DesktopLabeler"; ValueData: """{app}\{#MyAppExeName}"""; Flags: uninsdeletevalue; Tasks: autostart
; Clean up legacy key value if present
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: none; ValueName: "DesktopNameManager"; Flags: deletevalue uninsdeletevalue

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
