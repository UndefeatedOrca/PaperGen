#define MyAppName "PaperGen"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "PaperGen"
#define MyAppExeName "PaperGen.exe"

[Setup]
AppId={{B9A3CBE3-2BAA-4BA8-9BEA-4B9D7A3DAA42}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
PrivilegesRequired=lowest
OutputDir=..\dist\installer
OutputBaseFilename=PaperGenSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\dist\PaperGen\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion
Source: "..\templates\*"; DestDir: "{userappdata}\PaperGen\templates"; Flags: recursesubdirs ignoreversion onlyifdoesntexist
Source: "..\profile.example.json"; DestDir: "{userappdata}\PaperGen"; Flags: onlyifdoesntexist ignoreversion
Source: "..\classes.example.json"; DestDir: "{userappdata}\PaperGen"; Flags: onlyifdoesntexist ignoreversion
Source: "..\config.example.json"; DestDir: "{userappdata}\PaperGen"; Flags: onlyifdoesntexist ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}\Configuration"; Filename: "{app}\{#MyAppExeName}"; Parameters: "--config"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Parameters: "--config"; WorkingDir: "{app}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Registry]
Root: HKCU; Subkey: "Software\Classes\Directory\Background\shell\PaperGen"; ValueType: string; ValueName: ""; ValueData: "New Paper from Template"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\Directory\Background\shell\PaperGen"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#MyAppExeName}"
Root: HKCU; Subkey: "Software\Classes\Directory\Background\shell\PaperGen\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%V"""; Flags: uninsdeletekey

[Run]
Filename: "{app}\{#MyAppExeName}"; Parameters: "--config"; Description: "Open PaperGen configuration"; Flags: postinstall nowait skipifsilent unchecked
