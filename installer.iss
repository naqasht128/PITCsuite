[Setup]
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
AppName=PITC Suite
AppVersion=2.8.86
AppPublisher=Tahir Naqash
DefaultDirName={localappdata}\Programs\PITCSuite
DefaultGroupName=PITC Suite
OutputDir=dist
OutputBaseFilename=PITCSuite_Setup
SetupIconFile=pitcsuite\assets\pitcsuite_icon.ico
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
DisableWelcomePage=no

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"

[Files]
; The main executable built by PyInstaller
Source: "dist\pitcsuite.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\PITC Suite"; Filename: "{app}\pitcsuite.exe"
Name: "{autodesktop}\PITC Suite"; Filename: "{app}\pitcsuite.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\pitcsuite.exe"; Description: "Launch PITC Suite"; Flags: nowait postinstall skipifsilent shellexec
Filename: "{app}\pitcsuite.exe"; Flags: nowait shellexec; Check: IsAutoUpdate

[Code]
function IsAutoUpdate: Boolean;
begin
  Result := ExpandConstant('{param:AUTOUPDATE|0}') = '1';
end;


