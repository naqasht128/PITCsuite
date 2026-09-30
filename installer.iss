[Setup]
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
AppName=PITC Suite
AppVersion=2.8.85
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
CloseApplications=force
CloseApplicationsFilter=*.exe
RestartApplications=no

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"

[Files]
; The main executable built by PyInstaller
Source: "dist\pitcsuite.exe"; DestDir: "{app}"; Flags: ignoreversion restartreplace

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

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ResultCode: Integer;
begin
  // Force close any running or lingering instances of pitcsuite.exe before attempting file replacement
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM pitcsuite.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  // Give Windows file system half a second to release the file lock
  Sleep(600);
  Result := '';
end;
