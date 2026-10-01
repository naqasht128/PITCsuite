[Setup]
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
AppName=PITC Suite
AppVersion=2.8.89
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
CloseApplicationsFilter=pitcsuite.exe
RestartApplications=no

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"

[Files]
; The main executable built by PyInstaller
Source: "dist\pitcsuite.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\PITC Suite"; Filename: "{app}\pitcsuite.exe"
Name: "{autodesktop}\PITC Suite"; Filename: "{app}\pitcsuite.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\pitcsuite.exe"; Description: "Launch PITC Suite"; Flags: nowait postinstall skipifsilent shellexec; WorkingDir: "{app}"
Filename: "{app}\pitcsuite.exe"; Flags: nowait shellexec; Check: IsAutoUpdate; WorkingDir: "{app}"

[Code]
function SetEnvironmentVariable(lpName, lpValue: string): BOOL;
external 'SetEnvironmentVariableW@kernel32.dll stdcall';

function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  // Clear any inherited PyInstaller variables so the relaunched app unpacks cleanly
  SetEnvironmentVariable('_MEIPASS2', '');
  SetEnvironmentVariable('_MEIPASS', '');
  SetEnvironmentVariable('PYTHONHOME', '');
  SetEnvironmentVariable('PYTHONPATH', '');

  // Terminate any running or lingering instances of pitcsuite.exe (without /T so installer itself is never terminated)
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM pitcsuite.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(400);
  Result := True;
end;

function IsAutoUpdate: Boolean;
begin
  Result := ExpandConstant('{param:AUTOUPDATE|0}') = '1';
end;



