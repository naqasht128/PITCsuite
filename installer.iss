#define AppName "PITC Suite"
#define AppPublisher "TN Developers Co."
#define AppExeName "pitcsuite.exe"
#ifndef AppVersion
#define AppVersion "2.8.100"
#endif

[Setup]
AppId={{9C56E2F0-4C9D-428E-BB5A-9D8B6F3E1B92}}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=mailto:tahir.naqash4@gmail.com
AppSupportURL=https://wa.me/923339962474
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
DefaultDirName={localappdata}\Programs\PITCSuite
DisableDirPage=yes
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=PITCSuite_Setup
SetupIconFile=pitcsuite\assets\pitcsuite_icon.ico
UninstallDisplayIcon={app}\{#AppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
DisableWelcomePage=no
CloseApplications=force
RestartApplications=no
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppName} Setup
VersionInfoVersion={#AppVersion}
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVersion}
VersionInfoCopyright=Copyright (C) 2026 {#AppPublisher}

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; The main executable built by PyInstaller
Source: "dist\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent shellexec; WorkingDir: "{app}"
Filename: "{app}\{#AppExeName}"; Flags: nowait shellexec; Check: ShouldLaunchSilent; WorkingDir: "{app}"

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

function ShouldLaunchSilent: Boolean;
begin
  Result := WizardSilent or (ExpandConstant('{param:AUTOUPDATE|0}') = '1');
end;
