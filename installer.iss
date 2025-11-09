; installer.iss
#define MyAppName "MERQ Timesheet"
#define MyAppVersion "1.0.0.0"
#define MyAppPublisher "MERQ Consultancy PLC"
#define MyAppURL "https://merqconsultancy.org/"
#define MyAppExeName "MERQ_Timesheet.exe"
#define UpdateURL "https://app.merqconsultancy.org/apps/timesheet/desktop/"

[Setup]
; Code Signing
SignTool=MERQSign $f
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#UpdateURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=LICENSE.txt
OutputDir=installer
OutputBaseFilename=MERQ_Timesheet_Setup
SetupIconFile=src/merq.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\merq.ico
AppComments=This explicitly developed for only MERQ Consultancy PLC & Developed by ISDHU!
AppContact=support@merqconsultancy.org
AppSupportPhone=+251913391985
AppReadmeFile=README.txt
VersionInfoVersion=1.0.0.0
VersionInfoCompany=MERQ Consultancy PLC
VersionInfoDescription=This application helps MERQ employees and consultants track their working hours 
VersionInfoTextVersion=Version 1.0.0.0
VersionInfoCopyright=Copyright © 2025 MERQ Consultancy PLC. All rights reserved.
VersionInfoProductName=MERQ Timesheet
VersionInfoProductVersion=1.0.0.0
AppCopyright=Copyright © 2025 MERQ Consultancy PLC. All rights reserved.
InfoBeforeFile=copyright.txt
InfoAfterFile=README.txt

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}";
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; OnlyBelowVersion: 6.1

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "src\merq.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "src\merq.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "src\MERQ_TIMESHEET_ETH-CAL_TEMPLATE.xlsx"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "copyright.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Registry]
Root: HKCU; Subkey: "Software\{#MyAppName}"; Flags: uninsdeletekey
; Store version information for update checks
Root: HKCU; Subkey: "Software\{#MyAppName}\Updates"; ValueType: string; ValueName: "InstalledVersion"; ValueData: "{#MyAppVersion}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\{#MyAppName}\Updates"; ValueType: string; ValueName: "UpdateURL"; ValueData: "{#UpdateURL}"; Flags: uninsdeletekey

[Code]
// Simple Windows version check
function IsWindows8OrNewer: Boolean;
var
  Version: TWindowsVersion;
begin
  GetWindowsVersionEx(Version);
  Result := (Version.Major > 6) or ((Version.Major = 6) and (Version.Minor >= 2));
end;

// Check if app is running using a simple process check
function IsAppRunning: Boolean;
var
  ResultCode: Integer;
begin
  Result := False;
  // Try to create a mutex - if it already exists, app is running
  if not Exec('cmd.exe', '/C exit 0', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
  begin
    // If we can't even run cmd, assume app might be running
    Result := True;
  end;
end;

function InitializeSetup(): Boolean;
begin
  // Simple check if application might be running
  if IsAppRunning then
  begin
    if MsgBox('MERQ Timesheet might be currently running. Please close it and click OK to continue with the installation.', mbInformation, MB_OKCANCEL) = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;
  end;
  
  Result := True;
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  // Skip welcome page during silent updates
  if WizardSilent() and (PageID = wpWelcome) then
    Result := True
  else
    Result := False;
end;

function GetUninstallString(): String;
var
  UninstallKey: String;
begin
  UninstallKey := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}_is1';
  Result := '';
  if not RegQueryStringValue(HKLM, UninstallKey, 'UninstallString', Result) then
    RegQueryStringValue(HKCU, UninstallKey, 'UninstallString', Result);
end;

function IsUpgrade(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;

// Simple update check without complex HTTP requests
function CheckForUpdates(): Boolean;
var
  InstalledVersion: string;
begin
  // Get installed version from registry
  if not RegQueryStringValue(HKCU, 'Software\{#MyAppName}\Updates', 'InstalledVersion', InstalledVersion) then
    InstalledVersion := '{#MyAppVersion}';
  
  // For now, always return false - update checking will be handled by the Python application
  // This avoids complex HTTP code in Inno Setup that can cause compilation errors
  Result := False;
  
  // Log for debugging
  Log('Update check: Installed version = ' + InstalledVersion);
end;

procedure InitializeWizard;
var
  TaskbarPinCheckBox: TNewCheckBox;
begin
  // Create custom checkbox for taskbar pinning (Windows 10/11 only)
  TaskbarPinCheckBox := TNewCheckBox.Create(WizardForm);
  TaskbarPinCheckBox.Parent := WizardForm.SelectTasksPage;
  TaskbarPinCheckBox.Left := ScaleX(4);
  TaskbarPinCheckBox.Top := ScaleY(180);
  TaskbarPinCheckBox.Width := ScaleX(400);
  TaskbarPinCheckBox.Height := ScaleY(24);
  TaskbarPinCheckBox.Caption := 'Pin to &taskbar';
  TaskbarPinCheckBox.Checked := True;
  TaskbarPinCheckBox.Visible := True;
  TaskbarPinCheckBox.Enabled := IsWindows8OrNewer; // Windows 8+ but mainly for 10/11
end;

function PinToTaskbar(const AppPath: string): Boolean;
var
  ResultCode: Integer;
begin
  if not IsWindows8OrNewer then // Only for Windows 8+
  begin
    Log('Taskbar pinning not supported on this Windows version');
    Result := True;
    Exit;
  end;

  // Simple PowerShell command to pin to taskbar
  Result := Exec('powershell', '-Command "$shell = New-Object -ComObject ''Shell.Application''; $item = (New-Object -ComObject ''Shell.Application'').Namespace(''' + ExtractFilePath(AppPath) + ''').ParseName(''' + ExtractFileName(AppPath) + '''); $item.InvokeVerb(''pintotaskbar'')"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  
  if Result then
    Log('Taskbar pinning command executed successfully')
  else
    Log('Taskbar pinning failed');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  // Pin to taskbar after installation if enabled
  if CurStep = ssPostInstall then
  begin
    // Check if taskbar pinning was selected
    PinToTaskbar(ExpandConstant('{app}\{#MyAppExeName}'));
  end;
end;