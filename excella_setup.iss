#define MyAppName "Excella"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "frenzywall"
#define MyAppURL "https://github.com/frenzywall/Excella"
#define MyAppExeName "Excella.exe"
#define MyAppDescription "Excella"
#define MyAppCopyright "Copyright © 2025 Sreeram. All rights reserved."

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
AppId={{E5A7B8C9-D4F6-4A3B-9E2C-1F8G7H6I5J4K}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases
AppCopyright={#MyAppCopyright}
AppComments={#MyAppDescription}
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppDescription}
VersionInfoCopyright={#MyAppCopyright}
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer
OutputBaseFilename=Excella_Setup_v{#MyAppVersion}
SetupIconFile=icon.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
; Minimum Windows version (Windows 10)
MinVersion=10.0
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Privileges
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
; Code signing (uncomment and configure if you have a certificate)
;SignTool=signtool
;SignedUninstaller=yes

; Prevent multiple instances during installation
AppMutex=ExcellaAppMutex
; Close applications before installation
CloseApplications=yes
RestartApplications=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1
Name: "associate"; Description: "Associate with .xlsx and .xls files"; GroupDescription: "File associations:"
Name: "associatecsv"; Description: "Associate with .csv files"; GroupDescription: "File associations:"; Flags: unchecked
Name: "addtopath"; Description: "Add to system PATH (for command line usage)"; GroupDescription: "System integration:"
Name: "startup"; Description: "Add to Windows startup"; GroupDescription: "System integration:"; Flags: unchecked
Name: "contextmenu"; Description: "Add ""Compare with Excella"" to context menu"; GroupDescription: "System integration:"

[Files]
Source: "dist\Excella\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\Excella\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "requirements.txt"; DestDir: "{app}"; Flags: ignoreversion; Check: FileExists(ExpandConstant('{src}\requirements.txt'))
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion; Check: FileExists(ExpandConstant('{src}\icon.ico'))



[Dirs]
Name: "{userappdata}\{#MyAppName}"; Flags: uninsneveruninstall
Name: "{userappdata}\{#MyAppName}\logs"; Flags: uninsneveruninstall
Name: "{userappdata}\{#MyAppName}\temp"; Flags: uninsneveruninstall
Name: "{userappdata}\{#MyAppName}\exports"; Flags: uninsneveruninstall
Name: "{userappdata}\{#MyAppName}\settings"; Flags: uninsneveruninstall
Name: "{userappdata}\{#MyAppName}\profiles"; Flags: uninsneveruninstall

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Comment: "{#MyAppDescription}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{group}\{#MyAppName} Documentation"; Filename: "{app}\README.txt"; IconFilename: "{app}\icon.ico"; Check: FileExists(ExpandConstant('{app}\README.txt'))
Name: "{group}\User Manual"; Filename: "{app}\docs"; IconFilename: "{app}\icon.ico"; Check: DirExists(ExpandConstant('{app}\docs'))

Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; Comment: "{#MyAppDescription}"
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon; Comment: "{#MyAppDescription}"

[Registry]
; File associations for XLSX
Root: HKA; Subkey: "Software\Classes\.xlsx\OpenWithProgids"; ValueType: string; ValueName: "{#MyAppName}.xlsx"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associate
Root: HKA; Subkey: "Software\Classes\{#MyAppName}.xlsx"; ValueType: string; ValueName: ""; ValueData: "{#MyAppName} Excel File"; Flags: uninsdeletekey; Tasks: associate
Root: HKA; Subkey: "Software\Classes\{#MyAppName}.xlsx\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#MyAppExeName},0"; Tasks: associate
Root: HKA; Subkey: "Software\Classes\{#MyAppName}.xlsx\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: associate

; File associations for XLS
Root: HKA; Subkey: "Software\Classes\.xls\OpenWithProgids"; ValueType: string; ValueName: "{#MyAppName}.xls"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associate
Root: HKA; Subkey: "Software\Classes\{#MyAppName}.xls"; ValueType: string; ValueName: ""; ValueData: "{#MyAppName} Excel File"; Flags: uninsdeletekey; Tasks: associate
Root: HKA; Subkey: "Software\Classes\{#MyAppName}.xls\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#MyAppExeName},0"; Tasks: associate
Root: HKA; Subkey: "Software\Classes\{#MyAppName}.xls\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: associate

; File associations for CSV (optional)
Root: HKA; Subkey: "Software\Classes\.csv\OpenWithProgids"; ValueType: string; ValueName: "{#MyAppName}.csv"; ValueData: ""; Flags: uninsdeletevalue; Tasks: associatecsv
Root: HKA; Subkey: "Software\Classes\{#MyAppName}.csv"; ValueType: string; ValueName: ""; ValueData: "{#MyAppName} CSV File"; Flags: uninsdeletekey; Tasks: associatecsv
Root: HKA; Subkey: "Software\Classes\{#MyAppName}.csv\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#MyAppExeName},0"; Tasks: associatecsv
Root: HKA; Subkey: "Software\Classes\{#MyAppName}.csv\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: associatecsv

; Context menu integration
Root: HKA; Subkey: "Software\Classes\*\shell\ExcellaCompare"; ValueType: string; ValueName: ""; ValueData: "Compare with Excella"; Tasks: contextmenu
Root: HKA; Subkey: "Software\Classes\*\shell\ExcellaCompare"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#MyAppExeName},0"; Tasks: contextmenu
Root: HKA; Subkey: "Software\Classes\*\shell\ExcellaCompare\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: contextmenu; Flags: uninsdeletekey

; Application settings
Root: HKCU; Subkey: "Software\{#MyAppName}"; ValueType: string; ValueName: "Version"; ValueData: "{#MyAppVersion}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\{#MyAppName}"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\{#MyAppName}"; ValueType: string; ValueName: "DataPath"; ValueData: "{userappdata}\{#MyAppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\{#MyAppName}"; ValueType: dword; ValueName: "FirstRun"; ValueData: 1; Flags: uninsdeletekey

; Add to PATH
Root: HKCU; Subkey: "Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{app}"; Tasks: addtopath; Check: NeedsAddPath('{app}')

; Startup entry
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "{#MyAppName}"; ValueData: """{app}\{#MyAppExeName}"" --minimized"; Tasks: startup; Flags: uninsdeletevalue

; Add to Windows Defender exclusions (if needed)
; Root: HKLM; Subkey: "SOFTWARE\Microsoft\Windows Defender\Exclusions\Paths"; ValueType: dword; ValueName: "{app}"; ValueData: 0; Flags: uninsdeletevalue; MinVersion: 10.0

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
Filename: "{app}\README.txt"; Description: "View README file"; Flags: postinstall skipifsilent shellexec unchecked; Check: FileExists(ExpandConstant('{app}\README.txt'))
Filename: "https://github.com/frenzywall/Excella"; Description: "Visit project website"; Flags: postinstall skipifsilent shellexec unchecked
Filename: "https://github.com/frenzywall/Excella/releases"; Description: "Check for updates"; Flags: postinstall skipifsilent shellexec unchecked

[UninstallRun]
Filename: "{cmd}"; Parameters: "/c taskkill /f /im {#MyAppExeName}"; Flags: runhidden; RunOnceId: "KillExcella"
Filename: "{cmd}"; Parameters: "/c reg delete ""HKCU\Software\Microsoft\Windows\CurrentVersion\Run"" /v ""{#MyAppName}"" /f"; Flags: runhidden; RunOnceId: "RemoveStartup"

[UninstallDelete]
Type: files; Name: "{userappdata}\{#MyAppName}\*.log"
Type: files; Name: "{userappdata}\{#MyAppName}\temp\*"
Type: files; Name: "{userappdata}\{#MyAppName}\cache\*"
Type: dirifempty; Name: "{userappdata}\{#MyAppName}\temp"
Type: dirifempty; Name: "{userappdata}\{#MyAppName}\cache"
Type: dirifempty; Name: "{userappdata}\{#MyAppName}\logs"
; Don't delete settings and exports folders - user data

[Code]
function NeedsAddPath(Param: string): boolean;
var
  OrigPath: string;
begin
  if not RegQueryStringValue(HKEY_CURRENT_USER, 'Environment', 'Path', OrigPath)
  then begin
    Result := True;
    exit;
  end;
  Result := Pos(';' + UpperCase(Param) + ';', ';' + UpperCase(OrigPath) + ';') = 0;
end;

function IsAppRunning(const FileName: string): Boolean;
var
  FSWbemLocator: Variant;
  FWMIService: Variant;
  FWbemObjectSet: Variant;
begin
  Result := false;
  try
    FSWbemLocator := CreateOleObject('WBEMScripting.SWBEMLocator');
    FWMIService := FSWbemLocator.ConnectServer('', 'root\CIMV2', '', '');
    FWbemObjectSet := FWMIService.ExecQuery(Format('SELECT Name FROM Win32_Process WHERE Name=''%s''', [FileName]));
    Result := (FWbemObjectSet.Count > 0);
  except
    Result := false;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then begin
    // Check if application is running
    if IsAppRunning('{#MyAppExeName}') then begin
      if MsgBox('Excella is currently running. Setup will attempt to close it automatically. Continue?', 
                mbConfirmation, MB_YESNO) = IDNO then begin
        Abort;
      end;
    end;
  end;
  
  if CurStep = ssPostInstall then begin
    // Clean up any previous installation remnants
    DelTree(ExpandConstant('{userappdata}\{#MyAppName}\cache'), True, True, True);
    
    // Create initial config if needed
    if not FileExists(ExpandConstant('{userappdata}\{#MyAppName}\settings\config.ini')) then begin
      SaveStringToFile(ExpandConstant('{userappdata}\{#MyAppName}\settings\config.ini'), 
                      '[Settings]' + #13#10 + 
                      'FirstRun=true' + #13#10 +
                      'Version=' + '{#MyAppVersion}' + #13#10 +
                      'AutoUpdate=true' + #13#10, False);
    end;
  end;
end;

function GetUninstallString(): String;
var
  sUnInstPath: String;
  sUnInstallString: String;
begin
  sUnInstPath := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}_is1');
  sUnInstallString := '';
  if not RegQueryStringValue(HKLM, sUnInstPath, 'UninstallString', sUnInstallString) then
    RegQueryStringValue(HKCU, sUnInstPath, 'UninstallString', sUnInstallString);
  Result := sUnInstallString;
end;

function IsUpgrade(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;

function UnInstallOldVersion(): Integer;
var
  sUnInstallString: String;
  iResultCode: Integer;
begin
  Result := 0;
  sUnInstallString := GetUninstallString();
  if sUnInstallString <> '' then begin
    sUnInstallString := RemoveQuotes(sUnInstallString);
    if Exec(sUnInstallString, '/SILENT /NORESTART /SUPPRESSMSGBOXES','', SW_HIDE, ewWaitUntilTerminated, iResultCode) then
      Result := 3
    else
      Result := 2;
  end else
    Result := 1;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if (CurPageID = wpSelectTasks) then begin
    if IsUpgrade() then begin
      WizardForm.TasksList.Checked[0] := False; // Don't create desktop icon on upgrade by default
    end;
  end;
end;

function InitializeSetup(): Boolean;
begin
  // Check for minimum requirements
  Result := True;
  
  // Add any prerequisite checks here
  // For example, check for .NET Framework, Visual C++ Redistributables, etc.
  
  if not Result then begin
    MsgBox('System requirements not met. Please install required components and try again.', 
           mbError, MB_OK);
  end;
end;

procedure InitializeWizard();
begin
  // Customize the installer appearance if needed
  // WizardForm.WizardBitmapImage.Bitmap.LoadFromFile(ExpandConstant('{src}\installer_banner.bmp'));
end;