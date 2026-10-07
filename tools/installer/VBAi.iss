; SPDX-License-Identifier: MPL-2.0
; Build through Build-Installer.ps1; compile definitions are validated there.
#ifndef PackageDirectory
  #error PackageDirectory is required
#endif
#ifndef ProductVersion
  #error ProductVersion is required
#endif

[Setup]
AppId=VBAi.win-x64
AppName=VBAi
AppVersion={#ProductVersion}
AppPublisher=VBAi contributors
AppPublisherURL=https://github.com/blackcancer/VBAi
AppSupportURL=https://github.com/blackcancer/VBAi/issues
DefaultDirName={localappdata}\Programs\VBAi
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
WizardStyle=modern
SetupIconFile=..\..\assets\icons\assistant.ico
UninstallDisplayIcon={app}\assistant.ico
UninstallDisplayName=VBAi
LicenseFile=..\..\LICENSE
OutputDir={#SetupOutputDirectory}
OutputBaseFilename=VBAi-Setup-win-x64
Compression=lzma2
SolidCompression=yes
SetupMutex=VBAi.Setup.win-x64
CloseApplications=no
RestartApplications=no
AlwaysRestart=no
#ifdef SignedBuild
SignTool=vbai
SignedUninstaller=yes
SignedUninstallerDir={#SetupOutputDirectory}\signed-uninstaller
#else
SignedUninstaller=no
#endif

[Languages]
Name: english; MessagesFile: compiler:Default.isl
Name: french; MessagesFile: compiler:Languages\French.isl

[Files]
; The installation marker is written after successful registration, preserving its ID.
Source: "{#PackageDirectory}\*"; DestDir: "{app}"; Excludes: "vbai-installation.json"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "SetupRegistration.ps1"; DestDir: "{app}\Setup"; Flags: ignoreversion
Source: "..\Install-VBAi.ps1"; DestDir: "{app}\Setup"; Flags: ignoreversion
Source: "..\Uninstall-VBAi.ps1"; DestDir: "{app}\Setup"; Flags: ignoreversion
Source: "..\Register-VBAiTypeLib.ps1"; DestDir: "{app}\Setup"; Flags: ignoreversion
Source: "..\Register-ChatToolWindow.ps1"; DestDir: "{app}\Setup"; Flags: ignoreversion
Source: "..\Register-VbaTestRuntime.ps1"; DestDir: "{app}\Setup"; Flags: ignoreversion
Source: "..\..\assets\icons\assistant.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{userprograms}\VBAi\User guide"; Filename: "{app}\Help\VBAi.en-US.chm"; Check: not IsFrench
Name: "{userprograms}\VBAi\Guide utilisateur"; Filename: "{app}\Help\VBAi.fr-FR.chm"; Check: IsFrench
Name: "{userprograms}\VBAi\Uninstall VBAi"; Filename: "{uninstallexe}"

[Code]
function IsFrench: Boolean;
begin
  Result := ActiveLanguage = 'french';
end;

function Registration(Script, Action: String; Removing: Boolean): Boolean;
var
  Params, OutputLog: String;
  ExitCode: Integer;
begin
  OutputLog := ExpandConstant('{localappdata}\VBAi\SetupLogs');
  if not ForceDirectories(OutputLog) then begin Result := False; exit; end;
  OutputLog := OutputLog + '\' + GetDateTimeString('yyyymmdd-hhnnss', '-', ':') + '-' + Action + '.log';
  Params := '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + Script +
    '" -Action ' + Action + ' -InstallationDirectory "' + ExpandConstant('{app}') +
    '" -Version "{#ProductVersion}" -LogPath "' + OutputLog + '"';
  if Removing then Params := Params + ' -ForUninstall';
  Result := Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
    Params, '', SW_HIDE, ewWaitUntilTerminated, ExitCode);
  if Result then Result := ExitCode = 0;
  if not Result then Log('Registration failed; inspect ' + OutputLog);
end;

function InitializeSetup: Boolean;
var
  Release: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release);
  if Result then Result := Release >= 528040;
  if not Result then MsgBox('VBAi requires .NET Framework 4.8 or later.', mbError, MB_OK);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  ExtractTemporaryFile('SetupRegistration.ps1');
  if Registration(ExpandConstant('{tmp}\SetupRegistration.ps1'), 'Preflight', False) then
    Result := ''
  else
    Result := 'Save your work and close VBA hosts. Check %LOCALAPPDATA%\VBAi\SetupLogs for registration or path conflicts.';
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    if not Registration(ExpandConstant('{app}\Setup\SetupRegistration.ps1'), 'Install', False) then
      RaiseException('VBAi registration failed. Preserve the setup log and recovery snapshot; installation is not accepted.');
end;

function InitializeUninstall: Boolean;
begin
  Result := Registration(ExpandConstant('{app}\Setup\SetupRegistration.ps1'), 'Preflight', True);
  if not Result then MsgBox('Close VBA hosts and check %LOCALAPPDATA%\VBAi\SetupLogs. VBAi has not been removed.', mbError, MB_OK);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
    if not Registration(ExpandConstant('{app}\Setup\SetupRegistration.ps1'), 'Uninstall', True) then
      RaiseException('COM removal failed. Files must be preserved for recovery.');
end;
