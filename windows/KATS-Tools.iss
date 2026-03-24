#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

#define MyAppName "KATS-Tools"
#define MyAppPublisher "Your Company"
#define MyAppId "{{D6F5E9F2-4D37-4D1E-AF2B-1E7F4F5D1A11}}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
SourceDir=..
DefaultDirName={userappdata}\KATS-Tools
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
Compression=lzma
SolidCompression=yes
WizardStyle=modern
Uninstallable=no
CreateUninstallRegKey=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "assets\KATS-Tools.dotm"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"; Flags: ignoreversion
Source: "assets\KATSUpdater.bat"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"; Flags: ignoreversion

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    MsgBox(
      'KATS-Tools har installerats.' + #13#10#13#10 +
      'Starta om Word för att ladda tillägget.',
      mbInformation, MB_OK);
  end;
end;

