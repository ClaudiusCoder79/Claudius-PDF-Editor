; Claudius PDF Editor installer - build with Inno Setup 6.3 or newer (free, jrsoftware.org)
; Needs the folder dist\Claudius PDF Editor (run build_exe.bat first) and claudius.ico / README.txt next to this file.

#define MyAppName "Claudius PDF Editor"
#define MyAppVersion "3.7"
#define MyAppExe "Claudius PDF Editor.exe"
; the program folder is taken from dist\Claudius PDF Editor (made by build_exe.bat), or from a folder "Claudius PDF Editor" placed next to this script
#if FileExists("dist\Claudius PDF Editor\Claudius PDF Editor.exe")
  #define ExeSource "dist\Claudius PDF Editor"
#elif FileExists("Claudius PDF Editor\Claudius PDF Editor.exe")
  #define ExeSource "Claudius PDF Editor"
#else
  #error The program folder was not found. Run build_exe.bat first (it makes dist\Claudius PDF Editor) and compile again.
#endif

[Setup]
AppId={{B7C1E5A2-3F4D-4A6B-9C21-5D8E7F0A1B33}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher=Michal Czapla & Claude AI
AppPublisherURL=https://github.com/ClaudiusCoder79/Claudius-PDF-Editor
AppSupportURL=https://github.com/ClaudiusCoder79/Claudius-PDF-Editor
VersionInfoVersion={#MyAppVersion}.0.0
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
UninstallDisplayName={#MyAppName}
; installs into C:\Program Files (needs administrator rights)
PrivilegesRequired=admin
OutputDir=installer
OutputBaseFilename=Claudius-PDF-Editor-Setup-{#MyAppVersion}
LicenseFile=FREEWARE.txt
#if FileExists("claudius.ico")
SetupIconFile=claudius.ico
#endif
UninstallDisplayIcon={app}\{#MyAppExe}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
CloseApplications=yes
RestartApplications=no
ChangesAssociations=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "openwith"; Description: "Add Claudius PDF Editor to the ""Open with"" list and to Windows Default apps (PDF files)"; GroupDescription: "File integration:"

[Files]
Source: "{#ExeSource}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "claudius.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.txt"; DestDir: "{app}"; Flags: ignoreversion isreadme
Source: "FREEWARE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"; Tasks: desktopicon

[Registry]
Root: HKA; Subkey: "Software\Classes\CLAUDIUS.PDF"; ValueType: string; ValueName: ""; ValueData: "PDF document (Claudius PDF Editor)"; Flags: uninsdeletekey; Tasks: openwith
Root: HKA; Subkey: "Software\Classes\CLAUDIUS.PDF\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#MyAppExe},0"; Tasks: openwith
Root: HKA; Subkey: "Software\Classes\CLAUDIUS.PDF\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExe}"" ""%1"""; Tasks: openwith
Root: HKA; Subkey: "Software\Classes\.pdf\OpenWithProgids"; ValueType: none; ValueName: "CLAUDIUS.PDF"; Flags: uninsdeletevalue; Tasks: openwith

Root: HKA; Subkey: "Software\CLAUDIUS\Capabilities"; ValueType: string; ValueName: "ApplicationName"; ValueData: "Claudius PDF Editor"; Flags: uninsdeletekey; Tasks: openwith
Root: HKA; Subkey: "Software\CLAUDIUS\Capabilities"; ValueType: string; ValueName: "ApplicationDescription"; ValueData: "PDF editor"; Tasks: openwith
Root: HKA; Subkey: "Software\CLAUDIUS\Capabilities\FileAssociations"; ValueType: string; ValueName: ".pdf"; ValueData: "CLAUDIUS.PDF"; Tasks: openwith
Root: HKA; Subkey: "Software\RegisteredApplications"; ValueType: string; ValueName: "CLAUDIUS"; ValueData: "Software\CLAUDIUS\Capabilities"; Flags: uninsdeletevalue; Tasks: openwith

[Run]
Filename: "ms-settings:defaultapps?registeredAppUser=CLAUDIUS"; Description: "Open Windows Default apps to make Claudius PDF Editor your PDF app"; Flags: shellexec nowait postinstall skipifsilent unchecked; Tasks: openwith
Filename: "{app}\{#MyAppExe}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent
