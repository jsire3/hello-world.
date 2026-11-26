; Inno Setup script for Doc2Word.
; Build with:  iscc /DAppVersion=0.1.0 packaging\installer.iss
; (build.ps1 does this for you after PyInstaller has produced dist\Doc2Word\)
;
; Silent install for IT / Intune / GPO:
;   Doc2Word-Setup-x.y.z.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
; Silent uninstall:
;   "C:\Program Files\Doc2Word\unins000.exe" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART

#ifndef AppVersion
  #define AppVersion "0.2.0"
#endif

[Setup]
; Never change AppId - Windows uses it to recognise upgrades of the same app.
AppId={{F4E73D84-130A-479B-989A-8D020CCC2C22}
AppName=Doc2Word
AppVersion={#AppVersion}
AppVerName=Doc2Word {#AppVersion}
AppPublisher=Internal tools
DefaultDirName={autopf}\Doc2Word
DefaultGroupName=Doc2Word
DisableProgramGroupPage=yes
OutputDir=..\dist
OutputBaseFilename=Doc2Word-Setup-{#AppVersion}
SetupIconFile=doc2word.ico
UninstallDisplayIcon={app}\Doc2Word.exe
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Installs for all users on the PC (needs admin) - right for managed rollouts.
PrivilegesRequired=admin
CloseApplications=yes
MinVersion=10.0

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Shortcuts:"

[Files]
Source: "..\dist\Doc2Word\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\Doc2Word"; Filename: "{app}\Doc2Word.exe"; Comment: "Convert files to Word"
Name: "{autodesktop}\Doc2Word"; Filename: "{app}\Doc2Word.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Doc2Word.exe"; Description: "Start Doc2Word now"; Flags: nowait postinstall skipifsilent
