; Inno Setup Script - WF-Mag Konwerter
; Instalacja BEZ uprawnien administratora (instaluje do folderu uzytkownika)
;
; Aby zbudowac instalke:
; 1. Najpierw zbuduj EXE: uruchom build.bat
; 2. Zainstaluj Inno Setup: https://jrsoftware.org/isdl.php
; 3. Otworz ten plik w Inno Setup i kliknij "Compile"

#define MyAppName "WF-Mag Konwerter"
#define MyAppVersion "1.0"
#define MyAppPublisher "MOREX GM"
#define MyAppExeName "WFMag-Konwerter.exe"

[Setup]
AppId={{B8E3F2A1-4C7D-4E8F-9A1B-2C3D4E5F6A7B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=WFMag-Konwerter-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline

[Languages]
Name: "polish"; MessagesFile: "compiler:Languages\Polish.isl"

[Tasks]
Name: "desktopicon"; Description: "Utwórz skrót na pulpicie"; GroupDescription: "Dodatkowe opcje:"; Flags: unchecked

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Uruchom WF-Mag Konwerter"; Flags: nowait postinstall skipifsilent
