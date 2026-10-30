; Inno Setup: пакує папку dist\PlenaryAgenda у звичайний Windows-установник.
#define AppName "Порядок денний"
#define AppExe "PlenaryAgenda.exe"

[Setup]
AppId={{7E1C5B52-6C0B-4B0E-9E43-5C1A3B7A1D01}
AppName={#AppName}
AppVersion=1.0
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
OutputDir=installer
OutputBaseFilename=PlenaryAgenda-Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#AppExe}

[Files]
Source: "dist\PlenaryAgenda\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"

[Run]
Filename: "{app}\{#AppExe}"; Description: "Запустити {#AppName}"; Flags: nowait postinstall skipifsilent
