[Setup]
AppName=Sinergia
AppVersion=1.0
DefaultDirName={localappdata}\Sinergia
DefaultGroupName=Sinergia
OutputDir=installer_output
OutputBaseFilename=Sinergia_Installer
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Files]
Source: "dist\portable_manager.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Sinergia"; Filename: "{app}\portable_manager.exe"
Name: "{autodesktop}\Sinergia"; Filename: "{app}\portable_manager.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el Escritorio"; GroupDescription: "Adicionales:"