; Inno Setup Script para CORREAGRO Extractor
; Requiere:
; 1) Compilar la app con PyInstaller (genera dist\correagro_extractor\correagro_extractor.exe)
; 2) Compilar este script con Inno Setup (ISCC)

[Setup]
AppId={{6A1D92B0-CE61-4F50-9B8F-9F7E9A7081F7}
AppName=CORREAGRO Extractor
AppVersion=1.0.0
AppPublisher=CORREAGRO S.A.
AppPublisherURL=https://www.correagro.com.co
DefaultDirName={pf}\CORREAGRO\Extractor
DefaultGroupName=CORREAGRO
DisableProgramGroupPage=yes
OutputDir=..\..\dist\installer
OutputBaseFilename=Correagro-Extractor-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el Escritorio"; GroupDescription: "Tareas adicionales:"; Flags: unchecked

[Files]
; Copia todo el contenido generado por PyInstaller
Source: "..\..\dist\correagro_extractor\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\CORREAGRO Extractor"; Filename: "{app}\correagro_extractor.exe"
Name: "{userdesktop}\CORREAGRO Extractor"; Filename: "{app}\correagro_extractor.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\correagro_extractor.exe"; Description: "Ejecutar CORREAGRO Extractor"; Flags: nowait postinstall skipifsilent

