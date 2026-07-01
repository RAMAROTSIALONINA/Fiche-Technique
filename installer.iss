; ============================================================
;  Inno Setup Script — Fiche Technique BOGOTA
;  Compiler : Inno Setup 6  (https://jrsoftware.org/isinfo.php)
; ============================================================

[Setup]
AppName=Fiche Technique BOGOTA
AppVersion=1.0
AppPublisher=BOGOTA
AppPublisherURL=http://localhost:8000

; Installation dans AppData (pas besoin de droits admin)
DefaultDirName={localappdata}\Programs\FicheTechnique-BOGOTA
DefaultGroupName=Fiche Technique BOGOTA
PrivilegesRequired=lowest

; Fichier de sortie
OutputDir=installer-output
OutputBaseFilename=Setup-FicheTechnique-BOGOTA

; Compression maximale
Compression=lzma2/ultra64
SolidCompression=yes

; Icône de l'installeur et des raccourcis
SetupIconFile=static\logo.ico

; Infos désinstallation
UninstallDisplayName=Fiche Technique BOGOTA
UninstallDisplayIcon={app}\FicheTechnique-BOGOTA.exe

; Écran de bienvenue
WizardStyle=modern
WizardSizePercent=100

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Files]
; Tout le dossier dist (exe + _internal)
Source: "dist\FicheTechnique-BOGOTA\FicheTechnique-BOGOTA.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\FicheTechnique-BOGOTA\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Raccourci Bureau
Name: "{userdesktop}\Fiche Technique BOGOTA"; Filename: "{app}\FicheTechnique-BOGOTA.exe"; WorkingDir: "{app}"; Comment: "Ouvrir Fiche Technique BOGOTA"

; Raccourci Menu Démarrer
Name: "{group}\Fiche Technique BOGOTA"; Filename: "{app}\FicheTechnique-BOGOTA.exe"; WorkingDir: "{app}"; Comment: "Ouvrir Fiche Technique BOGOTA"
Name: "{group}\Désinstaller Fiche Technique BOGOTA"; Filename: "{uninstallexe}"

[Run]
; Proposer de lancer l'app à la fin de l'installation
Filename: "{app}\FicheTechnique-BOGOTA.exe"; Description: "Lancer Fiche Technique BOGOTA maintenant"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Supprimer le log et la base de données à la désinstallation
Type: files; Name: "{app}\fiche-technique.log"
; Note : on garde fiches_bogota.db pour ne pas perdre les données
