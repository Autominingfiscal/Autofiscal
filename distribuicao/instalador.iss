; Instalador do Autofiscal (Inno Setup 6).
; Nao compile direto: use gerar_instalador.py, que passa Versao, Programa,
; Saida e Icone com /D.
;
; Instala na pasta do USUARIO (%LOCALAPPDATA%\Programs\Autofiscal), sem pedir
; administrador: as ferramentas gravam configuracao, logs e backups na propria
; pasta, e em "Arquivos de Programas" o Windows nao deixa.

#ifndef Versao
  #error "Rode pelo gerar_instalador.py"
#endif

[Setup]
AppId={{E28AF974-C1A8-4D37-9708-06CEBF0DF454}
AppName=Autofiscal
AppVersion={#Versao}
AppVerName=Autofiscal {#Versao}
AppPublisher=Autofiscal
DefaultDirName={localappdata}\Programs\Autofiscal
DefaultGroupName=Autofiscal
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir={#Saida}
OutputBaseFilename=Autofiscal-{#Versao}-instalador
SetupIconFile={#Icone}
UninstallDisplayIcon={app}\Autofiscal.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes

[Languages]
Name: "pt"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "atalho"; Description: "Criar atalho na Área de Trabalho"; GroupDescription: "Atalhos:"

[InstallDelete]
; bibliotecas da versao anterior: troca tudo, para nao sobrar arquivo velho
Type: filesandordirs; Name: "{app}\_internal"

[Files]
; codigo: sempre substituido pela versao nova
Source: "{#Programa}\*"; DestDir: "{app}"; Excludes: "*.ini"; \
  Flags: ignoreversion recursesubdirs createallsubdirs
; configuracoes (.ini): so na primeira instalacao. Atualizar nao apaga o que o
; usuario ajustou, e desinstalar tambem nao.
Source: "{#Programa}\*.ini"; DestDir: "{app}"; \
  Flags: onlyifdoesntexist uninsneveruninstall recursesubdirs

[Icons]
Name: "{group}\Autofiscal"; Filename: "{app}\Autofiscal.exe"
Name: "{group}\Novidades desta versão"; Filename: "{app}\CHANGELOG.md"
Name: "{group}\Desinstalar o Autofiscal"; Filename: "{uninstallexe}"
Name: "{userdesktop}\Autofiscal"; Filename: "{app}\Autofiscal.exe"; Tasks: atalho

[Run]
Filename: "{app}\Autofiscal.exe"; Description: "Abrir o Autofiscal agora"; \
  Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\_internal"
Type: filesandordirs; Name: "{app}\*\__pycache__"
Type: filesandordirs; Name: "{app}\comum\__pycache__"
