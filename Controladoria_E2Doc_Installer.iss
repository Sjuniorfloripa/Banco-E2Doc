; ==============================================================
;  Instalador Oficial - Controladoria Automate E2Doc
; ==============================================================

[Setup]
AppName=Controladoria Automate E2Doc
AppVersion=1.0.0
AppPublisher=EQS Engenharia
DefaultDirName={pf}\Controladoria Automate E2Doc
OutputDir=Output
OutputBaseFilename=Controladoria-Automate-E2Doc-Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
SetupIconFile="assets\icons\app_e2doc_icon.ico"

; ================== ARQUIVOS A SEREM INSTALADOS =====================

[Files]
Source: "dist\EQS Automate Conversor.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: ".env_example"; DestDir: "{app}"; DestName: ".env.example"; Flags: ignoreversion onlyifdoesntexist

; ================== CRIA .env INICIAL CASO NÃO EXISTA =====================

[Run]
Filename: "{cmd}"; \
Parameters: "/c if not exist ""{app}\.env"" (copy ""{app}\.env.example"" ""{app}\.env"")"; \
Flags: runhidden

; ==================== ATALHO NA ÁREA DE TRABALHO ====================

[Icons]
Name: "{commondesktop}\Controladoria Automate E2Doc"; \
Filename: "{app}\EQS Automate Conversor.exe"; \
WorkingDir: "{app}"

; ==================== DESINSTALADOR (LIMPEZA) ====================

[UninstallDelete]
Type: filesandordirs; Name: "{app}\*"

