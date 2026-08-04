#define MyAppName "Word Continuation Checker"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Your Company"
#define MyAppExeName "WordContinuationServer.exe"

[Setup]
AppId={{8F6E4A16-7A78-4C48-9B6A-92EFC9E44B13}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={autopf}\Word Continuation Checker
DefaultGroupName={#MyAppName}

OutputDir=output
OutputBaseFilename=WordContinuationCheckerSetup

Compression=lzma
SolidCompression=yes
WizardStyle=modern

PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

UninstallDisplayIcon={app}\{#MyAppExeName}
CloseApplications=yes
RestartApplications=no

[Files]
; Background HTTPS server
Source: "..\local-server\WordContinuationServer\bin\Release\net8.0\win-x64\publish\WordContinuationServer.exe"; \
    DestDir: "{app}"; \
    Flags: ignoreversion

; Complete production add-in output
Source: "..\local-server\WordContinuationServer\bin\Release\net8.0\win-x64\publish\dist\*"; \
    DestDir: "{app}\dist"; \
    Flags: ignoreversion recursesubdirs createallsubdirs

; Manifest used by the shared-folder Office catalog
Source: "..\local-server\WordContinuationServer\bin\Release\net8.0\win-x64\publish\dist\manifest.xml"; \
    DestDir: "{commonappdata}\WordContinuationChecker\Catalog"; \
    Flags: ignoreversion

; HTTPS certificate files
Source: "..\local-server\WordContinuationServer\bin\Release\net8.0\win-x64\publish\certificate\*"; \
    DestDir: "{app}\certificate"; \
    Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
; Register the shared-folder catalog with Microsoft Office
Root: HKCU; \
    Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{{A98B612F-DBE8-4E45-A9EC-89AC3729FC41}"; \
    ValueType: string; \
    ValueName: "Id"; \
    ValueData: "{{A98B612F-DBE8-4E45-A9EC-89AC3729FC41}"; \
    Flags: uninsdeletekey

Root: HKCU; \
    Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{{A98B612F-DBE8-4E45-A9EC-89AC3729FC41}"; \
    ValueType: string; \
    ValueName: "Url"; \
    ValueData: "\\{computername}\WordContinuationCatalog"; \
    Flags: uninsdeletevalue

Root: HKCU; \
    Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{{A98B612F-DBE8-4E45-A9EC-89AC3729FC41}"; \
    ValueType: dword; \
    ValueName: "Flags"; \
    ValueData: "1"; \
    Flags: uninsdeletevalue

[Icons]
; Start Menu shortcut
Name: "{group}\Word Continuation Checker Server"; \
    Filename: "{app}\{#MyAppExeName}"; \
    WorkingDir: "{app}"

; Start automatically when the user signs in
Name: "{userstartup}\Word Continuation Checker Server"; \
    Filename: "{app}\{#MyAppExeName}"; \
    WorkingDir: "{app}"

[Run]
; Trust the localhost HTTPS certificate
Filename: "{cmd}"; \
    Parameters: "/C certutil -addstore -f Root ""{app}\certificate\WordContinuationChecker.cer"""; \
    Flags: runhidden waituntilterminated; \
    StatusMsg: "Installing the local HTTPS certificate..."

; Replace any previous share with the same name
Filename: "{cmd}"; \
    Parameters: "/C net share WordContinuationCatalog /delete /yes >nul 2>&1"; \
    Flags: runhidden waituntilterminated; \
    StatusMsg: "Preparing the Word add-in catalog..."

; Create the catalog share
Filename: "{cmd}"; \
    Parameters: "/C net share WordContinuationCatalog=""{commonappdata}\WordContinuationChecker\Catalog"" /GRANT:Everyone,READ"; \
    Flags: runhidden waituntilterminated; \
    StatusMsg: "Creating the Word add-in catalog..."

; Start the local HTTPS server
Filename: "{app}\{#MyAppExeName}"; \
    WorkingDir: "{app}"; \
    Description: "Start Word Continuation Checker"; \
    Flags: nowait postinstall skipifsilent

[UninstallRun]
; Stop the background server
Filename: "{cmd}"; \
    Parameters: "/C taskkill /IM WordContinuationServer.exe /F >nul 2>&1"; \
    Flags: runhidden waituntilterminated; \
    RunOnceId: "StopWordContinuationServer"

; Remove the Windows share
Filename: "{cmd}"; \
    Parameters: "/C net share WordContinuationCatalog /delete /yes >nul 2>&1"; \
    Flags: runhidden waituntilterminated; \
    RunOnceId: "DeleteWordContinuationCatalogShare"

; Read the certificate thumbprint from the installed CER and remove it
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; \
    Parameters: "-NoProfile -ExecutionPolicy Bypass -Command ""$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('{app}\certificate\WordContinuationChecker.cer'); & certutil.exe -delstore Root $cert.Thumbprint"""; \
    Flags: runhidden waituntilterminated; \
    RunOnceId: "RemoveWordContinuationCertificate"

[UninstallDelete]
Type: filesandordirs; Name: "{commonappdata}\WordContinuationChecker"