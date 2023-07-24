; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "������ �������� ���������"
#define MyAppVersion "1.0"
#define MyAppPublisher "BihmaVD."
#define MyAppURL "https://github.com/Bihma-Vladyslav/Dyplom1"
#define MyAppExeName "Dyplom1.exe"
#define MyAppAssocName "���� ������ �������� ��������� "
#define MyAppAssocExt ".exe"
#define MyAppAssocKey StringChange(MyAppAssocName, " ", "") + MyAppAssocExt

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{F48CD9D8-A884-4C5C-AF06-938AB431A35C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
ChangesAssociations=yes
DisableProgramGroupPage=yes
; Uncomment the following line to run in non administrative install mode (install for current user only.)
;PrivilegesRequired=lowest
OutputDir=C:\Users\Vlad\Documents\���� �����
OutputBaseFilename=���� ������������
SetupIconFile=D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "ukrainian"; MessagesFile: "compiler:Languages\Ukrainian.isl"

[Files]
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\Diplom.db"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\Dyplom1.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\Dyplom1.exe.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\Dyplom1.pdb"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\EntityFramework.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\EntityFramework.SqlServer.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\EntityFramework.SqlServer.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\EntityFramework.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\System.Data.SQLite.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\System.Data.SQLite.EF6.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\System.Data.SQLite.Linq.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\System.Data.SQLite.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\x64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "D:\Visual Studio ������������\source\repos\Dyplom1\Dyplom1\bin\Release\x86\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Registry]
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocExt}\OpenWithProgids"; ValueType: string; ValueName: "{#MyAppAssocKey}"; ValueData: ""; Flags: uninsdeletevalue
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}"; ValueType: string; ValueName: ""; ValueData: "{#MyAppAssocName}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#MyAppExeName},0"
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".myp"; ValueData: ""

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

