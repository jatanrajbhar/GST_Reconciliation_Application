; Inno Setup Script for GST Reconciliation Tool
; This script creates a professional Windows installer (setup.exe)

#define MyAppName "GST Reconciliation Tool"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "GSC in time"
#define MyAppURL "https://www.gscintime.com"
#define MyAppExeName "GST Reconciliation Tool.exe"

[Setup]
; Application identification
AppId={{8F3E9D2A-5C7B-4A6E-9D8F-1E2A3B4C5D6E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; Installation directory
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes

; License and documentation
LicenseFile=Terms and Conditions.txt
InfoBeforeFile=Privacy Policy.txt

; Output configuration
OutputDir=setup_output
OutputBaseFilename=GST_Reconciliation_Tool_Setup

; Visual configuration - use generated ICO file
SetupIconFile=app_icon.ico
WizardImageFile=wizard_image.bmp
WizardSmallImageFile=wizard_small.bmp

; Compression settings
Compression=lzma2/ultra64
SolidCompression=yes

; Modern installer style
WizardStyle=modern

; Privileges - allow non-admin install
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; 64-bit support
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
; Main application from PyInstaller
Source: "dist\GST Reconciliation Tool\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Template file for users
Source: "Template.xlsx"; DestDir: "{app}"; Flags: ignoreversion

; Documentation files (included in _internal folder from PyInstaller)
; Source: "Privacy Policy.docx"; DestDir: "{app}\Documentation"; Flags: ignoreversion
; Source: "Terms and Conditions.docx"; DestDir: "{app}\Documentation"; Flags: ignoreversion

; Logo file
Source: "logo small.png"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu shortcut
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

; Desktop shortcut (optional)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Launch application after installation
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Clean up any generated files on uninstall
Type: filesandordirs; Name: "{app}\__pycache__"
Type: filesandordirs; Name: "{app}\*.pyc"
; Remove activation data so the license slot is freed on this machine
Type: files; Name: "{userappdata}\GST_Reconciliation_Tool\activation.dat"
