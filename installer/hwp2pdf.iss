#define AppVersion GetEnv("HWP2PDF_VERSION")
#define AppRoot GetEnv("HWP2PDF_ROOT")

[Setup]
AppId={{8F377E11-3EB4-4F62-8D62-626D8C8241F1}
AppName=hwp2pdf
AppVersion={#AppVersion}
AppPublisher=Namun Cho
AppPublisherURL=https://github.com/z0nam/hwp2pdf
AppSupportURL=https://github.com/z0nam/hwp2pdf/issues
AppUpdatesURL=https://github.com/z0nam/hwp2pdf/releases/latest
DefaultDirName={autopf}\hwp2pdf
DefaultGroupName=hwp2pdf
DisableProgramGroupPage=yes
OutputDir={#AppRoot}\release
OutputBaseFilename=hwp2pdf-setup-{#AppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile={#AppRoot}\assets\hwp_to_pdf_final.ico
UninstallDisplayIcon={app}\hwp2pdf.exe
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "{#AppRoot}\dist\hwp2pdf-{#AppVersion}.exe"; DestDir: "{app}"; DestName: "hwp2pdf.exe"; Flags: ignoreversion
Source: "{#AppRoot}\dist\hwp2pdf-cli-{#AppVersion}.exe"; DestDir: "{app}"; DestName: "hwp2pdf-cli.exe"; Flags: ignoreversion
Source: "{#AppRoot}\vendor\x86\FilePathCheckerModule.dll"; DestDir: "{app}\vendor\x86"; Flags: ignoreversion
Source: "{#AppRoot}\vendor\x64\FilePathCheckerModule.dll"; DestDir: "{app}\vendor\x64"; Flags: ignoreversion

[Registry]
; Register Hancom HWP file-access security module so headless conversions skip the
; "파일에 접근하려는 시도가 있습니다. 접근을 허용하시겠습니까?" dialog. HKCU so no
; admin rights are required. Defaults to x86 (32-bit HWP is the common case);
; the app re-validates on launch and re-points to vendor\x64 when the installed HWP is 64-bit.
Root: HKCU; Subkey: "Software\HNC\HwpAutomation\Modules"; ValueType: string; ValueName: "FilePathCheckerModule"; ValueData: "{app}\vendor\x86\FilePathCheckerModule.dll"; Flags: uninsdeletevalue

[Icons]
Name: "{group}\hwp2pdf"; Filename: "{app}\hwp2pdf.exe"
Name: "{group}\hwp2pdf CLI"; Filename: "{app}\hwp2pdf-cli.exe"
Name: "{group}\Uninstall hwp2pdf"; Filename: "{uninstallexe}"
Name: "{autodesktop}\hwp2pdf"; Filename: "{app}\hwp2pdf.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\hwp2pdf.exe"; Description: "{cm:LaunchProgram,hwp2pdf}"; Flags: nowait postinstall skipifsilent
