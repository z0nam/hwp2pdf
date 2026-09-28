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
PrivilegesRequired=admin
ChangesEnvironment=yes
SetupIconFile={#AppRoot}\assets\hwp_to_pdf_final.ico
UninstallDisplayIcon={app}\hwp2pdf.exe
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
korean.AddToPath=명령줄에서 hwp2pdf-cli 사용(PATH에 추가)
english.AddToPath=Use hwp2pdf-cli from the command line (add to PATH)

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "addtopath"; Description: "{cm:AddToPath}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "{#AppRoot}\dist\hwp2pdf-{#AppVersion}.exe"; DestDir: "{app}"; DestName: "hwp2pdf.exe"; Flags: ignoreversion
Source: "{#AppRoot}\dist\hwp2pdf-cli-{#AppVersion}.exe"; DestDir: "{app}"; DestName: "hwp2pdf-cli.exe"; Flags: ignoreversion
Source: "{#AppRoot}\dist\hwp2pdf-serve-{#AppVersion}.exe"; DestDir: "{app}"; DestName: "hwp2pdf-serve.exe"; Flags: ignoreversion
Source: "{#AppRoot}\THIRD_PARTY_NOTICES.md"; DestDir: "{app}"; Flags: ignoreversion
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
; Conversion server for macOS/Linux clients. Windowless on purpose: there is no
; console window to close by accident, and it logs to %LOCALAPPDATA%\hwp2pdf\server.log.
; It runs in this logged-in desktop session -- Hangul automation does not work
; as a Windows Service.
Name: "{group}\hwp2pdf 변환 서버 (Conversion Server)"; Filename: "{app}\hwp2pdf-serve.exe"; Parameters: "--bind tailscale --init"
Name: "{group}\Uninstall hwp2pdf"; Filename: "{uninstallexe}"
Name: "{autodesktop}\hwp2pdf"; Filename: "{app}\hwp2pdf.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\hwp2pdf.exe"; Description: "{cm:LaunchProgram,hwp2pdf}"; Flags: nowait postinstall skipifsilent
Filename: "{app}\hwp2pdf.exe"; Flags: nowait runasoriginaluser skipifnotsilent; Check: IsAutoUpdate

[Code]
const
  EnvironmentKey = 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment';
  AppRegistryKey = 'Software\hwp2pdf';
  PathMarkerName = 'InstallerPathEntry';

function IsAutoUpdate: Boolean;
begin
  Result := ExpandConstant('{param:HWP2PDFAUTOUPDATE|0}') = '1';
end;

function NormalizedPath(Value: String): String;
begin
  Result := Trim(Value);
  if (Length(Result) >= 2) and (Result[1] = '"') and
     (Result[Length(Result)] = '"') then
    Result := Copy(Result, 2, Length(Result) - 2);
  StringChangeEx(Result, '/', '\', True);
  while (Length(Result) > 3) and (Result[Length(Result)] = '\') do
    Delete(Result, Length(Result), 1);
  Result := Lowercase(Result);
end;

function PathContains(const PathValue, Entry: String): Boolean;
var
  Remaining, Part, Wanted: String;
  Separator: Integer;
begin
  Result := False;
  Wanted := NormalizedPath(Entry);
  Remaining := PathValue + ';';
  while Remaining <> '' do
  begin
    Separator := Pos(';', Remaining);
    Part := Copy(Remaining, 1, Separator - 1);
    Delete(Remaining, 1, Separator);
    if NormalizedPath(Part) = Wanted then
    begin
      Result := True;
      Exit;
    end;
  end;
end;

function AddPathEntry(const Entry: String): Boolean;
var
  PathValue: String;
begin
  if not RegQueryStringValue(HKLM, EnvironmentKey, 'Path', PathValue) then
    PathValue := '';
  if PathContains(PathValue, Entry) then
  begin
    Result := False;
    Exit;
  end;

  while (Length(PathValue) > 0) and (PathValue[Length(PathValue)] = ';') do
    Delete(PathValue, Length(PathValue), 1);
  if PathValue <> '' then
    PathValue := PathValue + ';';
  Result := RegWriteExpandStringValue(HKLM, EnvironmentKey, 'Path', PathValue + Entry);
end;

function RemovePathEntry(const Entry: String): Boolean;
var
  PathValue, NewValue, Remaining, Part: String;
  Separator: Integer;
  FirstKept: Boolean;
begin
  Result := False;
  if not RegQueryStringValue(HKLM, EnvironmentKey, 'Path', PathValue) then
    Exit;

  Remaining := PathValue + ';';
  NewValue := '';
  FirstKept := True;
  while Remaining <> '' do
  begin
    Separator := Pos(';', Remaining);
    Part := Copy(Remaining, 1, Separator - 1);
    Delete(Remaining, 1, Separator);
    if NormalizedPath(Part) <> NormalizedPath(Entry) then
    begin
      if not FirstKept then
        NewValue := NewValue + ';';
      NewValue := NewValue + Part;
      FirstKept := False;
    end;
  end;

  if NewValue <> PathValue then
    Result := RegWriteExpandStringValue(HKLM, EnvironmentKey, 'Path', NewValue);
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  AppPath, Marker: String;
begin
  if CurStep <> ssPostInstall then
    Exit;

  AppPath := ExpandConstant('{app}');
  if WizardIsTaskSelected('addtopath') then
  begin
    if AddPathEntry(AppPath) then
      RegWriteStringValue(HKLM, AppRegistryKey, PathMarkerName, AppPath);
  end
  else if RegQueryStringValue(HKLM, AppRegistryKey, PathMarkerName, Marker) and
          (NormalizedPath(Marker) = NormalizedPath(AppPath)) then
  begin
    RemovePathEntry(AppPath);
    RegDeleteValue(HKLM, AppRegistryKey, PathMarkerName);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  AppPath, Marker: String;
begin
  if CurUninstallStep <> usUninstall then
    Exit;

  AppPath := ExpandConstant('{app}');
  if RegQueryStringValue(HKLM, AppRegistryKey, PathMarkerName, Marker) and
     (NormalizedPath(Marker) = NormalizedPath(AppPath)) then
  begin
    RemovePathEntry(AppPath);
    RegDeleteValue(HKLM, AppRegistryKey, PathMarkerName);
    RegDeleteKeyIfEmpty(HKLM, AppRegistryKey);
  end;
end;
