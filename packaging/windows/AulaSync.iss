; Installationsprogram til AulaSync på Windows (Inno Setup 6.6 eller nyere). Installerer for brugeren selv, uden
; administrator, i %LOCALAPPDATA%\Programs\AulaSync og lægger AulaSync i Start-menuen. Bruges ved download og af winget.
;
;   iscc /DVersion=3.2.0 /DSource=out\win /Oout\setup packaging\windows\AulaSync.iss   -> out\setup\AulaSync-Setup.exe
;
; Kører AulaSync, bedes den om at afslutte (AulaSync.exe --quit; ældre udgaver lukkes), og den startes igen bagefter.
; Har brugeren slået start ved login til, peger den bagefter på den installerede AulaSync, også når den før lå et andet
; sted (fx den løse AulaSync.exe i Dokumenter). /LAUNCH=1 (winget-manifestet) starter AulaSync efter første installation.
; Afinstallation fjerner alt uden at spørge: også brugerens data (%LOCALAPPDATA%\AulaSync) og start ved login.

#if VER < EncodeVer(6,6,0)
  #error Inno Setup 6.6 eller nyere kræves (WizardStyle "dynamic")
#endif
#ifndef Version
  #error Angiv versionen: iscc /DVersion=X.Y.Z
#endif
#ifndef Source
  #define Source SourcePath + "..\..\out\win"
#endif
; Må aldrig ændres: binder opgradering, afinstallation og winget (ProductCode "{GUID}_is1") sammen.
#define AppGuid "EB8DDE1A-F008-42A1-AA73-7110F88E4044"

[Setup]
AppId={{{#AppGuid}}
AppName=AulaSync
AppVersion={#Version}
AppVerName=AulaSync {#Version}
AppPublisher=rpaasch
AppPublisherURL=https://github.com/rpaasch/AulaSync
AppSupportURL=https://github.com/rpaasch/AulaSync/blob/master/docs/vejledning.md
AppUpdatesURL=https://github.com/rpaasch/AulaSync/releases/latest
AppCopyright=Copyright (c) 2026 Rasmus Paasch
AppComments=Holder skemaer fra Aula opdateret i din kalender
VersionInfoVersion={#Version}
VersionInfoDescription=AulaSync-installationsprogram
UninstallDisplayName=AulaSync
UninstallDisplayIcon={app}\AulaSync.exe
PrivilegesRequired=lowest
DefaultDirName={localappdata}\Programs\AulaSync
DisableDirPage=yes
DisableProgramGroupPage=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
; En kørende AulaSync lukkes i PrepareToInstall; Windows' egen lukning (Restart Manager) er kun en reserve.
CloseApplications=force
RestartApplications=no
SetupMutex=AulaSyncSetup
WizardStyle=modern dynamic
SetupIconFile=..\icon\AulaSync.ico
; AulaSyncs ikon øverst til højre i installations- og afinstallationsprogrammet (i stedet for Inno Setups eget).
WizardSmallImageFile=..\icon\wizard\AulaSync-58.png,..\icon\wizard\AulaSync-77.png,..\icon\wizard\AulaSync-97.png,..\icon\wizard\AulaSync-116.png,..\icon\wizard\AulaSync-124.png,..\icon\wizard\AulaSync-143.png,..\icon\wizard\AulaSync-159.png
ShowLanguageDialog=no
OutputBaseFilename=AulaSync-Setup
Compression=lzma2
SolidCompression=yes
SetupLogging=yes

[Languages]
Name: "da"; MessagesFile: "compiler:Languages\Danish.isl"

[Messages]
; Vises kun med vinduer; /SILENT og /VERYSILENT (winget) springer begge over.
da.ConfirmUninstall=Vil du fjerne %1 helt?%n%nDine indstillinger, dit login til Aula og kalenderfilerne bliver også slettet.
da.UninstalledAll=%1 er fjernet fra computeren.%n%nKalenderne fra AulaSync bliver ikke længere opdateret. Slet dem i dit kalenderprogram.
da.UninstalledMost=%1 er fjernet fra computeren, men nogle filer kunne ikke slettes. De kan slettes manuelt.%n%nKalenderne fra AulaSync bliver ikke længere opdateret. Slet dem i dit kalenderprogram.

[CustomMessages]
da.DataLeft=Nogle af AulaSyncs filer var stadig i brug og kunne ikke slettes. Slet det, der er tilbage, når du har genstartet computeren:%n%n%1
da.OtherCopy=Der ligger stadig en AulaSync uden installation:%n%n%1%n%nDen starter ikke længere ved login. Slet den, hvis du ikke bruger den (fra winget: skriv winget uninstall rpaasch.AulaSync i Terminal).

[Files]
Source: "{#Source}\AulaSync.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Direkte i Start-menuen, uden undermappe: samme sted som AulaSync selv skriver genvejen (StartMenuShortcut.cs).
Name: "{autoprograms}\AulaSync"; Filename: "{app}\AulaSync.exe"; WorkingDir: "{app}"; Comment: "Holder skemaer fra Aula opdateret i din kalender"

[Run]
; Med vinduer: "Start AulaSync" på sidste side (ikke som administrator, se Elevated).
Filename: "{app}\AulaSync.exe"; Description: "{cm:LaunchProgram,AulaSync}"; Flags: nowait postinstall skipifsilent; Check: not Elevated
; Uden vinduer (winget): kørte AulaSync før, startes den igen i systembakken; ved første installation med /LAUNCH=1
; åbnes den.
Filename: "{app}\AulaSync.exe"; Parameters: "{code:LaunchParameters}"; Flags: nowait skipifnotsilent; Check: LaunchSilently

[Code]
const
  RunKey = 'Software\Microsoft\Windows\CurrentVersion\Run';
  ApprovedKey = 'Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run';
  UninstallKey = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{{#AppGuid}}_is1';
  // AulaSync 2 holder ingen låsefil, men denne mutex (OldVersion.cs).
  OldMutex = 'AulaSync_SingleInstance';

var
  WasRunning: Boolean;
  FirstInstall: Boolean;
  Installed: Boolean;
  // Afinstallation: noget af dataene kunne ikke slettes; en løs AulaSync ligger stadig et andet sted.
  DataLeft: String;
  OtherCopy: String;

const
  TOKEN_QUERY = $0008;
  TokenElevationType = 18;
  TokenElevationTypeFull = 2;

function GetCurrentProcess: THandle; external 'GetCurrentProcess@kernel32.dll stdcall';
function OpenProcessToken(Process: THandle; Access: DWORD; var Token: THandle): BOOL; external 'OpenProcessToken@advapi32.dll stdcall';
function GetTokenInformation(Token: THandle; InfoClass: Integer; var Info: Integer; Len: DWORD; var RetLen: DWORD): BOOL; external 'GetTokenInformation@advapi32.dll stdcall';
function CloseHandle(Handle: THandle): BOOL; external 'CloseHandle@kernel32.dll stdcall';

// Kører installationsprogrammet med forhøjede rettigheder (fx winget i et administrator-vindue eller "Kør som
// administrator")? Så startes AulaSync ikke: den ville køre som administrator, og en almindelig start eller opdatering
// kunne ikke nå eller lukke den. Er brugerkontokontrol slået fra, kører alt sådan, og så startes den.
function Elevated: Boolean;
var
  Token: THandle;
  Kind: Integer;
  Len: DWORD;
begin
  Result := IsAdmin;
  if OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, Token) then
  begin
    if GetTokenInformation(Token, TokenElevationType, Kind, 4, Len) then
      Result := Kind = TokenElevationTypeFull;
    CloseHandle(Token);
  end;
end;

function AppExe: String;
begin
  Result := ExpandConstant('{app}\AulaSync.exe');
end;

// Brugerens data: indstillinger, abonnementer, kalenderfiler, log, låsefil og login (webview\). Den samme mappe for
// alle AulaSync'er, uanset hvor exe'en ligger (AppPaths.cs).
function DataDir: String;
begin
  Result := ExpandConstant('{localappdata}\AulaSync');
end;

function LockFile: String;
begin
  Result := DataDir + '\aulasync.lock';
end;

// AulaSync 2's mappe med det gamle login; tom, hvis USERPROFILE mangler (så slettes intet).
function OldDataDir: String;
begin
  Result := GetEnv('USERPROFILE');
  if Result <> '' then Result := Result + '\.aulasync';
end;

// Her pakker .NET AulaSyncs egne DLL'er ud (én exe med IncludeNativeLibrariesForSelfExtract): %TEMP%\.net\<app>\<id>.
// .NET bruger samme midlertidige mappe som Windows (TMP før TEMP).
function ExtractDir: String;
begin
  Result := GetEnv('TMP');
  if Result = '' then Result := GetEnv('TEMP');
  if Result <> '' then Result := Result + '\.net\AulaSync';
end;

// Kører AulaSync? Den holder låsefilen åben uden deling.
function AulaSyncRunning: Boolean;
var
  Lock: TFileStream;
begin
  Result := False;
  if not FileExists(LockFile) then Exit;
  try
    Lock := TFileStream.Create(LockFile, fmOpenReadWrite or fmShareExclusive);
    Lock.Free;
  except
    Result := True;
  end;
end;

// Lukker alle brugerens AulaSync.exe, også løse kopier og AulaSync 2.
procedure KillAulaSync;
var
  Code: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM AulaSync.exe /FI "USERNAME eq ' + GetUserNameString + '"',
    '', SW_HIDE, ewWaitUntilTerminated, Code);
end;

// Beder AulaSync om at afslutte (Exe --quit). Svarer den ikke (ældre udgave), lukkes brugerens AulaSync-processer.
procedure StopAulaSync(const Exe: String);
var
  Code, I: Integer;
begin
  if FileExists(Exe) then
    Exec(Exe, '--quit', '', SW_HIDE, ewWaitUntilTerminated, Code);
  if AulaSyncRunning then
  begin
    Log('AulaSync afsluttede ikke på --quit; den lukkes.');
    KillAulaSync;
    I := 0;
    while AulaSyncRunning and (I < 50) do
    begin
      Sleep(100);
      I := I + 1;
    end;
  end;
end;

// Antal msedgewebview2.exe, der bruger AulaSyncs login (webview\, i AulaSync 2 .aulasync\webview2\): kommandolinjen
// nævner mappen (--user-data-dir, og crashpad med --database). Edge (msedge.exe) og andre programmers WebView2 har
// andre mapper og røres ikke; andre brugeres processer giver CommandLine = Null. Kill lukker dem. Uden WMI svares 0,
// og sletningen prøver blot igen.
function OurWebViews(Kill: Boolean): Integer;
var
  Locator, Wmi, Procs, Proc: Variant;
  Dir, OldDir, Cmd: String;
  I, N, Pid, Code: Integer;
begin
  Result := 0;
  Dir := Lowercase(DataDir + '\');
  OldDir := Lowercase(OldDataDir + '\');
  if OldDir = '\' then OldDir := Dir;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    Wmi := Locator.ConnectServer('.', 'root\CIMV2');
    Procs := Wmi.ExecQuery('SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name = ''msedgewebview2.exe''');
    N := Procs.Count;
    for I := 0 to N - 1 do
    begin
      Proc := Procs.ItemIndex(I);
      if not VarIsNull(Proc.CommandLine) then
      begin
        // Lowercase tager ikke en Variant i Inno Setup 6 (Type mismatch).
        Cmd := Proc.CommandLine;
        Cmd := Lowercase(Cmd);
        if (Pos(Dir, Cmd) > 0) or (Pos(OldDir, Cmd) > 0) then
        begin
          Result := Result + 1;
          if Kill then
          begin
            Pid := Proc.ProcessId;
            Log(Format('Lukker WebView2-processen %d', [Pid]));
            Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /T /PID ' + IntToStr(Pid), '', SW_HIDE, ewWaitUntilTerminated, Code);
          end;
        end;
      end;
    end;
  except
    Log('Kunne ikke finde WebView2-processerne: ' + GetExceptionMessage);
  end;
end;

// WebView2 afslutter sine processer lidt efter AulaSync og kan så længe holde filer i login-mappen åbne. Venter op til
// 5 s og lukker så dem, der er tilbage.
procedure StopOurWebViews;
var
  I: Integer;
begin
  if not DirExists(DataDir + '\webview') and not DirExists(OldDataDir + '\webview2') then Exit;
  I := 0;
  while (OurWebViews(False) > 0) and (I < 10) do
  begin
    Sleep(500);
    I := I + 1;
  end;
  if OurWebViews(True) > 0 then Sleep(500);
end;

// Sletter Dir helt. En fil, der lige er sluppet, kan være i brug et øjeblik endnu, så der prøves igen i op til 10 s;
// halvvejs lukkes WebView2-processer, der stadig er der. False, hvis noget er tilbage.
function DeleteAll(const Dir: String): Boolean;
var
  I: Integer;
begin
  Result := True;
  if (Dir = '') or not DirExists(Dir) then Exit;
  I := 0;
  while not DelTree(Dir, True, True, True) and DirExists(Dir) do
  begin
    I := I + 1;
    if I >= 20 then
    begin
      Log('Kunne ikke slette alt i ' + Dir);
      Result := False;
      Exit;
    end;
    if I = 10 then OurWebViews(True);
    Sleep(500);
  end;
  Log('Slettet: ' + Dir);
end;

// Den exe, startværdien peger på: "C:\...\AulaSync.exe" --silent.
function RunTarget(Command: String): String;
var
  P: Integer;
begin
  Command := Trim(Command);
  if Copy(Command, 1, 1) = '"' then
  begin
    Delete(Command, 1, 1);
    P := Pos('"', Command);
  end
  else
    P := Pos(' ', Command);
  if P > 0 then Command := Copy(Command, 1, P - 1);
  Result := Command;
end;

// Start ved login fjernes, når den starter en AulaSync, uanset hvor exe'en ligger: uden data ville en løs AulaSync fra
// før 3.2 ellers starte forfra ved næste login. Findes den løse exe stadig, returneres stien; den slettes ikke.
function RemoveAutostart: String;
var
  Command, Target: String;
begin
  Result := '';
  if RegQueryStringValue(HKCU, RunKey, 'AulaSync', Command) and (Pos('aulasync.exe', Lowercase(Command)) > 0) then
  begin
    Target := RunTarget(Command);
    if (CompareText(Target, AppExe) <> 0) and FileExists(Target) then
      Result := Target;
    RegDeleteValue(HKCU, RunKey, 'AulaSync');
  end;
  if not RegValueExists(HKCU, RunKey, 'AulaSync') then
    RegDeleteValue(HKCU, ApprovedKey, 'AulaSync');
end;

// winget's portable AulaSync fra før 3.2 (henvisningen Links\AulaSync.exe). Den hører til winget og røres ikke.
function WingetCopy: String;
begin
  Result := ExpandConstant('{localappdata}\Microsoft\WinGet\Links\AulaSync.exe');
  if FileExists(Result) then Exit;
  // Uden henvisning (winget kunne ikke lave den uden udviklertilstand) ligger den kun i pakkemappen.
  Result := ExpandConstant('{localappdata}\Microsoft\WinGet\Packages\rpaasch.AulaSync_Microsoft.Winget.Source_8wekyb3d8bbwe\AulaSync.exe');
  if not FileExists(Result) then Result := '';
end;

function InitializeSetup: Boolean;
begin
  FirstInstall := not RegKeyExists(HKCU, UninstallKey);
  Result := True;
end;

// Den nye AulaSync.exe pakkes ud midlertidigt, så den også kan lukke en AulaSync, der ligger et andet sted.
function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  WasRunning := AulaSyncRunning;
  if WasRunning then
  begin
    ExtractTemporaryFile('AulaSync.exe');
    StopAulaSync(ExpandConstant('{tmp}\AulaSync.exe'));
  end;
  Result := '';
end;

function LaunchSilently: Boolean;
begin
  Result := not Elevated and (WasRunning or (FirstInstall and (ExpandConstant('{param:LAUNCH|0}') = '1')));
end;

function LaunchParameters(Param: String): String;
begin
  if WasRunning then Result := '--silent' else Result := '';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Command: String;
begin
  if CurStep = ssPostInstall then
  begin
    Installed := True;
    // Var start ved login slået til, peger den nu på den installerede AulaSync.
    if RegQueryStringValue(HKCU, RunKey, 'AulaSync', Command) then
      RegWriteStringValue(HKCU, RunKey, 'AulaSync', '"' + AppExe + '" --silent');
  end;
end;

// Blev installationen afbrudt eller mislykkedes den, efter at AulaSync blev lukket, startes den igen (den gamle exe er
// rullet tilbage).
procedure DeinitializeSetup;
var
  Code: Integer;
begin
  if WasRunning and not Installed and not Elevated and FileExists(AppExe) then
    Exec(AppExe, '--silent', '', SW_SHOWNORMAL, ewNoWait, Code);
end;

// Alt fjernes i usUninstall, før Inno Setup selv fjerner filerne: winget venter kun på afinstallationsprogrammet, til
// Inno Setup har fjernet sine egne filer; usPostUninstall kommer først bagefter.
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  I: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    // Alle AulaSync'er lukkes, uanset hvor de ligger: de deler låsefilen og dataene.
    if AulaSyncRunning then StopAulaSync(AppExe);
    if CheckForMutexes(OldMutex) then
    begin
      Log('AulaSync 2 kører; den lukkes.');
      KillAulaSync;
    end;
    StopOurWebViews;
    // Uden at spørge: start ved login, brugerens data, AulaSync 2's mappe og .NET's udpakkede filer.
    OtherCopy := RemoveAutostart;
    if OtherCopy = '' then OtherCopy := WingetCopy;
    DataLeft := '';
    if not DeleteAll(DataDir) then DataLeft := DataDir;
    if not DeleteAll(OldDataDir) then DataLeft := Trim(DataLeft + #13#10 + OldDataDir);
    DeleteAll(ExtractDir);
    // Afinstallationen prøver kun én gang at slette filen; er den lige afsluttet, kan den være i brug et øjeblik endnu.
    I := 0;
    while FileExists(AppExe) and not DeleteFile(AppExe) and (I < 50) do
    begin
      Sleep(100);
      I := I + 1;
    end;
  end;
  // winget afinstallerer med /SILENT uden /SUPPRESSMSGBOXES; en MsgBox ville vente på et klik, så intet vises dér.
  if (CurUninstallStep = usPostUninstall) and not UninstallSilent then
  begin
    if DataLeft <> '' then MsgBox(FmtMessage(CustomMessage('DataLeft'), [DataLeft]), mbInformation, MB_OK);
    if OtherCopy <> '' then MsgBox(FmtMessage(CustomMessage('OtherCopy'), [OtherCopy]), mbInformation, MB_OK);
  end;
end;
