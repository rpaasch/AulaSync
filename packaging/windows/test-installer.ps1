# Prøver installationsprogrammet af på Windows (prøvebygningen kører det på GitHubs Windows-maskine):
#   1. Den gamle, løse AulaSync kører, og start ved login peger på den.
#   2. Installation uden vinduer, som winget gør det: den gamle lukkes, AulaSync ligger i Start-menuen, start ved login
#      peger på den installerede, og den installerede startes igen i systembakken.
#   3. Opdatering, mens den kører: den startes igen; --quit (som administrator, som et installationsprogram i et
#      administrator-vindue) lukker den pænt.
#   4. Afinstallation, som winget gør det ("/SILENT" uden /SUPPRESSMSGBOXES, som almindelig bruger): må ikke vente på et
#      spørgsmål; alt er væk, før Installerede apps er det: brugerens data, også login-mappen, mens en WebView2-proces
#      holder en fil i den åben, AulaSync 2's mappe, .NET's udpakkede filer og start ved login (også når den peger på
#      en løs AulaSync.exe, som ikke slettes, og i Windows' liste). Et andet programs WebView2 kører videre.
#   5. Første installation med /LAUNCH=1 (winget-manifestet): AulaSync åbner bagefter; afinstallation som administrator,
#      mens den kører, fjerner også alt.
#   6. "Afinstallér AulaSync…" for en løs AulaSync.exe (AulaSync.exe --uninstall, som knappen i Indstillinger starter):
#      den kørende AulaSync afslutter, og så fjernes exe'en, genvejen, start ved login (også i Windows' liste), brugerens
#      data, AulaSync 2's mappe og .NET's udpakkede filer.
#   7. Det samme for AulaSync fra winget fra før 3.2 (portabel, uden henvisning): winget's post under Installerede apps,
#      mappen i PATH og pakkemappen er også væk.
# Installationsprogrammet og AulaSync køres som almindelig bruger (runas /trustlevel), ligesom når winget kører i et
# almindeligt vindue; GitHubs maskine kører ellers alt som administrator.
#
#   pwsh packaging/windows/test-installer.ps1 out/setup/AulaSync-Setup.exe out/win/AulaSync.exe
param(
    [Parameter(Mandatory)][string]$Setup,
    [Parameter(Mandatory)][string]$Exe,
    [string]$OldVersion = '3.1.1'
)

$ErrorActionPreference = 'Stop'
$app = Join-Path $env:LOCALAPPDATA 'Programs\AulaSync\AulaSync.exe'
$lnk = Join-Path ([Environment]::GetFolderPath('Programs')) 'AulaSync.lnk'
$run = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{EB8DDE1A-F008-42A1-AA73-7110F88E4044}_is1'
$data = Join-Path $env:LOCALAPPDATA 'AulaSync'
$extracted = Join-Path ([IO.Path]::GetTempPath()) '.net\AulaSync' # som .NET: TMP før TEMP
$oldData = Join-Path $env:USERPROFILE '.aulasync'
$approved = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run'
$lock = Join-Path $data 'aulasync.lock'
# Brugerens egen midlertidige mappe, som den almindelige bruger kan skrive i; installationsprogrammet kopieres derhen.
$work = Join-Path $env:LOCALAPPDATA 'Temp\aulasync-test' # lang sti; TEMP kan være kort (C:\Users\RUNNER~1\…)
New-Item -ItemType Directory -Force $work | Out-Null
Copy-Item $Setup (Join-Path $work 'AulaSync-Setup.exe') -Force
$Setup = Join-Path $work 'AulaSync-Setup.exe'
$loose = Join-Path $work 'loes\AulaSync.exe'
New-Item -ItemType Directory -Force (Split-Path $loose) | Out-Null
Copy-Item $Exe $loose -Force
if ($work -match ' ') { throw "Stien må ikke indeholde mellemrum (runas): $work" }

function Check($ok, $text) {
    if (-not $ok) { throw "FEJL: $text" }
    Write-Host "ok: $text"
}
function Wait-Until([scriptblock]$condition, [int]$seconds, [string]$text) {
    $end = (Get-Date).AddSeconds($seconds)
    while (-not (& $condition)) {
        if ((Get-Date) -gt $end) { throw "FEJL: $text (ventede $seconds s)" }
        Start-Sleep -Milliseconds 500
    }
    Write-Host "ok: $text"
}
# Kører en AulaSync? Den holder låsefilen åben uden deling.
function Locked {
    if (-not (Test-Path $lock)) { return $false }
    try { [IO.File]::Open($lock, 'Open', 'ReadWrite', 'None').Dispose(); return $false }
    catch [IO.IOException] { return $true }
}
function Running($path) { [bool](Get-Process AulaSync -ErrorAction SilentlyContinue | Where-Object Path -eq $path) }
function Target($path) { (New-Object -ComObject WScript.Shell).CreateShortcut($path).TargetPath }
function RunValue { (Get-ItemProperty $run -ErrorAction SilentlyContinue).AulaSync }
# Som almindelig bruger, uden administrator.
function As-User([string]$commandLine) { & runas.exe /trustlevel:0x20000 $commandLine | Out-Null }
# msedgewebview2.exe-processer, hvis kommandolinje nævner mappen.
function WebViews([string]$dir) {
    @(Get-CimInstance Win32_Process -Filter "Name = 'msedgewebview2.exe'" |
        Where-Object { $_.CommandLine -and $_.CommandLine.IndexOf($dir, [StringComparison]::OrdinalIgnoreCase) -ge 0 })
}
# Som en WebView2-proces, der bliver hængende efter AulaSync: en kopi af Windows PowerShell med navnet msedgewebview2.exe
# holder selv en fil i $dir åben uden deling, og kommandolinjen nævner mappen, ligesom --user-data-dir. (Ikke cmd.exe med
# en omdirigering: så holder et underprogram filen, og det hedder ikke msedgewebview2.exe.)
$fakeWebView = Join-Path $work 'webview2\msedgewebview2.exe'
function Start-WebView([string]$dir) {
    New-Item -ItemType Directory -Force $dir | Out-Null
    As-User "$fakeWebView -NoProfile -NonInteractive -Command `$f=[IO.File]::Open('$dir\held.txt','OpenOrCreate','ReadWrite','None');Start-Sleep 600"
}
# Installation uden vinduer; venter på, at loggen er lukket, og kræver, at installationen lykkedes.
function Install([string]$extra = '') {
    $log = Join-Path $work "setup-$([guid]::NewGuid().ToString('N')).log"
    As-User "$Setup /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /LOG=$log $extra"
    Wait-Until { (Test-Path $log) -and ((Get-Content $log -Raw) -match 'Log closed\.') } 300 'installationsprogrammet er færdigt'
    $text = Get-Content $log -Raw
    if ($text -notmatch 'Installation process succeeded') { Write-Host $text; throw 'FEJL: installationen lykkedes ikke' }
}

try {
    if (-not (Test-Path $run)) { New-Item $run -Force | Out-Null }

    Write-Host "== 1. AulaSync $OldVersion kører fra en løs exe med start ved login"
    $old = Join-Path $work 'gammel\AulaSync.exe'
    New-Item -ItemType Directory -Force (Split-Path $old) | Out-Null
    Invoke-WebRequest "https://github.com/rpaasch/AulaSync/releases/download/v$OldVersion/AulaSync.exe" -OutFile $old
    Set-ItemProperty $run AulaSync "`"$old`" --silent"
    As-User "$old --silent"
    Wait-Until { Locked } 60 "AulaSync $OldVersion kører"

    Write-Host '== 2. Installation uden vinduer'
    Install
    Check (Test-Path $app) "AulaSync.exe ligger i $app"
    Check ((Target $lnk) -eq $app) 'genvejen i Start-menuen peger på den installerede AulaSync'
    Check ((RunValue) -eq "`"$app`" --silent") 'start ved login peger på den installerede AulaSync'
    Check ((Get-ItemProperty $uninstallKey).DisplayName -eq 'AulaSync') 'AulaSync står under Installerede apps'
    Check ((Get-ItemProperty $uninstallKey).DisplayIcon -eq $app) 'Installerede apps viser AulaSyncs ikon'
    Check (-not (Running $old)) "AulaSync $OldVersion er lukket"
    Wait-Until { Running $app } 30 'den installerede AulaSync er startet igen'
    Wait-Until { Locked } 60 'den installerede AulaSync kører'

    Write-Host '== 3. Opdatering, mens AulaSync kører'
    $marker = Join-Path $data 'kalendere\marker.txt' # som brugerens kalenderfiler
    New-Item -ItemType Directory -Force (Split-Path $marker) | Out-Null
    Set-Content $marker 'x'
    Install
    Check (Test-Path $marker) 'opdateringen bevarer brugerens data'
    Remove-Item $marker
    Wait-Until { (Running $app) -and (Locked) } 60 'AulaSync kører igen efter opdateringen'
    $q = Start-Process $app '--quit' -Wait -PassThru
    Check ($q.ExitCode -eq 2) "AulaSync.exe --quit lukkede den kørende AulaSync (fik $($q.ExitCode))"
    Check (-not (Locked)) 'låsen er fri'
    Wait-Until { -not (Running $app) } 30 'AulaSync er afsluttet'

    Write-Host '== 4. Afinstallation, som winget gør det'
    Check (Test-Path $extracted) ".NET har pakket filer ud i $extracted"
    # Login-mappen holdes åben af en WebView2-proces; et andet program har sin egen. AulaSync 2's mappe ligger der stadig.
    New-Item -ItemType Directory -Force (Split-Path $fakeWebView), (Join-Path $oldData 'webview2') | Out-Null
    Copy-Item "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" $fakeWebView -Force
    $ours = Join-Path $data 'webview\EBWebView'
    $other = Join-Path $work 'andet-program\EBWebView'
    Start-WebView $ours
    Start-WebView $other
    Wait-Until { (WebViews $ours).Count -eq 1 -and (WebViews $other).Count -eq 1 -and (Test-Path "$ours\held.txt") } 30 'en WebView2-proces holder en fil i login-mappen'
    # Start ved login peger på den løse AulaSync.exe og står i Windows' liste (Jobliste › Start).
    Set-ItemProperty $run AulaSync "`"$old`" --silent"
    if (-not (Test-Path $approved)) { New-Item $approved -Force | Out-Null }
    New-ItemProperty $approved AulaSync -PropertyType Binary -Value ([byte[]](2,0,0,0,0,0,0,0,0,0,0,0)) -Force | Out-Null
    $quiet = (Get-ItemProperty $uninstallKey).QuietUninstallString
    Check ($quiet -match '^"(\S+)" (.*)$') "afinstallation uden vinduer: $quiet"
    $ulog = Join-Path $work "setup-afinstallation-$([guid]::NewGuid().ToString('N')).log"
    As-User "$($Matches[1]) $($Matches[2]) /LOG=$ulog"
    Wait-Until { -not (Test-Path $app) -and -not (Test-Path $uninstallKey) } 120 'AulaSync er afinstalleret uden at spørge'
    # Loggen lukkes først, når afinstallationen er helt færdig: en boks, der venter på et klik, ville holde den åben.
    Wait-Until { (Test-Path $ulog) -and ((Get-Content $ulog -Raw) -match 'Log closed\.') } 60 'afinstallationen blev færdig uden at vente på et klik'
    # Alt andet fjernes i usUninstall, før Inno Setup fjerner sine egne filer og Installerede apps (og før winget er færdig).
    Check (-not (Test-Path $data)) 'brugerens data er væk, også login-mappen'
    Check ((WebViews $ours).Count -eq 0) 'WebView2-processen i login-mappen er lukket'
    Check ((WebViews $other).Count -eq 1) 'det andet programs WebView2 kører stadig'
    Check (-not (Test-Path $oldData)) "AulaSync 2's mappe er væk"
    Check (-not (Test-Path $extracted)) '.NET-filerne er væk'
    Check ($null -eq (RunValue)) 'start ved login er væk, også når den pegede på den løse AulaSync'
    Check ($null -eq (Get-ItemProperty $approved -ErrorAction SilentlyContinue).AulaSync) "start ved login er væk fra Windows' liste"
    Check (Test-Path $old) 'den løse AulaSync.exe er ikke slettet'
    # Afinstallationsprogrammet sletter sig selv og mappen til sidst, efter nøglen under Installerede apps.
    Wait-Until { -not (Test-Path (Split-Path $app)) } 30 'mappen er væk'
    Check (-not (Test-Path $lnk)) 'genvejen i Start-menuen er væk'

    Write-Host '== 5. Første installation med /LAUNCH=1'
    Install '/LAUNCH=1'
    Wait-Until { (Running $app) -and (Locked) } 60 'AulaSync er åbnet efter installationen'
    $quiet = (Get-ItemProperty $uninstallKey).QuietUninstallString
    $null = $quiet -match '^"(.+)" (.*)$'
    Start-Process $Matches[1] $Matches[2] | Out-Null
    Wait-Until { -not (Test-Path $app) -and -not (Test-Path $uninstallKey) } 120 'AulaSync er afinstalleret, mens den kørte'
    Check (-not (Running $app)) 'AulaSync er lukket'
    Check (-not (Test-Path $data)) 'brugerens data er væk'
    Wait-Until { -not (Test-Path (Split-Path $app)) } 30 'mappen er væk'

    Write-Host '== 6. Afinstallér AulaSync… for en løs AulaSync.exe'
    As-User "$loose --silent"
    Wait-Until { (Running $loose) -and (Locked) } 60 'den løse AulaSync kører'
    Set-ItemProperty $run AulaSync "`"$loose`" --silent"
    Wait-Until { (Test-Path $lnk) -and ((Target $lnk) -eq $loose) } 30 'genvejen i Start-menuen peger på den løse AulaSync'
    New-Item -ItemType Directory -Force (Join-Path $oldData 'webview2') | Out-Null
    if (-not (Test-Path $approved)) { New-Item $approved -Force | Out-Null }
    New-ItemProperty $approved AulaSync -PropertyType Binary -Value ([byte[]](2,0,0,0,0,0,0,0,0,0,0,0)) -Force | Out-Null
    Check (Test-Path $extracted) ".NET har pakket filer ud i $extracted"
    As-User "$loose --uninstall"
    Wait-Until { -not (Test-Path $loose) } 90 'den løse AulaSync.exe er slettet'
    Check (-not (Running $loose)) 'den løse AulaSync er afsluttet'
    Check (-not (Test-Path $data)) 'brugerens data er væk'
    Check (-not (Test-Path $lnk)) 'genvejen i Start-menuen er væk'
    Check ($null -eq (RunValue)) 'start ved login er væk'
    Check ($null -eq (Get-ItemProperty $approved -ErrorAction SilentlyContinue).AulaSync) "start ved login er væk fra Windows' liste"
    Check (-not (Test-Path $oldData)) "AulaSync 2's mappe er væk"
    Wait-Until { -not (Test-Path $extracted) } 30 '.NET-filerne er væk'

    Write-Host '== 7. Afinstallér AulaSync… for AulaSync fra winget fra før 3.2'
    $package = Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages\rpaasch.AulaSync_Microsoft.Winget.Source_8wekyb3d8bbwe'
    $wingetExe = Join-Path $package 'AulaSync.exe'
    $wingetKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\rpaasch.AulaSync_Microsoft.Winget.Source_8wekyb3d8bbwe'
    New-Item -ItemType Directory -Force $package | Out-Null
    Copy-Item $Exe $wingetExe -Force
    # Som winget uden udviklertilstand: ingen henvisning i Links, men mappen i brugerens PATH.
    New-Item $wingetKey -Force | Out-Null
    New-ItemProperty $wingetKey DisplayName -Value 'AulaSync' -Force | Out-Null
    New-ItemProperty $wingetKey InstallDirectoryAddedToPath -PropertyType DWord -Value 1 -Force | Out-Null
    $envKey = Get-Item 'HKCU:\Environment'
    $userPath = $envKey.GetValue('Path', '', 'DoNotExpandEnvironmentNames')
    $pathKind = if ($envKey.GetValueNames() -contains 'Path') { $envKey.GetValueKind('Path') } else { 'ExpandString' }
    Set-ItemProperty 'HKCU:\Environment' Path (($userPath, $package | Where-Object { $_ }) -join ';') -Type $pathKind
    As-User "$wingetExe --silent"
    Wait-Until { (Running $wingetExe) -and (Locked) } 60 'AulaSync fra winget kører'
    As-User "$wingetExe --uninstall"
    Wait-Until { -not (Test-Path $package) } 90 "winget's pakkemappe er slettet"
    Check (-not (Test-Path $wingetKey)) "winget's post under Installerede apps er væk"
    $after = (Get-Item 'HKCU:\Environment').GetValue('Path', '', 'DoNotExpandEnvironmentNames')
    Check (-not ($after -split ';' | Where-Object { $_.TrimEnd('\') -eq $package })) 'mappen er væk fra PATH'
    Check ($after -eq $userPath) 'resten af PATH er uændret'
    Check (-not (Test-Path $data)) 'brugerens data er væk'
}
catch {
    $log = Join-Path $data 'aulasync.log'
    if (Test-Path $log) { Write-Host '--- aulasync.log'; Get-Content $log -Tail 40 }
    Get-ChildItem $work -Filter 'setup-*.log' -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "--- $($_.Name)"; Get-Content $_.FullName -Tail 40 }
    throw
}
finally {
    Get-Process AulaSync -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Get-Process msedgewebview2 -ErrorAction SilentlyContinue | Where-Object Path -eq $fakeWebView |
        Stop-Process -Force -ErrorAction SilentlyContinue
}
