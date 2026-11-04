# Signerer AulaSync til Windows med et kodesigneringscertifikat i Windows' certifikatlager (Certum SimplySign eller, i
# prøven på GitHub, et selvsigneret certifikat): AulaSync.exe signeres, installationsprogrammet bygges med Inno Setup ud
# fra den signerede exe, og Inno signerer selv installations- og afinstallationsprogrammet. Skriver
# <Out>\win\AulaSync.exe og <Out>\setup\AulaSync-Setup.exe og tjekker, at alle signaturer er gyldige og fra certifikatet.
# Virker i Windows PowerShell 5.1 og PowerShell 7.
#
#   sign-windows.ps1 -Exe out\win\AulaSync.exe -Version 3.2.2 -Out out\signed -Thumbprint <aftryk>
#   sign-windows.ps1 -Test -Thumbprint <aftryk>       (prøvesignering af en kopi af signtool.exe; intet bygges)
#
# signtool.exe hentes fra NuGet-pakken Microsoft.Windows.SDK.BuildTools (fast version, Microsoft-signeret) til -ToolDir,
# medmindre -SignTool er angivet. Inno Setup (ISCC.exe) findes de vanlige steder, eller angives med -Iscc. Certifikatets
# navn skrives aldrig ud, kun aftrykket og udløbsdatoen.
[CmdletBinding(DefaultParameterSetName = 'Build')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Build')][string]$Exe,
    [Parameter(Mandatory, ParameterSetName = 'Build')][string]$Version,
    [Parameter(Mandatory, ParameterSetName = 'Build')][string]$Out,
    [Parameter(Mandatory, ParameterSetName = 'Test')][switch]$Test,
    [Parameter(Mandatory)][string]$Thumbprint,
    [string]$TimestampUrl = 'http://time.certum.pl',
    [string]$SignTool,
    [string]$ToolDir = (Join-Path $env:LOCALAPPDATA 'AulaSync-udgivelse\tools'),
    [string]$Iscc,
    [string]$Iss
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue' # ellers henter Invoke-WebRequest i Windows PowerShell meget langsomt
$CodeSigning = '1.3.6.1.5.5.7.3.3'
# signtool fra Windows SDK Build Tools; pakkens SHA-512 (fra nuget.org) tjekkes, før den pakkes ud.
$BuildToolsVersion = '10.0.28000.2705'
$BuildToolsSha512 = '47kp42eMb8rF3bjDuZHVkzGRb9Fr7wyFZDIT2sWr3StwnTCoDHIisDe0SPWGwerwVUw/eQQO/OxDb0YYHbfNvA=='

function Fail($text) { throw $text }

# Certifikatet med privat nøgle og formålet kodesignering, gyldigt i mindst et døgn endnu.
function Get-SigningCertificate([string]$thumb) {
    $thumb = ($thumb -replace '\s', '').ToUpperInvariant()
    $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $thumb } | Select-Object -First 1
    if (-not $cert) { Fail "Certifikatet $thumb findes ikke i dit certifikatlager. Er du logget ind i SimplySign Desktop?" }
    if (-not $cert.HasPrivateKey) { Fail "Certifikatet $thumb har ingen privat nøgle her. Log ind i SimplySign Desktop, og prøv igen." }
    if (-not ($cert.EnhancedKeyUsageList | Where-Object { $_.ObjectId -eq $CodeSigning })) { Fail "Certifikatet $thumb kan ikke bruges til kodesignering." }
    if ($cert.NotAfter -lt (Get-Date).AddDays(1)) { Fail "Certifikatet $thumb udløber $($cert.NotAfter.ToString('yyyy-MM-dd')). Forny det hos Certum." }
    if ($cert.NotAfter -lt (Get-Date).AddDays(30)) { Write-Warning "Certifikatet udløber $($cert.NotAfter.ToString('yyyy-MM-dd')). Forny det hos Certum i god tid." }
    Write-Host "Certifikat: $thumb, gyldigt til $($cert.NotAfter.ToString('yyyy-MM-dd'))"
    return $cert
}

# signtool.exe: den angivne, eller den fra NuGet-pakken i ToolDir (hentes første gang). Skal være signeret af Microsoft.
function Get-SignTool {
    if ($SignTool) { $path = $SignTool }
    else {
        $dir = Join-Path $ToolDir "Microsoft.Windows.SDK.BuildTools.$BuildToolsVersion"
        $path = Get-ChildItem (Join-Path $dir 'bin') -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
            Where-Object { $_.Directory.Name -eq 'x64' } | Select-Object -First 1 -ExpandProperty FullName
        if (-not $path) {
            Write-Host "Henter signtool (Microsoft.Windows.SDK.BuildTools $BuildToolsVersion)..."
            $zip = Join-Path ([IO.Path]::GetTempPath()) "buildtools-$BuildToolsVersion-$PID.zip"
            $old = [Net.ServicePointManager]::SecurityProtocol
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Invoke-WebRequest "https://www.nuget.org/api/v2/package/Microsoft.Windows.SDK.BuildTools/$BuildToolsVersion" -OutFile $zip -UseBasicParsing
            }
            finally { [Net.ServicePointManager]::SecurityProtocol = $old }
            $sha = [Convert]::ToBase64String([Security.Cryptography.SHA512]::Create().ComputeHash([IO.File]::ReadAllBytes($zip)))
            if ($sha -ne $BuildToolsSha512) { Remove-Item $zip -Force; Fail 'Den hentede signtool-pakke har en anden kontrolsum end forventet. Den bruges ikke.' }
            if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
            Expand-Archive $zip $dir
            Remove-Item $zip -Force
            $path = Get-ChildItem (Join-Path $dir 'bin') -Recurse -Filter signtool.exe | Where-Object { $_.Directory.Name -eq 'x64' } |
                Select-Object -First 1 -ExpandProperty FullName
            if (-not $path) { Fail 'signtool.exe (x64) findes ikke i Microsoft.Windows.SDK.BuildTools-pakken.' }
        }
    }
    if (-not (Test-Path $path)) { Fail "signtool findes ikke: $path" }
    $sig = Get-AuthenticodeSignature $path
    if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') { Fail "signtool er ikke signeret af Microsoft: $path" }
    return $path
}

# ISCC.exe: den angivne, i PATH, hvor Inno Setup 7 eller 6 er installeret (Installerede apps), eller de vanlige steder.
function Get-Iscc {
    if ($Iscc) { if (Test-Path $Iscc) { return $Iscc } else { Fail "ISCC.exe findes ikke: $Iscc" } }
    $cmd = Get-Command iscc -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $keys = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    foreach ($name in 'Inno Setup 7_is1', 'Inno Setup 6_is1') {
        foreach ($key in $keys) {
            $dir = (Get-ItemProperty (Join-Path $key $name) -ErrorAction SilentlyContinue).InstallLocation
            if ($dir -and (Test-Path (Join-Path $dir 'ISCC.exe'))) { return (Join-Path $dir 'ISCC.exe') }
        }
    }
    foreach ($base in @(${env:ProgramFiles(x86)}, $env:ProgramFiles, (Join-Path $env:LOCALAPPDATA 'Programs'))) {
        if (-not $base) { continue }
        foreach ($name in 'Inno Setup 7', 'Inno Setup 7 (32-bit)', 'Inno Setup 6') {
            $p = Join-Path $base "$name\ISCC.exe"
            if (Test-Path $p) { return $p }
        }
    }
    Fail 'Inno Setup findes ikke. Installér det med:  winget install -e --id JRSoftware.InnoSetup --scope user'
}

# Signerer én fil med signtool (tidsstempel fra TimestampUrl; tidsstempeltjenesten kan svigte et øjeblik, så tre forsøg).
function Sign-File([string]$tool, [string]$thumb, [string]$file) {
    $a = @('sign', '/sha1', $thumb, '/fd', 'sha256')
    if ($TimestampUrl) { $a += @('/tr', $TimestampUrl, '/td', 'sha256') }
    $a += @('/q', $file)
    for ($i = 1; ; $i++) {
        & $tool @a
        if ($LASTEXITCODE -eq 0) { return }
        if ($i -ge 3) { Fail "signtool kunne ikke signere $(Split-Path $file -Leaf) (se fejlen ovenfor)." }
        Write-Host "signtool fejlede; prøver igen om $(5 * $i) sekunder..."
        Start-Sleep (5 * $i)
    }
}

# Gyldig signatur fra certifikatet (og med tidsstempel, når der signeres med et).
function Assert-Signed([string]$file, [string]$thumb) {
    $sig = Get-AuthenticodeSignature $file
    $name = Split-Path $file -Leaf
    if ($sig.Status -ne 'Valid') { Fail "$name er ikke gyldigt signeret ($($sig.Status): $($sig.StatusMessage))." }
    if ($sig.SignerCertificate.Thumbprint -ne $thumb) { Fail "$name er signeret med et andet certifikat ($($sig.SignerCertificate.Thumbprint))." }
    if ($TimestampUrl -and -not $sig.TimeStamperCertificate) { Fail "$name har intet tidsstempel." }
    Write-Host "ok: $name er signeret"
}

$thumb = ($Thumbprint -replace '\s', '').ToUpperInvariant()
$null = Get-SigningCertificate $thumb
$tool = Get-SignTool

if ($Test) {
    # En kopi af signtool.exe signeres om (erstatter Microsofts signatur i kopien); det viser, at certifikatet, SimplySign
    # og tidsstemplet virker, uden at noget bygges.
    $copy = Join-Path ([IO.Path]::GetTempPath()) "aulasync-proevesignering-$PID.exe"
    Copy-Item $tool $copy -Force
    try {
        Sign-File $tool $thumb $copy
        Assert-Signed $copy $thumb
    }
    finally { Remove-Item $copy -Force -ErrorAction SilentlyContinue }
    Write-Host 'Signeringen virker.' -ForegroundColor Green
    return
}

if (-not (Test-Path $Exe)) { Fail "AulaSync.exe findes ikke: $Exe" }
$fileVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo((Resolve-Path $Exe).Path).FileVersion
if ($fileVersion -notmatch "^$([regex]::Escape($Version))(\.0)?$") { Fail "AulaSync.exe har version $fileVersion, ikke $Version." }
if ((Get-AuthenticodeSignature $Exe).Status -ne 'NotSigned') { Fail 'AulaSync.exe er allerede signeret. Brug den usignerede exe fra GitHub-bygningen.' }
if (-not $Iss) { $Iss = Join-Path $PSScriptRoot 'AulaSync.iss' }
if (-not (Test-Path $Iss)) { Fail "AulaSync.iss findes ikke: $Iss" }
$isccPath = Get-Iscc

$winDir = Join-Path $Out 'win'
$setupDir = Join-Path $Out 'setup'
foreach ($d in $winDir, $setupDir) { if (Test-Path $d) { Remove-Item $d -Recurse -Force }; New-Item -ItemType Directory -Force $d | Out-Null }
$signedExe = Join-Path (Resolve-Path $winDir).Path 'AulaSync.exe'
Copy-Item $Exe $signedExe -Force

Sign-File $tool $thumb $signedExe
Assert-Signed $signedExe $thumb

# Inno Setup signerer afinstallationsprogrammet og installationsprogrammet med samme kommando. $q er et anførselstegn og
# $f filen (Innos egne pladsholdere; enkelte anførselstegn, så PowerShell ikke udfylder dem).
$innoSign = '$q' + $tool + '$q sign /sha1 ' + $thumb + ' /fd sha256'
if ($TimestampUrl) { $innoSign += ' /tr ' + $TimestampUrl + ' /td sha256' }
$innoSign += ' /q $f'
# Uden /Q, så en fejl fra signtool (SimplySign eller tidsstemplet) står i Innos udskrift.
& $isccPath "/DVersion=$Version" "/DSource=$((Resolve-Path $winDir).Path)" "/O$((Resolve-Path $setupDir).Path)" /DSign "/Saulasync=$innoSign" $Iss
if ($LASTEXITCODE -ne 0) { Fail "Inno Setup kunne ikke bygge det signerede installationsprogram (exit $LASTEXITCODE)." }
$setup = Join-Path (Resolve-Path $setupDir).Path 'AulaSync-Setup.exe'
Assert-Signed $setup $thumb

foreach ($f in $signedExe, $setup) {
    Write-Host ("{0}  {1}" -f (Get-FileHash $f -Algorithm SHA256).Hash, (Split-Path $f -Leaf))
}
Write-Host 'Windows-filerne er signeret.' -ForegroundColor Green
