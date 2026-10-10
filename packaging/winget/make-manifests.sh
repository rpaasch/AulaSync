#!/usr/bin/env bash
# Skriver winget-manifesterne for rpaasch.AulaSync <version> til <ud-mappe>/manifests/... InstallerUrl peger på udgivelsen
# i <owner/repo> (fx rpaasch/AulaSync); SHA256 regnes ud fra filerne. Indsend mappen som pull request til microsoft/winget-pkgs.
#
# To installationsmåder:
# - inno: installationsprogrammet (AulaSync-Setup.exe) for brugeren selv, med AulaSync i Start-menuen. Nye installationer
#   får det (winget foretrækker inno frem for portable), og /LAUNCH=1 åbner AulaSync efter første installation.
# - portable: den løse AulaSync.exe, som før 3.2. winget kan ikke skifte en portabel installation til et
#   installationsprogram, så dem, der har den, bliver ved med at få den med winget upgrade. AulaSync lægger selv sin
#   genvej i Start-menuen.
#
# InstallerType, InstallerSwitches og Commands står øverst som i 3.1.1 (winget-pkgs' kontrol melder det som en
# uoverensstemmelse, hvis de mangler dér i forhold til den forrige version) og gælder den løse exe. Det, der står øverst,
# arver installationsprogrammet også, felt for felt, så det har sin egen type og Innos stille-parametre; ellers kørte
# winget det med --silent, som Inno ikke kender, og guiden kom frem. Scope må ikke stå øverst (portable har intet Scope).
#
#   make-manifests.sh 3.2.0 out/setup/AulaSync-Setup.exe out/win/AulaSync.exe rpaasch/AulaSync out/winget
set -euo pipefail

version=$1 setup=$2 exe=$3 repo=$4 out=$5
sha() { sha256sum "$1" | cut -d' ' -f1 | tr 'a-f' 'A-F'; }
dir="$out/manifests/r/rpaasch/AulaSync/$version"
url="https://github.com/$repo/releases/download/v$version"
mkdir -p "$dir"

cat > "$dir/rpaasch.AulaSync.yaml" <<YAML
# yaml-language-server: \$schema=https://aka.ms/winget-manifest.version.1.12.0.schema.json

PackageIdentifier: rpaasch.AulaSync
PackageVersion: $version
DefaultLocale: da-DK
ManifestType: version
ManifestVersion: 1.12.0
YAML

cat > "$dir/rpaasch.AulaSync.installer.yaml" <<YAML
# yaml-language-server: \$schema=https://aka.ms/winget-manifest.installer.1.12.0.schema.json

PackageIdentifier: rpaasch.AulaSync
PackageVersion: $version
InstallerType: portable
InstallerSwitches:
  Silent: --silent
  SilentWithProgress: --silent
Commands:
  - AulaSync
Installers:
  - Architecture: x64
    InstallerType: inno
    Scope: user
    InstallerUrl: $url/AulaSync-Setup.exe
    InstallerSha256: $(sha "$setup")
    InstallerSwitches:
      Silent: /SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
      SilentWithProgress: /SP- /SILENT /SUPPRESSMSGBOXES /NORESTART
      Custom: /LAUNCH=1
    UpgradeBehavior: install
    ProductCode: '{EB8DDE1A-F008-42A1-AA73-7110F88E4044}_is1'
  - Architecture: x64
    InstallerUrl: $url/AulaSync.exe
    InstallerSha256: $(sha "$exe")
ManifestType: installer
ManifestVersion: 1.12.0
YAML

cat > "$dir/rpaasch.AulaSync.locale.da-DK.yaml" <<YAML
# yaml-language-server: \$schema=https://aka.ms/winget-manifest.defaultLocale.1.12.0.schema.json

PackageIdentifier: rpaasch.AulaSync
PackageVersion: $version
PackageLocale: da-DK
Publisher: rpaasch
PublisherUrl: https://github.com/rpaasch
PackageName: AulaSync
PackageUrl: https://github.com/rpaasch/AulaSync
License: MIT
ShortDescription: Holder skemaer fra Aula opdateret i din kalender
Description: |-
  AulaSync henter skemaer fra Aula (medarbejdere, klasser og lokaler) og holder dem opdateret som kalendere i
  klassisk Outlook og andre kalenderprogrammer på samme computer. Ny Outlook og Outlook på nettet får skemaerne som
  import. AulaSync ligger i Start-menuen og kører i systembakken. Ikke tilknyttet eller godkendt af Aula, KOMBIT,
  Netcompany eller KMD. Fra version 3 synkroniseres beskeder ikke længere, og indstillinger fra 2.x overføres ikke.
Tags:
  - aula
  - outlook
  - skole
  - sync
  - education
  - kalender
  - ics
  - skema
ReleaseNotesUrl: https://github.com/$repo/releases/tag/v$version
ManifestType: defaultLocale
ManifestVersion: 1.12.0
YAML

echo "$dir"
