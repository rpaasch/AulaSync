#!/usr/bin/env bash
# Skriver winget-manifesterne for rpaasch.AulaSync <version> (portabel exe, som i 2.x) til <ud-mappe>/manifests/...
# InstallerUrl peger på udgivelsen i <owner/repo> (fx rpaasch/AulaSync); SHA256 regnes ud fra exe'en.
# Indsend mappen som pull request til microsoft/winget-pkgs.
#
#   make-manifests.sh 3.0.0 out/win/AulaSync.exe rpaasch/AulaSync out/winget
set -euo pipefail

version=$1 exe=$2 repo=$3 out=$4
sha=$(sha256sum "$exe" | cut -d' ' -f1 | tr 'a-f' 'A-F')
dir="$out/manifests/r/rpaasch/AulaSync/$version"
mkdir -p "$dir"

cat > "$dir/rpaasch.AulaSync.yaml" <<YAML
# yaml-language-server: \$schema=https://aka.ms/winget-manifest.version.1.6.0.schema.json

PackageIdentifier: rpaasch.AulaSync
PackageVersion: $version
DefaultLocale: da-DK
ManifestType: version
ManifestVersion: 1.6.0
YAML

cat > "$dir/rpaasch.AulaSync.installer.yaml" <<YAML
# yaml-language-server: \$schema=https://aka.ms/winget-manifest.installer.1.6.0.schema.json

PackageIdentifier: rpaasch.AulaSync
PackageVersion: $version
InstallerType: portable
Commands:
  - AulaSync
InstallerSwitches:
  Silent: --silent
  SilentWithProgress: --silent
Installers:
  - Architecture: x64
    InstallerUrl: https://github.com/$repo/releases/download/v$version/AulaSync.exe
    InstallerSha256: $sha
ManifestType: installer
ManifestVersion: 1.6.0
YAML

cat > "$dir/rpaasch.AulaSync.locale.da-DK.yaml" <<YAML
# yaml-language-server: \$schema=https://aka.ms/winget-manifest.defaultLocale.1.6.0.schema.json

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
  import. AulaSync kører i systembakken. Ikke tilknyttet eller godkendt af Aula, KOMBIT, Netcompany eller KMD.
  Fra version 3 synkroniseres beskeder ikke længere, og indstillinger fra 2.x overføres ikke.
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
ManifestVersion: 1.6.0
YAML

echo "$dir"
