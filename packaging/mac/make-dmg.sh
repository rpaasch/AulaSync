#!/usr/bin/env bash
# Laver AulaSync.app og AulaSync-<version>-<arm64|x64>.dmg ud fra en udgivet osx-arm64- eller osx-x64-mappe.
# Appen er kun ad hoc-signeret (ingen Apple-konto endnu): første gang skal brugeren vælge "Åbn alligevel" i
# Systemindstillinger › Anonymitet og sikkerhed. Kører på macOS (codesign, hdiutil); andre steder bygges kun app'en.
#
#   make-dmg.sh osx-arm64 out/osx-arm64 3.0.0 out
set -euo pipefail

rid=$1 publish=$2 version=$3 out=$4
arch=${rid#osx-}
here=$(cd "$(dirname "$0")" && pwd)
stage=$(mktemp -d)
app="$stage/AulaSync.app"

mkdir -p "$app/Contents/MacOS" "$app/Contents/Resources"
cp -R "$publish"/. "$app/Contents/MacOS/"
rm -f "$app/Contents/MacOS/"*.pdb
sed "s/@VERSION@/$version/g" "$here/Info.plist" > "$app/Contents/Info.plist"
cp "$here/../icon/AulaSync.icns" "$app/Contents/Resources/AulaSync.icns"
# macOS 26 og nyere viser kun ikoner i Apples nye format pænt (ellers på en grå plade): Assets.car oversættes fra
# AulaSync.icon med actool (Xcode 26). Mangler actool, eller fejler den, bruges kun AulaSync.icns.
if command -v xcrun > /dev/null && xcrun --find actool > /dev/null 2>&1; then
  car=$(mktemp -d)
  if xcrun actool "$here/../icon/AulaSync.icon" --compile "$car" --app-icon AulaSync --include-all-app-icons \
       --target-device mac --platform macosx --minimum-deployment-target 15.0 --enable-on-demand-resources NO \
       --output-partial-info-plist "$car/partial.plist" --errors --warnings > "$car/actool.log" 2>&1 && [ -s "$car/Assets.car" ]; then
    cp "$car/Assets.car" "$app/Contents/Resources/Assets.car"
    /usr/libexec/PlistBuddy -c "Add :CFBundleIconName string AulaSync" "$app/Contents/Info.plist"
  else
    echo "::warning::actool kunne ikke lave Assets.car af AulaSync.icon; kun AulaSync.icns bruges"
    cat "$car/actool.log"
  fi
  rm -rf "$car"
fi
chmod +x "$app/Contents/MacOS/AulaSync"

mkdir -p "$out"
if command -v codesign > /dev/null; then
  codesign --force --deep --sign - "$app"
  codesign --verify --deep --strict "$app"
fi
# Træk app'en over i Programmer (genvej til /Applications i diskbilledet)
ln -s /Applications "$stage/Programmer"
if command -v hdiutil > /dev/null; then
  hdiutil create -volname "AulaSync $version" -srcfolder "$stage" -ov -format UDZO "$out/AulaSync-$version-$arch.dmg"
else
  cp -R "$app" "$out/AulaSync-$version-$arch.app"
  echo "hdiutil findes kun på macOS; app'en ligger i $out/AulaSync-$version-$arch.app"
fi
rm -rf "$stage"
