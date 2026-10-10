#!/usr/bin/env bash
# Laver AulaSync.app og AulaSync-<version>-<arm64|x64>.dmg ud fra en udgivet osx-arm64- eller osx-x64-mappe. Kører på
# macOS (codesign, notarytool, hdiutil); andre steder bygges kun app'en.
#
# Signering:
# - MAC_SIGN_IDENTITY (SHA-1 for en "Developer ID Application"-identitet; release.yml finder den ved kørslen i en midlertidig
#   nøglering, MAC_KEYCHAIN): Developer ID med hardened runtime og tidsstempel, indefra og ud uden --deep: hver fil i
#   Contents/MacOS (codesign regner også .dll og .json dér for kode og gemmer deres signatur i udvidede attributter, som
#   hdiutil, ditto og Homebrew bevarer), så app'en med AulaSync.entitlements. App'en notariseres (ditto-zip) og hæftes, før
#   den lægges i dmg'en, så også Homebrews kopi har billetten (også uden net); derefter signeres, notariseres og hæftes
#   dmg'en. Så åbner macOS AulaSync uden "Åbn alligevel".
#   Notarisering: NOTARY_KEY (.p8-fil) + NOTARY_KEY_ID (+ NOTARY_ISSUER for en teamnøgle) fra App Store Connect, eller
#   NOTARY_APPLE_ID + NOTARY_TEAM_ID + NOTARY_PASSWORD (app-specifik adgangskode), eller NOTARY_KEYCHAIN_PROFILE
#   (xcrun notarytool store-credentials, på egen Mac). NOTARY_TIMEOUT (standard 30m) er ventetiden pr. indsendelse.
#   Notar-loggene lægges i <ud-mappe>/notary/.
# - Ellers kun ad hoc-signeret som før (forks og prøvebygninger uden Apples nøgler): første gang "Åbn alligevel".
#
#   make-dmg.sh osx-arm64 out/osx-arm64 3.0.0 out
set -euo pipefail

rid=$1 publish=$2 version=$3 out=$4
arch=${rid#osx-}
here=$(cd "$(dirname "$0")" && pwd)
stage=$(mktemp -d)
work=$(mktemp -d)
trap 'rm -rf "$stage" "$work"' EXIT
app="$stage/AulaSync.app"
dmg="$out/AulaSync-$version-$arch.dmg"

# Bash 3.2 (macOS' egen) med set -u: tomme arrays kun som ${a[@]+"${a[@]}"}.
identity=${MAC_SIGN_IDENTITY:-}
keychain=()
[ -n "${MAC_KEYCHAIN:-}" ] && keychain=(--keychain "$MAC_KEYCHAIN")
notary=() notarize_on=0
if [ -n "${NOTARY_KEY:-}" ]; then
  notary=(--key "$NOTARY_KEY" --key-id "${NOTARY_KEY_ID:?NOTARY_KEY_ID mangler}") notarize_on=1
  # Kun en teamnøgle har en udsteder; en personlig nøgle afvises (401), hvis den får en.
  [ -n "${NOTARY_ISSUER:-}" ] && notary=("${notary[@]}" --issuer "$NOTARY_ISSUER")
elif [ -n "${NOTARY_APPLE_ID:-}" ]; then
  notary=(--apple-id "$NOTARY_APPLE_ID" --team-id "${NOTARY_TEAM_ID:?NOTARY_TEAM_ID mangler}"
          --password "${NOTARY_PASSWORD:?NOTARY_PASSWORD mangler}") notarize_on=1
elif [ -n "${NOTARY_KEYCHAIN_PROFILE:-}" ]; then
  notary=(--keychain-profile "$NOTARY_KEYCHAIN_PROFILE") notarize_on=1
fi
if [ -n "$identity" ] && [ "$notarize_on" = 0 ]; then
  echo "::error::Developer ID uden notarisering: macOS ville stadig stoppe AulaSync. Angiv NOTARY_KEY, NOTARY_APPLE_ID eller NOTARY_KEYCHAIN_PROFILE."
  exit 1
fi
if [ -z "$identity" ] && [ "$notarize_on" = 1 ]; then echo "::error::Notarisering kræver MAC_SIGN_IDENTITY (Developer ID)"; exit 1; fi

# codesign med tre forsøg: tidsstempeltjenesten (timestamp.apple.com) svarer af og til ikke.
sign() {
  local n
  for n in 1 2 3; do
    codesign "$@" && return 0
    echo "::warning::codesign fejlede (forsøg $n)"; sleep $((n * 10))
  done
  return 1
}

# Sender en fil til Apples notartjeneste og venter på svaret. Loggen hentes altid (også ved succes: advarsler).
notarize() {
  local file=$1 name json sub status rc=0
  name=$(basename "$file")
  mkdir -p "$out/notary"
  echo "Notariserer $name ..."
  json=$(xcrun notarytool submit "$file" "${notary[@]}" --wait --timeout "${NOTARY_TIMEOUT:-30m}" --output-format json) || rc=$?
  sub=$(jq -r '.id // empty' <<< "$json" 2> /dev/null || true)
  status=$(jq -r '.status // empty' <<< "$json" 2> /dev/null || true)
  echo "$name: id=${sub:-?} status=${status:-?} (notarytool: $rc)"
  if [ -n "$sub" ]; then
    if xcrun notarytool log "$sub" "${notary[@]}" "$out/notary/$name.json" > /dev/null 2>&1; then
      jq '{status, statusSummary, issues}' "$out/notary/$name.json"
    else
      echo "::warning::Notar-loggen for $name ($sub) kunne ikke hentes endnu"
    fi
  else
    echo "$json"
  fi
  if [ "$rc" -ne 0 ] || [ "$status" != Accepted ]; then
    if [ "$status" = "In Progress" ]; then
      echo "::error::$name er ikke notariseret efter ${NOTARY_TIMEOUT:-30m}; Apple arbejder videre på $sub. Kør jobbet igen senere."
    else
      echo "::error::Notariseringen af $name fejlede (${status:-intet svar}); se loggen ovenfor og artefaktet notary-logs"
    fi
    return 1
  fi
}

# Hæfter billetten på (app eller dmg). Lige efter "Accepted" kan billetten mangle et øjeblik hos Apple.
staple() {
  local n
  for n in 1 2 3; do
    xcrun stapler staple "$1" && xcrun stapler validate "$1" && return 0
    echo "::warning::stapler fejlede for $(basename "$1") (forsøg $n)"; sleep $((n * 20))
  done
  return 1
}

# Gatekeepers vurdering, som brugeren møder den. Kun ét -v, så navnet i certifikatet (origin=) ikke står i loggen. Er
# Gatekeeper slået fra på maskinen, kan spctl ikke sige "Notarized Developer ID"; så kun en advarsel.
gatekeeper() {
  local o rc=0
  if grep -q 'assessments disabled' <<< "$(spctl --status 2>&1 || true)"; then
    echo "::warning::Gatekeeper er slået fra på denne maskine; spctl-tjekket springes over"
    return 0
  fi
  o=$(spctl "$@" 2>&1) || rc=$?
  grep -v '^origin=' <<< "$o" || true
  [ "$rc" -eq 0 ] && grep -q '^source=Notarized Developer ID' <<< "$o"
}

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
  if [ -n "$identity" ]; then
    # Indefra og ud, uden --deep: hver fil i Contents/MacOS undtagen hovedprogrammet (hardened runtime, ingen
    # rettigheder), så app'en selv (hovedprogrammet med AulaSync.entitlements og forseglingen af resten).
    plutil -lint "$here/AulaSync.entitlements" > /dev/null
    while IFS= read -r -d '' f; do
      sign --force --timestamp --options runtime ${keychain[@]+"${keychain[@]}"} --sign "$identity" "$f"
    done < <(find "$app/Contents/MacOS" -type f ! -path "$app/Contents/MacOS/AulaSync" -print0)
    sign --force --timestamp --options runtime --entitlements "$here/AulaSync.entitlements" \
      ${keychain[@]+"${keychain[@]}"} --sign "$identity" "$app"
  else
    codesign --force --deep --sign - "$app"
  fi
  codesign --verify --deep --strict "$app"
fi

if [ -n "$identity" ]; then
  ditto -c -k --sequesterRsrc --keepParent "$app" "$work/AulaSync-$version-$arch.zip"
  notarize "$work/AulaSync-$version-$arch.zip"
  staple "$app"
fi

# Træk app'en over i Programmer (genvej til /Applications i diskbilledet)
ln -s /Applications "$stage/Programmer"
if command -v hdiutil > /dev/null; then
  # GitHubs Mac-maskiner siger af og til "Resource busy"; derfor op til tre forsøg.
  for n in 1 2 3; do
    hdiutil create -volname "AulaSync $version" -srcfolder "$stage" -ov -format UDZO "$dmg" && break
    [ "$n" = 3 ] && exit 1
    sleep 10
  done
  if [ -n "$identity" ]; then
    sign --force --timestamp ${keychain[@]+"${keychain[@]}"} --sign "$identity" -i dk.rpaasch.aulasync.dmg "$dmg"
    codesign --verify --strict "$dmg"
    notarize "$dmg"
    staple "$dmg"
    gatekeeper -a -t exec -v "$app" || { echo "::error::Gatekeeper godkender ikke AulaSync.app ($arch)"; exit 1; }
    gatekeeper -a -t open --context context:primary-signature -v "$dmg" \
      || { echo "::error::Gatekeeper godkender ikke $(basename "$dmg")"; exit 1; }
  fi
else
  cp -R "$app" "$out/AulaSync-$version-$arch.app"
  echo "hdiutil findes kun på macOS; app'en ligger i $out/AulaSync-$version-$arch.app"
fi
