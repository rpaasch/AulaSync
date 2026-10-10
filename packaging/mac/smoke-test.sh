#!/usr/bin/env bash
# Røgprøve af AulaSync.app fra en dmg på en Mac (release.yml): .NET starter (AulaSync --quit), menulinjen starter
# (--silent) og lukker igen (--quit), og Afinstallér AulaSync… (--uninstall) rydder op og lægger app'en i papirkurven.
#
# --hardened: er app'en kun ad hoc-signeret (bygget uden Apples nøgler), signeres en kopi først ad hoc med hardened runtime
# og rettighederne fra AulaSync.entitlements plus disable-library-validation (en ad hoc-signatur har intet team, så macOS
# ville ellers ikke indlæse .NET's biblioteker). Så prøves det også uden nøglerne, at .NET kører under hardened runtime.
# En app med Developer ID prøves, som den er (også bibliotekskontrollen).
#
#   smoke-test.sh out/dmg/AulaSync-3.2.1-arm64.dmg [--hardened]
set -euo pipefail

dmg=$1 hardened=${2:-}
here=$(cd "$(dirname "$0")" && pwd)
work=$(mktemp -d)
mnt=$(mktemp -d)
trap 'hdiutil detach "$mnt" > /dev/null 2>&1 || true; rm -rf "$work" "$mnt"' EXIT
app="$work/AulaSync.app"
exe="$app/Contents/MacOS/AulaSync"

ok() { echo "ok: $*"; }
fail() { echo "::error::$*"; exit 1; }

hdiutil attach -nobrowse -readonly -mountpoint "$mnt" "$dmg" > /dev/null
ditto "$mnt/AulaSync.app" "$app"
# "Resource busy" sker af og til; ellers frakobler trap'en diskbilledet til sidst.
hdiutil detach "$mnt" > /dev/null 2>&1 || { sleep 5; hdiutil detach -force "$mnt" > /dev/null 2>&1 || true; }
echo "== $(basename "$dmg") på $(uname -m)"

team=$(codesign -dv "$app" 2>&1 | sed -n 's/^TeamIdentifier=//p')
if [ -n "$team" ] && [ "$team" != "not set" ]; then
  ok "signeret med Developer ID"
elif [ "$hardened" = --hardened ]; then
  ent="$work/adhoc.entitlements"
  cp "$here/AulaSync.entitlements" "$ent"
  /usr/libexec/PlistBuddy -c "Add :com.apple.security.cs.disable-library-validation bool true" "$ent"
  while IFS= read -r -d '' f; do
    codesign --force --options runtime --sign - "$f"
  done < <(find "$app/Contents/MacOS" -type f ! -path "$exe" -print0)
  codesign --force --options runtime --entitlements "$ent" --sign - "$app"
  ok "signeret ad hoc med hardened runtime og AulaSync.entitlements (+ disable-library-validation)"
fi
codesign --verify --deep --strict "$app" || fail "signaturen holder ikke"
if [ "$hardened" = --hardened ] || [ "$team" != "not set" ]; then
  # Udskriften gemmes først: grep -q stopper ved første fund, og med pipefail kunne codesign så fejle på en lukket pipe.
  info=$(codesign -dv "$app" 2>&1)
  grep -q 'flags=.*runtime' <<< "$info" || fail "app'en har ikke hardened runtime"
fi
xattr -dr com.apple.quarantine "$app" 2> /dev/null || true

# .NET starter: uden en kørende AulaSync returnerer --quit straks 0.
rc=0; "$exe" --quit || rc=$?
[ "$rc" = 0 ] || fail "AulaSync --quit gav $rc (starter .NET ikke?)"
ok ".NET starter"

# Menulinjen: AulaSync skal stadig køre efter 15 sekunder og lukke pænt med --quit (2 = den kørte og er afsluttet).
"$exe" --silent > "$work/run.log" 2>&1 &
pid=$!
for _ in $(seq 15); do
  sleep 1
  kill -0 "$pid" 2> /dev/null || { cat "$work/run.log"; fail "AulaSync --silent stoppede af sig selv"; }
done
rc=0; "$exe" --quit || rc=$?
[ "$rc" = 2 ] || { kill "$pid" 2> /dev/null || true; fail "AulaSync --quit gav $rc, ikke 2"; }
for _ in $(seq 15); do kill -0 "$pid" 2> /dev/null || break; sleep 1; done
! kill -0 "$pid" 2> /dev/null || fail "AulaSync lukkede ikke efter --quit"
log="$HOME/Library/Application Support/AulaSync/aulasync.log"
if [ -f "$log" ] && grep -q 'Uventet fejl' "$log"; then cat "$log"; fail "loggen har en uventet fejl"; fi
ok "menulinjen starter og lukker"

# Afinstallér AulaSync…: data og start ved login væk, app'en i papirkurven.
rc=0; "$exe" --uninstall || rc=$?
[ "$rc" = 0 ] || fail "AulaSync --uninstall gav $rc"
[ ! -e "$app" ] || fail "AulaSync.app er ikke flyttet til papirkurven"
[ ! -e "$HOME/Library/Application Support/AulaSync" ] || fail "datamappen er der stadig"
[ ! -e "$HOME/Library/LaunchAgents/dk.rpaasch.aulasync.plist" ] || fail "start ved login er der stadig"
rm -rf "$HOME/.Trash/AulaSync"*.app
ok "Afinstallér AulaSync… rydder op"
