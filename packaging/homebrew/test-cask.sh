#!/usr/bin/env bash
# Prøver Homebrew-cask'en af på en Mac (release.yml) i en midlertidig tap (aulasync/ci), så den rigtige tap ikke røres:
# brew style og audit (Apple Silicon og Intel), installation fra den lokale dmg (udgivelsen findes ikke endnu, så kun
# url'en skiftes), afinstallation uden --zap (data og start ved login bliver) og med --zap (alt er væk), og til sidst
# Afinstallér AulaSync… i appen efter en installation med Homebrew.
#
#   test-cask.sh out/homebrew/aulasync.rb out/dmg
set -euo pipefail

cask=$1
dmgdir=$(cd "$2" && pwd)
export HOMEBREW_NO_AUTO_UPDATE=1 HOMEBREW_NO_INSTALL_CLEANUP=1 HOMEBREW_NO_ENV_HINTS=1 HOMEBREW_NO_ANALYTICS=1
name=aulasync/ci/aulasync
app=/Applications/AulaSync.app
lib="$HOME/Library"
data=(
  "$lib/Application Support/AulaSync"
  "$lib/Caches/dk.rpaasch.aulasync"
  "$lib/HTTPStorages/dk.rpaasch.aulasync"
  "$lib/LaunchAgents/dk.rpaasch.aulasync.plist"
  "$lib/Preferences/dk.rpaasch.aulasync.plist"
  "$lib/Saved Application State/dk.rpaasch.aulasync.savedState"
  "$lib/WebKit/dk.rpaasch.aulasync"
)

ok() { echo "ok: $*"; }
fail() { echo "::error::$*"; exit 1; }

brew --version | head -1
brew untap aulasync/ci > /dev/null 2>&1 || true
brew tap-new --no-git aulasync/ci > /dev/null
tap=$(brew --repository aulasync/ci)
mkdir -p "$tap/Casks"
cp "$cask" "$tap/Casks/aulasync.rb"

echo "== brew style og audit"
brew style --cask "$name"
for arch in arm intel; do brew audit --cask --strict --arch="$arch" "$name"; done
ok "brew style og brew audit"

version=$(sed -n 's/^  version "\(.*\)"$/\1/p' "$cask")
[ -n "$version" ] || fail "fandt ingen version i $cask"
sed "s#^  url \"https://github.com/[^\"]*/AulaSync-#  url \"file://$dmgdir/AulaSync-#" "$cask" > "$tap/Casks/aulasync.rb"
grep -q "url \"file://$dmgdir/" "$tap/Casks/aulasync.rb" || fail "url'en blev ikke skiftet"

echo "== Installation"
[ ! -e "$app" ] || fail "$app findes allerede"
brew install --cask "$name"
[ -d "$app" ] || fail "$app blev ikke installeret"
got=$(/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' "$app/Contents/Info.plist")
[ "$got" = "$version" ] || fail "AulaSync.app har version $got, ikke $version"
ok "AulaSync $version ligger i Programmer"
# Signaturerne (også dem i udvidede attributter på .dll og .json) skal have overlevet Homebrews kopi ud af dmg'en.
codesign --verify --deep --strict "$app" || fail "signaturen på $app holder ikke efter brew install"
ok "signaturen holder efter brew install"
# Developer ID og notariseret (MAC_EXPECT_NOTARIZED=true): Gatekeeper godkender appen, som Homebrew har lagt den, uden
# "Åbn alligevel", og billetten er hæftet på (virker også uden net). Kun ét -v, så navnet i certifikatet (origin=) ikke
# står i loggen.
if [ "${MAC_EXPECT_NOTARIZED:-false}" = true ]; then
  if grep -q 'assessments disabled' <<< "$(spctl --status 2>&1 || true)"; then
    echo "::warning::Gatekeeper er slået fra på denne maskine; kun hæftningen tjekkes"
  else
    gk=$(spctl -a -t exec -v "$app" 2>&1) || fail "Gatekeeper afviser $app: $(grep -v '^origin=' <<< "$gk")"
    grep -q '^source=Notarized Developer ID' <<< "$gk" || fail "$app er ikke notariseret: $(grep -v '^origin=' <<< "$gk")"
  fi
  xcrun stapler validate "$app" > /dev/null || fail "$app har ingen hæftet notar-billet"
  ok "Developer ID, notariseret og hæftet"
fi

# Det, AulaSync selv gemmer, lægges ind, som om den havde kørt.
seed() {
  for path in "${data[@]}"; do
    mkdir -p "$(dirname "$path")"
    case $path in *.plist) printf '<?xml version="1.0"?>\n<plist version="1.0"><dict/></plist>\n' > "$path" ;; *) mkdir -p "$path" && touch "$path/fil" ;; esac
  done
}
seed

echo "== Afinstallation uden --zap"
brew uninstall --cask "$name"
[ ! -e "$app" ] || fail "$app er der stadig"
for path in "${data[@]}"; do [ -e "$path" ] || fail "$path blev slettet uden --zap"; done
ok "programmet er væk, data og start ved login er der stadig"

echo "== Afinstallation med --zap"
brew install --cask "$name" > /dev/null
brew uninstall --zap --cask "$name"
[ ! -e "$app" ] || fail "$app er der stadig"
for path in "${data[@]}"; do [ ! -e "$path" ] || fail "$path er der stadig efter --zap"; done
ok "--zap fjerner alt, AulaSync har gemt"

# Afinstallér AulaSync… i appen (her uden vindue: AulaSync --uninstall) efter en installation med Homebrew: appen
# kommer i papirkurven, data og start ved login er væk, og Homebrew har glemt AulaSync, så brew upgrade ikke lægger
# den tilbage. Indstillingerne (Preferences) fjerner appen med defaults delete, ikke som fil, så de prøves ikke her.
echo "== Afinstallér AulaSync… efter installation med Homebrew"
brew install --cask "$name" > /dev/null
seed
caskroom="$(brew --prefix)/Caskroom/aulasync"
[ -d "$caskroom" ] || fail "Homebrews optegnelse $caskroom mangler"
# Som når brugeren har åbnet AulaSync første gang; ellers kan macOS stoppe en prøvebygning, der kun er ad hoc-signeret.
xattr -dr com.apple.quarantine "$app" 2> /dev/null || true
"$app/Contents/MacOS/AulaSync" --uninstall || fail "AulaSync --uninstall fejlede ($?)"
[ ! -e "$app" ] || fail "$app er der stadig"
[ -d "$HOME/.Trash/AulaSync.app" ] || fail "AulaSync.app er ikke i papirkurven"
for path in "${data[@]}"; do
  case $path in */Preferences/*) ;; *) [ ! -e "$path" ] || fail "$path er der stadig efter Afinstallér AulaSync…" ;; esac
done
[ ! -e "$caskroom" ] || fail "Homebrews optegnelse $caskroom er der stadig"
if brew list --cask aulasync > /dev/null 2>&1; then fail "brew list viser stadig aulasync"; fi
rm -rf "$HOME/.Trash/AulaSync.app" "$lib/Preferences/dk.rpaasch.aulasync.plist"
ok "Afinstallér AulaSync… fjerner også Homebrews optegnelse"

brew untap aulasync/ci > /dev/null
