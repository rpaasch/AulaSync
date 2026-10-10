# Homebrew-tap til AulaSync

[AulaSync](https://github.com/rpaasch/AulaSync) holder skemaer fra Aula opdateret i din kalender. Her ligger Mac-udgaven
som Homebrew-cask. Filen `Casks/aulasync.rb` opdateres ved hver udgivelse af AulaSync; ret den ikke i hånden.

## Installér

```sh
brew install --cask rpaasch/tap/aulasync
```

Kommandoen tilføjer tap'en og stoler kun på AulaSync herfra. AulaSync lægges i Programmer.

AulaSync er signeret af udvikleren og undersøgt af Apple (notariseret). Første gang du åbner AulaSync fra **Programmer**,
spørger macOS, om du vil åbne en app, der er hentet fra internettet. Klik **Åbn**.

## Opdatér

```sh
brew upgrade --cask rpaasch/tap/aulasync
```

Med det fulde navn henter Homebrew tap'en først, så en ny udgave ses med det samme (ellers kan der gå op til et døgn).
Kører AulaSync, lukker Homebrew den først og åbner den igen bagefter. macOS spørger ikke igen, fordi den nye udgave er
signeret af den samme udvikler.

Siger macOS, at "Terminal" blev forhindret i at ændre apps på din Mac, er opdateringen alligevel gået igennem: Homebrew
fjerner så den gamle AulaSync og lægger den nye ind. Vil du undgå beskeden, så slå Terminal til under
**Systemindstillinger › Anonymitet & sikkerhed › Appadministration**, og luk og åbn Terminal igen.

## Afinstallér

```sh
brew uninstall --zap --cask aulasync
```

Det fjerner AulaSync og flytter alt, den har gemt på Macen, til papirkurven: indstillinger, skemaer, log, Aula-login og
start ved login. Tøm papirkurven bagefter, så dit Aula-login er helt væk. Uden `--zap` fjernes kun programmet.
**Afinstallér AulaSync…** i AulaSyncs Indstillinger fjerner det samme og får Homebrew til at glemme AulaSync, men sletter
dine data med det samme; kun selve AulaSync lægges i papirkurven. Slet bagefter AulaSync-kalenderne i dit
kalenderprogram; dem kan AulaSync ikke slette.
