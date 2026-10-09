# AulaSync

Skemaer fra [Aula](https://www.aula.dk) som kalendere, der holder sig opdateret, på Windows og Mac.

Vælg medarbejdere, klasser og lokaler. Hvert skema bliver sin egen kalender i dit kalenderprogram, og AulaSync henter ændringer fra Aula hver 4. time (det kan du ændre i Indstillinger), så længe programmet kører. AulaSync taler kun med Aula og gemmer alt på din computer.

> **Kommer du fra 2.x?** Beskeder synkroniseres ikke længere, og indstillinger fra 2.x overføres ikke. Se [Opgradering fra 2.x](docs/vejledning.md#opgradering-fra-2x).

## Kalenderprogrammer

| Kalenderprogram | Knappen i AulaSync | Opdateres |
|---|---|---|
| Apple Kalender (Mac) | **Tilføj til Kalender** | automatisk |
| Outlook (klassisk) på Windows | **Tilføj til Outlook** | automatisk |
| Ny Outlook, Outlook til Mac og Outlook på nettet | **Importér…** | ikke automatisk (øjebliksbillede); AulaSync siger til, når skemaet er ændret |
| Andet program på samme computer, fx Thunderbird | **Kopiér adresse** | automatisk |

**AulaSync og kalenderprogrammet skal køre på samme computer.** Kalenderprogrammet henter skemaerne fra AulaSync på `http://localhost:9876` (eller en af de næste porte, hvis 9876 var optaget ved første start), altså din egen computer. En kalender på nettet eller på telefonen kan ikke nå den. Det gælder også ny Outlook, Outlook til Mac og Outlook på nettet, som henter kalendere gennem Microsofts servere. Til dem importerer du i stedet skemaet som en fil med **Importér…**. I Apple Kalender skal abonnementet ligge **På min Mac**, ikke i iCloud.

## Installation

### Windows 10 og 11 (64-bit)

Hent [AulaSync-Setup.exe](https://github.com/rpaasch/AulaSync/releases/latest/download/AulaSync-Setup.exe), og åbn den. Klik **Installer**, og lad **Start AulaSync** være markeret, når du klikker **Færdig**. AulaSync installeres kun for dig, i `%LOCALAPPDATA%\Programs\AulaSync`, og kræver ingen administratorrettigheder. Bagefter ligger AulaSync i Start-menuen.

Programmet er ikke signeret: viser Windows "Windows beskyttede din pc", så vælg **Flere oplysninger** › **Kør alligevel**. Gør det kun med en `AulaSync-Setup.exe`, du selv har hentet fra udgivelsessiden. Mangler **Kør alligevel**, eller siger Windows, at programmet er blokeret, har computerens administrator lukket for programmer, der ikke er godkendt. Så spørg skolens IT.

Med winget: højreklik på Start-knappen, vælg **Terminal** (på Windows 10: **Windows PowerShell**), og skriv:

```
winget install rpaasch.AulaSync
```

AulaSync installeres, lægges i Start-menuen og åbner. En ny udgave kommer først i winget, når Microsoft har godkendt den.

**Ny udgave:** Hent og åbn AulaSync-Setup.exe igen. Kører AulaSync, lukkes den og startes igen bagefter. Med winget: højreklik på AulaSync-ikonet i systembakken ved uret, vælg **Afslut AulaSync**, skriv `winget upgrade rpaasch.AulaSync` i Terminal, og start AulaSync fra Start-menuen bagefter.

**Har du en ældre udgave:** Har du lagt `AulaSync.exe` i en mappe selv, så installér AulaSync-Setup.exe. Start ved login flytter med til den installerede AulaSync, og bagefter kan du slette den gamle `AulaSync.exe`. Har du AulaSync fra winget fra før 3.2, så opdatér med winget som beskrevet ovenfor. winget kan ikke skifte til installationsprogrammet, så du beholder udgaven uden installation; den ligger også i Start-menuen.

Login bruger Microsoft Edge WebView2 Runtime, som følger med Windows 11 og findes på de fleste computere med Windows 10. Mangler den, siger AulaSync til.

### Mac (macOS 15 eller nyere)

Hent `AulaSync-<version>-arm64.dmg` (Apple-chip, M1 og nyere) eller `AulaSync-<version>-x64.dmg` (Intel) fra [den nyeste udgave](https://github.com/rpaasch/AulaSync/releases/latest), åbn den, og træk AulaSync over i **Programmer**.

Appen er endnu ikke signeret af Apple, så første gang skal du give lov. Gør det kun, hvis du selv har hentet `.dmg`-filen fra udgivelsessiden:

1. Åbn AulaSync fra **Programmer** (ikke fra `.dmg`-filen, ellers kan AulaSync ikke starte af sig selv, når du logger ind). macOS siger, at appen ikke blev åbnet. Klik **Udført**, ikke **Flyt til papirkurv**.
2. Åbn **Systemindstillinger › Anonymitet og sikkerhed**, rul ned, og klik **Åbn alligevel** ved AulaSync. Knappen står der kun i cirka en time. Er den væk, så åbn AulaSync fra Programmer igen.
3. Når macOS spørger igen, så klik **Åbn alligevel**, og bekræft med adgangskoden til din Mac eller med Touch ID.

## Kom i gang

Første gang guider AulaSync dig gennem fem korte trin. Efter velkomsten logger du ind på Aula (som du plejer, fx med MitID eller UniLogin), vælger dit kalenderprogram, vælger skemaer og tilføjer dem til din kalender. Bagefter kører AulaSync videre i menulinjen (Mac) eller systembakken (Windows) og starter selv, når du logger ind på computeren. Trinene kan vises igen med **Kom i gang igen…** under Hjælp i Indstillinger, hvor **Vejledning** også åbner vejledningen.

Hele vejledningen står i [docs/vejledning.md](docs/vejledning.md).

## Data og privatliv

AulaSync gemmer alt på din egen computer. Din adgangskode gemmer AulaSync ikke. Du logger ind på Aulas egen login-side, og AulaSync gemmer cookies fra Aula og fra login-tjenesten (fx UniLogin eller din kommunes login) i en browserprofil. Så forbliver du logget ind, og AulaSync kan selv logge ind igen i baggrunden, så længe login-tjenesten husker dig. Filerne ligger her:

| | |
|---|---|
| Windows | `%LOCALAPPDATA%\AulaSync` |
| Mac | `~/Library/Application Support/AulaSync` |

| Fil eller mappe | Indhold |
|---|---|
| `kalendere/` | Én `.ics`-fil pr. skema, fx `123456-AE-medarbejder-1001.ics`: institutionens nummer, medarbejderens initialer (eller navn, hvis Aula ikke har initialer) eller klassens eller lokalets navn, typen og skemaets Aula-id |
| `abonnementer.json` | De skemaer, du har valgt |
| `config.json` | Kalenderprogram, første start, kalender-serverens port, hvor tit skemaerne hentes, og om opdateringstiden vises i kalenderen |
| `aulasync.log` | Log. Indeholder bl.a. navne og initialer på de skemaer, du har valgt, deres Aula-id'er og din institutions nummer. Højst ca. 1 MB; ældre linjer flyttes til `aulasync.log.old` |
| `webview/` (Windows) | Login til Aula. Slettes, når du logger ud. Er mappen i brug, slettes cookies i stedet, næste gang AulaSync viser Aulas login |
| `webview-id.txt` (Mac) | Hvilket login-lager AulaSync bruger. Selve login gemmer macOS i AulaSyncs WebKit-data uden for mappen. Når du logger ud, skifter AulaSync til et nyt, tomt lager; det gamle bruges ikke igen |

Kalender-serveren tager kun imod forbindelser fra din egen computer (`127.0.0.1` og `::1`) og udleverer kun skemafilerne. En hjemmeside i din browser kan ikke læse dem. Serveren har ingen adgangskode, så andre programmer og andre brugere, der er logget ind på samme computer, kan hente skemaerne. Porten vælges ved første start: 9876 eller den første ledige op til 9899, så flere brugere på samme computer som regel kan køre AulaSync samtidig. Porten reserveres dog ikke: kørte en anden brugers AulaSync ikke, da du startede AulaSync første gang, kan I få samme port, og så kan kun én af jer ad gangen køre kalender-serveren.

**Log ud…** i Indstillinger glemmer dit login og sletter dine valgte skemaer, og første start vises igen. Kalenderfilerne bliver liggende, så kalenderne viser stadig de seneste skemaer, men de bliver ikke opdateret, før du vælger skemaerne igen. **Afinstallér AulaSync…** i Indstillinger fjerner alt, AulaSync har gemt på computeren: programmet, start ved login, dit login, dine indstillinger, loggen og kalenderfilerne. På Windows også genvejen i Start-menuen og mappen fra AulaSync 2 (`%USERPROFILE%\.aulasync`); på Mac også AulaSyncs WebKit-data med login og cache, og AulaSync flyttes til papirkurven. Det samme sker, når du afinstallerer AulaSync på Windows fra **Indstillinger › Apps** (**Installerede apps**, på Windows 10 **Apps og funktioner**) eller med `winget uninstall rpaasch.AulaSync`. Har du AulaSync fra winget fra før 3.2, så brug **Afinstallér AulaSync…**; **Indstillinger › Apps** og `winget uninstall` sletter dér kun programmet. Knappen findes fra AulaSync 3.2. Slet bagefter AulaSync-kalenderne i dit kalenderprogram; dem kan AulaSync ikke slette.

Importerer du et skema i ny Outlook, Outlook til Mac eller Outlook på nettet, ligger det bagefter også i din Outlook-konto hos Microsoft.

## Byg selv

Kræver [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0).

```
dotnet run --project src/AulaSync.App
dotnet test AulaSync.slnx
```

En udgivelse laves af GitHub Actions (`.github/workflows/release.yml`), når et versionsmærke som `v3.2.0` pushes: installationsprogrammet til Windows (`AulaSync-Setup.exe`, Inno Setup), den løse `AulaSync.exe` (til den portable udgave i winget og til IT), Mac-dmg'er og winget-manifester (se `packaging/`). Ikonet tegnes af `packaging/icon/make-icons.py`.

## Ansvarsfraskrivelse

AulaSync er et uafhængigt open source-projekt og er **ikke tilknyttet, godkendt af eller supporteret af Aula, KOMBIT, Netcompany eller KMD**. Appen bruger Aulas interne API, som ikke er offentligt dokumenteret og kan ændre sig uden varsel, så AulaSync kan holde op med at virke fra den ene dag til den anden.

Kalenderne indeholder personoplysninger om dine kolleger: navne, initialer og hvem der er vikar for hvem. Følg din skoles og kommunes regler for, hvor de må ligge, også når du importerer dem i ny Outlook, Outlook til Mac eller Outlook på nettet, som gemmer dem hos Microsoft. Brug på eget ansvar.

## Fejl og forslag

Opret et [issue på GitHub](https://github.com/rpaasch/AulaSync/issues), hvis du finder en fejl eller har et forslag. Vedhæft gerne de sidste linjer fra `aulasync.log` (Indstillinger › Fejlfinding › **Åbn log**). Alle kan læse et issue, så erstat først navne, initialer, din institutions nummer, Aula-id'erne og dit brugernavn i filstier med opdigtede værdier, fx `Anna Eksempel`. Filnavnene rummer også institutionens nummer, initialer eller navne og Aula-id'er, fx `123456-AE-medarbejder-1001.ics`; skriv dem fx som `1-XX-medarbejder-1.ics` (både nummer, initialer eller navn og id).

## Licens

MIT. Se [LICENSE](LICENSE).
