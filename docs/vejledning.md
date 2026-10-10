# Vejledning til AulaSync

AulaSync henter skemaer fra Aula og holder dem opdateret som kalendere i dit kalenderprogram. Du vælger selv, hvilke skemaer du vil have: dit eget, kollegers, klassers og lokalers.

## Før du starter

- **AulaSync og kalenderprogrammet skal køre på samme computer.** Kalenderprogrammet henter skemaerne fra AulaSync på din egen computer (`http://localhost:9876`, eller en af de næste porte, hvis 9876 var optaget ved første start). Derfor ser du skemaerne på denne computer, ikke på telefonen eller i en kalender på nettet. Undtagelsen er import i ny Outlook, Outlook til Mac og Outlook på nettet: Så ligger skemaet i din Outlook-konto hos Microsoft og kan også ses på telefonen, men det opdateres ikke af sig selv (se [Kalenderprogrammer](#kalenderprogrammer)).
- **Kalenderne opdateres kun, mens AulaSync kører.** AulaSync starter selv, når du logger ind på computeren, og kører i menulinjen (Mac) eller systembakken (Windows).
- Du skal kunne logge ind på Aula som medarbejder.
- Windows 10 eller 11 (64-bit), eller macOS 15 eller nyere.

## Installation

### Windows

1. Hent [AulaSync-Setup.exe](https://github.com/rpaasch/AulaSync/releases/latest/download/AulaSync-Setup.exe). Advarer browseren om, at filen ikke hentes ofte, så vælg at beholde den.
2. Åbn filen. Viser Windows "Windows beskyttede din pc", så klik **Flere oplysninger** › **Kør alligevel**. Det sker, fordi programmet ikke er signeret, så Windows ikke kan se, hvem der har udgivet det. Gør det kun med en `AulaSync-Setup.exe`, du selv har hentet fra udgivelsessiden.
3. Klik **Installer**, og klik **Færdig**, når installationen er færdig. AulaSync åbner af sig selv.

AulaSync installeres kun for dig og kræver ingen administratorrettigheder. Bagefter ligger AulaSync i Start-menuen, så du kan finde den ved at skrive *AulaSync* i Start. Mangler **Kør alligevel**, eller siger Windows, at programmet er blokeret, har skolens IT lukket for programmer, de ikke har godkendt. Så spørg IT.

**Med winget:** Højreklik på Start-knappen, vælg **Terminal** (på Windows 10: **Windows PowerShell**), og skriv `winget install rpaasch.AulaSync`. Spørger winget, om du accepterer vilkårene, så svar ja. AulaSync installeres, lægges i Start-menuen og åbner. En ny udgave kommer først i winget, når Microsoft har godkendt den.

**Ny udgave:** Hent og åbn AulaSync-Setup.exe igen. Kører AulaSync, lukkes den og startes igen bagefter. Med winget: højreklik på AulaSync-ikonet i systembakken ved uret, vælg **Afslut AulaSync**, skriv `winget upgrade rpaasch.AulaSync` i Terminal, og start AulaSync fra Start-menuen bagefter.

**Har du en ældre udgave uden installation?** Har du selv lagt `AulaSync.exe` i en mappe, fx **Dokumenter**, så installér AulaSync-Setup.exe som beskrevet ovenfor. Kører den gamle AulaSync, lukkes den. Start ved login og genvejen i Start-menuen flytter med til den installerede AulaSync. Bagefter kan du slette den gamle `AulaSync.exe`. Har du AulaSync fra winget fra før 3.2, så opdatér med winget som beskrevet ovenfor. winget kan ikke skifte til installationsprogrammet, så du beholder udgaven uden installation. Den ligger også i Start-menuen, når den har kørt én gang.

**Afinstallér:** Vælg **Afinstallér AulaSync…** i AulaSyncs [Indstillinger](#indstillinger), eller vælg **Indstillinger › Apps** og så **Installerede apps** (på Windows 10: **Apps og funktioner**), find AulaSync, og vælg **Afinstaller** (på Windows 11 under **⋯** ved AulaSync). Når du har bekræftet, slettes alt, AulaSync har gemt på computeren: programmet, genvejen i Start-menuen, start ved login, dit login, dine indstillinger, loggen, kalenderfilerne og mappen fra AulaSync 2. Det gælder også `winget uninstall rpaasch.AulaSync`, som ikke spørger. Slet bagefter AulaSync-kalenderne i dit kalenderprogram; dem kan AulaSync ikke slette. Har du AulaSync uden installation eller fra winget fra før 3.2, så brug **Afinstallér AulaSync…** i AulaSyncs Indstillinger; det sletter også `AulaSync.exe`. **Indstillinger › Apps** og `winget uninstall` sletter dér kun programmet. Knappen findes fra AulaSync 3.2; har du en ældre udgave, så hent den nyeste først (se **Ny udgave** ovenfor).

Login bruger Microsoft Edge WebView2 Runtime, som følger med Windows 11 og findes på de fleste computere med Windows 10. Mangler den, står der i login-vinduet, hvor du henter den.

### Mac

1. Hent den rigtige `.dmg`-fil fra [den nyeste udgave](https://github.com/rpaasch/AulaSync/releases/latest). Filerne står under **Assets**: `AulaSync-<version>-arm64.dmg` er til Mac med Apple-chip (M1 og nyere), og `AulaSync-<version>-x64.dmg` er til Mac med Intel-processor. Er du i tvivl, så klik på æblet øverst til venstre på skærmen, og vælg **Om denne Mac**. Står der **Chip** og **Apple M…**, skal du bruge `arm64`. Står der **Processor** og **Intel**, skal du bruge `x64`.
2. Åbn filen, og træk AulaSync over i **Programmer**.
3. Åbn AulaSync fra **Programmer**, ikke fra `.dmg`-filen. Ellers kan AulaSync ikke starte af sig selv, når du logger ind. Første gang spørger macOS, om du vil åbne AulaSync, fordi den er hentet fra internettet. Under spørgsmålet står, at Apple har undersøgt den og ikke fundet noget ondsindet software (AulaSync er signeret af udvikleren og notariseret af Apple). Klik **Åbn**.

Beder macOS om en administrators navn og adgangskode, når du trækker AulaSync over i **Programmer**, eller lader macOS dig ikke åbne AulaSync, fx på en skole-Mac, som IT styrer, så bed IT om hjælp.

Kommer der en ny udgave, så vælg **Afslut AulaSync** i ikonets menu. Hent den nye `.dmg`-fil, og træk AulaSync over i **Programmer**, så den erstatter den gamle. Åbn den fra **Programmer**, og klik **Åbn**, hvis macOS spørger.

**Med Homebrew:** Har du [Homebrew](https://brew.sh), kan du i stedet skrive `brew install --cask rpaasch/tap/aulasync` i Terminal. Homebrew henter selv den rigtige udgave til din Mac og lægger AulaSync i **Programmer**. Første gang du åbner den, spørger macOS som i trin 3. En ny udgave får du med `brew upgrade --cask rpaasch/tap/aulasync` (med det fulde navn ser Homebrew den nye udgave med det samme). Kører AulaSync, lukker Homebrew den og åbner den igen. macOS spørger ikke igen.

**Afinstallér:** Vælg **Afinstallér AulaSync…** i AulaSyncs [Indstillinger](#indstillinger). Start ved login, dit login, dine indstillinger, loggen og kalenderfilerne slettes, og AulaSync flyttes til papirkurven. Slet bagefter AulaSync-kalenderne i dit kalenderprogram; dem kan AulaSync ikke slette. Knappen findes fra AulaSync 3.2. Ligger AulaSync stadig i **Programmer** bagefter, fx fordi du ikke er administrator på Macen, så træk den selv til papirkurven. Trækker du bare AulaSync til papirkurven uden at afinstallere, bliver start ved login og dit login liggende. Har du AulaSync fra Homebrew, fjerner `brew uninstall --zap --cask aulasync` det samme, men flytter dit login og dine data til papirkurven i stedet for at slette dem, så tøm papirkurven bagefter; uden `--zap` fjernes kun programmet. Knappen får også Homebrew til at glemme AulaSync.

## Første start

AulaSync guider dig gennem fem trin:

1. **Velkommen.** Klik **Log ind med Aula**.
2. **Log ind.** Aulas login vises i vinduet. Log ind som medarbejder, som du plejer (MitID, UniLogin eller din kommunes login). Når du er logget ind, går AulaSync selv videre. Skemaerne hentes fra den institution, der står på din medarbejderprofil i Aula. Du kan se den i Indstillinger.
3. **Kalenderprogram.** Vælg det program, du bruger (se [Kalenderprogrammer](#kalenderprogrammer)), og klik **Fortsæt**. Bruger du Outlook på Windows, så se øverst i Outlook: står der **Ny Outlook** med en kontakt, der er slået til, bruger du ny Outlook. Ellers bruger du klassisk Outlook. Du kan altid ændre valget i Indstillinger.
4. **Vælg skemaer.** Dit eget skema er valgt på forhånd. Søg efter navn, initialer, klasse eller lokale, og klik **Tilføj**. Du kan tilføje flere senere. Klik **Fortsæt**, når du har valgt.
5. **Tilføj til din kalender.** Klik knappen ved hvert skema (se nedenfor), og klik **Færdig**.

Når du klikker **Færdig**, lukker vinduet, og AulaSync kører videre som et lille ikon: på Mac i menulinjen øverst til højre, på Windows i systembakken ved uret nederst til højre. Windows gemmer ofte nye ikoner under pilen **^**. Træk AulaSync-ikonet derfra ned på proceslinjen, så du altid kan se det. AulaSync starter nu selv, når du logger ind på computeren. På Mac viser macOS måske en besked om, at et program kan køre i baggrunden. Det er AulaSync. Slå det ikke fra, ellers starter AulaSync ikke selv.

Vil du igennem trinene igen, så vælg **Kom i gang igen…** under **Hjælp** i [Indstillinger](#indstillinger). Er du logget ind, går **Fortsæt** direkte til valget af kalenderprogram. De skemaer, du har valgt, er stadig valgt, og **Start AulaSync, når jeg logger ind** bliver, som du har sat det.

## Kalenderprogrammer

| Kalenderprogram | Knappen | Opdateres |
|---|---|---|
| Apple Kalender (Mac) | **Tilføj til Kalender** | automatisk, så længe AulaSync kører |
| Outlook (klassisk) på Windows | **Tilføj til Outlook** | automatisk, så længe AulaSync kører |
| Ny Outlook, Outlook til Mac og Outlook på nettet | **Importér…** | ikke automatisk (øjebliksbillede); AulaSync siger til, når skemaet er ændret, så du kan importere igen |
| Andet program, fx Thunderbird | **Kopiér adresse** | automatisk, så længe AulaSync kører |

### Apple Kalender

1. Klik **Tilføj til Kalender**. Kalender åbner og spørger, om du vil abonnere. Klik **Abonner**.
2. Vælg **Placering: På min Mac** (ikke iCloud). Med iCloud er det Apples servere, der henter kalenderen, og de kan ikke nå AulaSync på denne Mac.
3. Vælg gerne **Opdater automatisk: Hver time**, og klik **OK**.

### Outlook (klassisk)

1. Klik **Tilføj til Outlook**. Outlook spørger, om du vil tilføje kalenderen og abonnere på opdateringer. Klik **Ja**.
2. Kalenderen står under **Andre kalendere**.

Samtidig viser AulaSync altid vinduet "Åbnede Outlook ikke kalenderen?". Står kalenderen i Outlook, så klik bare **OK**. Åbnede ny Outlook i stedet, eller skete der ingenting, så klik **Kopiér adresse** i vinduet. Gå så til kalenderen i klassisk Outlook, vælg **Tilføj kalender › Fra internettet** under fanen **Hjem**, sæt adressen ind, og bekræft. Klik til sidst **OK** i AulaSync.

### Ny Outlook, Outlook til Mac og Outlook på nettet

Disse udgaver henter kalendere gennem Microsofts servere, som ikke kan nå din computer. Derfor importerer du skemaet som en fil i Outlook på nettet eller i ny Outlook. Bruger du Outlook til Mac, så importér i Outlook på nettet. Kalenderen kommer bagefter også frem i Outlook til Mac.

Klik **Importér…** ved skemaet. AulaSync viser tre trin, hver med sin egen knap:

1. **Opret en tom kalender** i Outlook med skemaets navn, fx med **Tilføj kalender › Opret tom kalender**. **Kopiér navn** kopierer navnet, så du kan sætte det ind.
2. **Åbn Outlook-kalenderen**, vælg **Tilføj kalender › Upload fra fil**, og vælg den nye kalender. **Åbn Outlook på nettet** åbner kalenderen i din browser.
3. **Vælg filen.** **Vis fil i Finder** (Mac) eller **Vis fil i Stifinder** (Windows) viser, hvor den ligger. Filen hedder fx `123456-7A-klasse-4711.ics`: institutionens nummer, skemaets navn (for en medarbejder initialerne, ellers navnet), typen og skemaets id i Aula. Afslut importen i Outlook, og klik **Færdig** i AulaSync.

En import er et øjebliksbillede. Når skemaet ændrer sig i Aula, siger AulaSync til: der kommer en besked ("7A er ændret siden import"), ikonet i menulinjen eller systembakken får en prik, og rækken viser **Ændret siden import**. Slet så kalenderen i Outlook, og importér igen med **Importér igen…**. En import dækker de næste ca. 90 dage. Når de er gået, siger AulaSync også til, så du kan importere igen.

### Andet program

Klik **Kopiér adresse**, og indsæt adressen i dit kalenderprogram, hvor man abonnerer på en kalender fra en adresse. Programmet skal køre på samme computer som AulaSync.

## Hovedvinduet "Mine skemaer"

Hver række er ét skema med type og antal lektioner. Til højre står knappen eller status:

| Står der | Betyder |
|---|---|
| **Venter på Kalender…** eller **Venter på Outlook…** | Du har klikket, og AulaSync venter på, at kalenderprogrammet henter skemaet. |
| **✓ Tilføjet** | Kalenderprogrammet har hentet skemaet fra AulaSync. |
| **Kalender hentede ikke skemaet** eller **Outlook hentede ikke skemaet** | Programmet har ikke hentet skemaet inden for 2 minutter efter klikket. Knappen står igen ved siden af. Se [Fejlfinding](#fejlfinding). |
| **Ikke hentet af Outlook siden …** med knappen ved siden af | Outlook henter dine andre skemaer, men ikke dette, så kalenderen er nok slettet i Outlook. Se [Fejlfinding](#ikke-hentet-af-outlook-siden-). |
| **Ikke hentet siden …** uden knap, fx **Ikke hentet af Kalender siden …** | Kalenderprogrammet henter dine andre skemaer, men ikke dette, eller har ikke hentet skemaet i tre dage (Apple Kalender: otte dage), fx fordi programmet ikke har været åbent, eller fordi kalenderen er slettet. Noten forsvinder, når programmet henter skemaet igen. Se [Fejlfinding](#ikke-hentet-af-outlook-siden-). |
| **✓ Kopieret** | Adressen er kopieret. Når et program henter den, står der **✓ Tilføjet**. |
| **Importeret i dag** eller fx **Importeret 5. okt.** | Skemaet er importeret i ny Outlook, Outlook til Mac eller Outlook på nettet. Datoen viser, hvornår du sidst importerede det. |
| **Ændret siden import** | Skemaet er ændret i Aula siden importen, eller importen dækker ikke længere fremad. Importér igen. |
| **Kunne ikke hentes kl. …** | Teksten står under navnet i stedet for type og antal lektioner. AulaSync kunne ikke hente skemaet fra Aula, fx fordi der ikke var forbindelse, eller fordi du ikke har adgang til skemaet. Den gamle kalender bliver stående. Knappen og status til højre er væk, til skemaet er hentet igen. Efter "prøver igen" står, hvornår AulaSync prøver igen; klik **Opdatér nu** for at prøve med det samme. |

**⋯** ved hver række: **Tilføj igen** (eller **Importér igen…**), **Kopiér adresse**, **Vis fil i Finder** (Mac) eller **Vis fil i Stifinder** (Windows) og **Fjern skema…**. Brug kun **Tilføj igen**, når kalenderen ikke står i kalenderprogrammet. Ellers abonnerer programmet en gang til, og lektionerne står to gange. Fjerner du et skema, sletter AulaSync kalenderfilen, og kalenderen bliver ikke længere opdateret. Slet også kalenderen i dit kalenderprogram, ellers kan programmet melde en fejl, hver gang det prøver at opdatere den.

Med **+ Tilføj skema** vælger du flere skemaer. Et nyt skema kommer ikke af sig selv i dit kalenderprogram: klik knappen ved skemaet, fx **Tilføj til Kalender**, ligesom ved første start. **Opdatér nu** henter alle skemaer fra Aula med det samme; ellers sker det hver 4. time, eller så tit, du har valgt i [Indstillinger](#indstillinger). Linjen nederst viser, hvornår der sidst blev opdateret, og hvornår næste gang er.

Hvert skema dækker 90 dage tilbage og 90 dage frem. Sådan står en lektion i kalenderen:

- Medarbejderskema: "Dansk | Lokale 53 | 7A" (fag, lokale, klasse)
- Klasseskema: "Dansk | Lokale 53 | AE" (fag, lokale, lærerens initialer)
- Lokaleskema: "Dansk | 7A | AE" (fag, klasse, lærerens initialer)

Er der vikar, står "Vikar" først, fx "Vikar | Dansk | Lokale 53 | 7A". I noterne til lektionen står lærernes fulde navne, klasse og lokale og ved vikar også, hvem vikaren dækker for, fx "Vikar for: Anna Eksempel (AE)". Del derfor kun kalenderne med nogen, der også må se oplysningerne i Aula, og importér dem kun i en Outlook-konto, hvor de må ligge, fx din arbejdskonto.

## Menulinjen og systembakken

Ikonet viser, hvordan det går. En prik midt i ikonet betyder, at AulaSync henter skemaer. En prik i øverste højre hjørne betyder, at noget kræver handling: du er logget ud af Aula, kalender-serveren kunne ikke starte, eller et importeret skema er ændret. Øverst i menuen står så, hvad du skal gøre, fx **Log ind igen…** eller **Start kalender-server igen**. Kalender-serveren er den del af AulaSync, som kalenderprogrammet henter skemaerne fra.

Menuen: status · **Åbn AulaSync…** · **Opdatér nu** · **Indstillinger…** · **Afslut AulaSync**. På Mac åbner et klik på ikonet menuen. På Windows åbner et klik på ikonet hovedvinduet, og et højreklik åbner menuen. Kan du ikke se ikonet på Windows, så klik på pilen **^** ved uret.

Du kan også få hovedvinduet frem ved at starte AulaSync igen. Der starter ikke en AulaSync mere; den, der allerede kører, viser hovedvinduet.

At lukke hovedvinduet afslutter ikke AulaSync. Vælg **Afslut AulaSync** i menuen, hvis du vil stoppe den; så holder kalenderne op med at blive opdateret.

## Indstillinger

Åbn Indstillinger fra ikonets menu (**Indstillinger…**). På Mac kan du også trykke ⌘,.

- **Kalenderprogram:** de samme fire valg som ved første start. Skifter du program, ændres kun knapperne i hovedvinduet. Skemaerne kommer ikke selv over i det nye program: klik knappen ved hvert skema, eller vælg **Tilføj igen** i ⋯-menuen, hvis rækken allerede viser **✓ Tilføjet**. Kalenderne i det gamle program bliver liggende, til du sletter dem dér.
- **Hent skemaer fra Aula:** hvor tit AulaSync henter skemaerne: hver halve time, hver time, hver 2. time, hver 4. time (standard) eller hver 8. time. Hver opdatering henter alle dine valgte skemaer fra Aula, så jo sjældnere, jo færre forespørgsler til Aula, men ændringer i Aula kommer også senere i kalenderen. **Opdatér nu** henter altid med det samme. Næste opdatering regnes fra den seneste, også når den kom fra **Opdatér nu** eller fra et login. Vælger du et kortere interval, end der er gået siden sidste opdatering, henter AulaSync med det samme. Uanset intervallet spørger AulaSync Aula hvert 10. minut, om du stadig er logget ind; det holder forbindelsen i live.
- **Vis opdateringstid i kalenderen:** Slået fra som standard. Slår du den til, får hver af AulaSyncs kalendere en privat aftale mandag kl. 5.45-6.00, der i titel og beskrivelse viser, hvornår skemaet sidst blev hentet fra Aula, fx "Opd. 071026@14:32" for 7. oktober 2026 kl. 14:32. Lørdag og søndag står aftalen i den kommende uge. Den optager ikke tid i kalenderen. Kalenderen viser det skema, kalenderprogrammet sidst har hentet fra AulaSync. Har AulaSync hentet et nyere siden, viser aftalen et tidligere tidspunkt end **Opdateret** nederst i AulaSyncs vindue, indtil kalenderprogrammet henter igen. Outlook henter typisk inden for en time, mens Outlook er åben; i Apple Kalender kan du trykke ⌘R. Ændringen kommer med ved næste opdatering, også når du slår den fra (eller klik **Opdatér nu**). Et importeret skema får aftalen først, når du importerer det igen, og den viser så, hvornår filen blev hentet før importen. Aftalen tæller ikke med, når AulaSync ser efter, om skemaet er ændret siden importen.
- **Start AulaSync, når jeg logger ind:** slået til efter første start.
- **Logget ind som …:** dit navn og din institution. **Log ud…** glemmer dit login og sletter dine valgte skemaer, og AulaSync begynder forfra med [første start](#første-start). Kalenderfilerne bliver liggende, så kalenderne viser stadig de seneste skemaer, men de bliver ikke opdateret. Vælger du de samme skemaer igen, bliver de opdateret igen, uden at du skal tilføje dem i kalenderprogrammet på ny; klik bare **Færdig** i trin 5. Ved skemaerne står **Tilføj til Kalender** eller **Tilføj til Outlook**, indtil kalenderprogrammet henter skemaet næste gang. Klik ikke på knappen, ellers står lektionerne to gange; rækken skifter selv til **✓ Tilføjet**. Har du importeret skemaer i Outlook, siger AulaSync ikke længere til, når de ændrer sig. Slet kalenderen i Outlook, og importér skemaet igen med **Importér…**. Vil du fjerne skemaerne helt, så slet filerne i kalendermappen (**Åbn kalendermappe**) og kalenderne i dit kalenderprogram.
- **Hjælp:** **Vejledning** åbner denne vejledning. **Kom i gang igen…** viser trinene fra [første start](#første-start) igen.
- **Afinstallér AulaSync…:** fjerner AulaSync og alt, den har gemt på computeren (se [Installation](#installation)). AulaSync spørger først. Kalenderne i dit kalenderprogram skal du selv slette.
- **Fejlfinding:** **Åbn kalendermappe** og **Åbn log**, og versionsnummeret. Kalendermappen har én fil pr. skema, fx `123456-AE-medarbejder-1001.ics` (se [Ny Outlook, Outlook til Mac og Outlook på nettet](#ny-outlook-outlook-til-mac-og-outlook-på-nettet)). Filer fra AulaSync 3.0.0 hed fx `medarbejder-1001.ics`; de får det nye navn ved næste opdatering. Adressen i kalenderprogrammet er den samme som før.

## Når Aula logger dig ud

AulaSync holder forbindelsen til Aula i live og logger selv ind igen i baggrunden, når det kan. Kan det ikke, kommer beskeden "Du er logget ud af Aula. Klik her for at logge ind igen.", og hovedvinduet viser **Log ind igen**. Sker det, når AulaSync starter af sig selv, efter at du har logget ind på computeren, står der i stedet: "Klik her for at logge ind på Aula, så AulaSync kan holde dine skemaer opdateret." Indtil da viser kalenderne de seneste skemaer.

## Genveje

Tilføj skema, Opdatér nu og Luk vinduet virker i hovedvinduet. Indstillinger og Afslut AulaSync virker i alle AulaSyncs vinduer på Mac.

| | Mac | Windows |
|---|---|---|
| Tilføj skema | ⌘N | Ctrl+N |
| Opdatér nu | ⌘R | F5 |
| Luk vinduet | ⌘W | |
| Indstillinger | ⌘, | |
| Afslut AulaSync | ⌘Q | |

## Opgradering fra 2.x

AulaSync 3 er skrevet forfra og gør én ting: skemaer som kalendere.

- Beskeder synkroniseres ikke længere til Outlook. De beskeder, 2.x har hentet, ligger stadig i mappen **Indbakke › Aula** i Outlook. Vil du ikke beholde dem, så slet mappen, og tøm **Slettet post**.
- Indstillinger og valgte kalendere fra 2.x overføres ikke. Log ind, og vælg dine skemaer igen.
- Slet de gamle kalendere fra 2.x i Outlook; de bliver ikke opdateret længere. De ligger i en kalendergruppe med skolens navn og hedder fx "7A" eller "Lokale 53". Klasse- og lokalekalendere hedder det samme i AulaSync 3, så se efter, at du sletter dem i skolens gruppe.
- Kører den gamle AulaSync stadig, viser AulaSync 3 en besked. Højreklik på det gamle ikon i systembakken, og vælg **Afslut**. Åbner der i stedet vinduet "AulaSync - Log ind", så luk det; det afslutter også den gamle AulaSync. Når AulaSync 3 er installeret eller har kørt, starter kun AulaSync 3, når du logger ind.
- Den gamle AulaSync lavede genvejen **AulaSync** i Start-menuen. AulaSync 3.1.1 og nyere retter den, når den starter, så genvejen starter den nye AulaSync. Har du en ældre AulaSync 3, så hent den nyeste.
- Den gamle mappe `%USERPROFILE%\.aulasync` bruges ikke længere. Slet den, når den gamle AulaSync er afsluttet. Den rummer stadig data fra 2.x: dit gamle Aula-login, loggen og de gamle skemafiler. Afinstallerer du AulaSync 3, slettes den også.

## Fejlfinding

### Apple Kalender: "Anmodningen om at opdatere … mislykkedes"

Kalender kunne ikke nå AulaSync, da den ville opdatere kalenderen. Det sker typisk, når AulaSync ikke kører. Start AulaSync, og tryk ⌘R i Kalender, eller vent til næste automatiske opdatering. Kører AulaSync allerede, så se, om hovedvinduet viser "Kalender-server kunne ikke starte" (se nedenfor). Brug ikke **Tilføj igen**: så abonnerer Kalender en gang til, og lektionerne står to gange.

### "Kalender hentede ikke skemaet" eller "Outlook hentede ikke skemaet"

Kalenderprogrammet hentede ikke skemaet inden for 2 minutter efter klikket. Det sker fx, hvis du ikke klikkede **Abonner** i Kalender eller **Ja** i Outlook. Tjek, at kalenderprogrammet kører på samme computer som AulaSync. Klik så knappen igen, og bekræft abonnementet.

Åbnede adressen ny Outlook i stedet for klassisk Outlook, så klik ikke knappen igen. Vælg i stedet **Kopiér adresse** under **⋯** ved skemaet, og tilføj adressen i klassisk Outlook med **Tilføj kalender › Fra internettet** (se [Outlook (klassisk)](#outlook-klassisk)). Henter programmet skemaet senere, skifter rækken selv til **✓ Tilføjet**.

### "Ikke hentet af Outlook siden …"

AulaSync kan ikke se ind i dit kalenderprogram, men den ser, hvornår programmet henter hvert skema. Henter Outlook dine andre skemaer, men ikke dette, i flere timer, er kalenderen nok slettet i Outlook, og knappen kommer igen. Står kalenderen ikke længere i dit kalenderprogram, så klik knappen ved skemaet, fx **Tilføj til Outlook**, eller vælg **Tilføj igen** under **⋯**. Står den der stadig, så klik ikke; så står lektionerne to gange. Så henter programmet den bare ikke: i Apple Kalender kan den fx være sat til at opdatere sjældent (højreklik på kalenderen, og vælg **Vis info**). Vil du ikke have skemaet længere, så vælg **Fjern skema…** under **⋯**.

### "Kalender-server kunne ikke starte: port … er optaget af et andet program"

AulaSync valgte porten ved første start, og dine kalenderabonnementer bruger den. Nu bruger et andet program den, fx en anden bruger på samme computer, der også kører AulaSync. Luk det andet program, og klik **Prøv igen** i hovedvinduet. Er det en anden bruger, der er logget ind på computeren, kan du ikke selv lukke det; bed brugeren om at vælge **Afslut AulaSync** eller logge ud.

### "Login kræver Microsoft Edge WebView2 Runtime" (Windows)

Login-vinduet bruger WebView2, og den mangler på computeren. Installér den fra adressen i beskeden (eller bed IT om det). Afslut så AulaSync (højreklik på ikonet i systembakken, og vælg **Afslut AulaSync**), og start AulaSync igen.

### Kalenderen viser ikke de nyeste ændringer

AulaSync henter fra Aula hver 4. time, eller så tit, du har valgt i Indstillinger; klik **Opdatér nu** for at hente med det samme. Slå **Vis opdateringstid i kalenderen** til i Indstillinger, hvis du vil kunne se i kalenderen, hvornår skemaet sidst blev hentet. Kalenderprogrammet henter derefter fra AulaSync efter sin egen plan. I Apple Kalender kan du trykke ⌘R for at hente med det samme, eller højreklikke på kalenderen, vælge **Vis info** og ændre **Opdater automatisk**.

Bliver en kalender i Apple Kalender aldrig opdateret, så se, om den står under **iCloud** i Kalenders liste over kalendere. Så henter Apples servere den, og de kan ikke nå AulaSync på din Mac. Slet kalenderen i Kalender, vælg **Tilføj igen** i ⋯-menuen i AulaSync, og vælg **Placering: På min Mac**.

### Noget andet

Åbn loggen (Indstillinger › Fejlfinding › **Åbn log**). Linjerne med "Kalender-server" viser, om kalender-serveren startede, hvilket program der først hentede hvert skema, og hvilke forespørgsler der blev afvist. Den samme slags hentning skrives kun én gang efter hver start. Linjen "et program forsøgte at hente via https" er ikke en fejl: Apple Kalender prøver først https og henter derefter skemaet via http. Forbindelsen bliver på din computer.

## Udrulning på flere computere (Windows)

`AulaSync-Setup.exe /VERYSILENT /SUPPRESSMSGBOXES` installerer AulaSync uden vinduer for den bruger, der kører den, i `%LOCALAPPDATA%\Programs\AulaSync`, og lægger genvejen `AulaSync` i brugerens Start-menu. Det kræver ingen administratorrettigheder. Kører installationen som administrator, installeres AulaSync for administratoren, ikke for brugeren. `/LAUNCH=1` åbner AulaSync efter første installation. Ved en opdatering startes AulaSync igen, hvis den kørte, men ikke når installationen kører som administrator. Afinstallation uden vinduer: `"%LOCALAPPDATA%\Programs\AulaSync\unins000.exe" /VERYSILENT`. Den fjerner alt uden at spørge, også brugerens data i `%LOCALAPPDATA%\AulaSync` og `%USERPROFILE%\.aulasync` og start ved login. Vil I hellere pakke AulaSync selv, fx til Intune, ligger den løse `AulaSync.exe` også under hver udgave; den kræver ingen installation.

`AulaSync.exe --silent` starter AulaSync i systembakken uden vindue. Har brugeren ikke gennemført første start, eller kan AulaSync ikke logge ind af sig selv, viser ikonet en prik, og der kommer en besked: "Klik her for at logge ind på Aula, så AulaSync kan holde dine skemaer opdateret." Et klik åbner første start, hvis den ikke er gennemført, og ellers login.

- **Start ved login** sættes først op, når brugeren klikker **Færdig** i første start eller slår **Start AulaSync, når jeg logger ind** til i Indstillinger. Findes værdien allerede, fx fra AulaSync 2 eller en tidligere AulaSync 3, retter AulaSync og installationsprogrammet den med det samme, så AulaSync 3 starter ved login også før første start. AulaSync skriver `AulaSync` med indholdet `"<sti til AulaSync.exe>" --silent` under `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`. Skal AulaSync starte ved login før første start, kan I selv skrive den samme værdi.
- Start ikke også AulaSync på anden vis, fx fra `HKLM` eller mappen Start. Så starter AulaSync to gange, og ved den anden start viser AulaSync et vindue: første start, hvis brugeren ikke har gennemført den, og ellers hovedvinduet.
- Startværdien peger på den `AulaSync.exe`, der kører. Starter brugeren AulaSync fra et nyt sted, retter AulaSync selv værdien, og installationsprogrammet retter den til den installerede AulaSync. Afinstallation fjerner værdien, uanset hvilken `AulaSync.exe` den peger på.
- AulaSync lægger genvejen `AulaSync.lnk` i brugerens egen Start-menu, `%APPDATA%\Microsoft\Windows\Start Menu\Programs`. Ved hver start laver AulaSync den igen, hvis den mangler, og retter den, hvis den peger på en anden fil. Det kan ikke slås fra. I behøver ikke selv lave en genvej. Installationsprogrammet laver den samme genvej, og afinstallation fjerner den.
- `AulaSync.exe --quit` beder en kørende AulaSync om at afslutte. Det bruger installationsprogrammet før en opdatering.
- Alt, også login, gemmes i `%LOCALAPPDATA%\AulaSync`, som ikke følger med en roaming-profil. På en ny computer skal brugeren igennem første start igen.
- Hver bruger får sin egen port ved første start: 9876 eller den første ledige op til 9899. Porten reserveres ikke. Var to brugere ikke logget ind samtidig, da de startede AulaSync første gang, kan de få samme port, og så kan kun den ene starte kalender-serveren, når begge er logget ind.
- Kalender-serveren svarer kun på computeren selv, men den har ingen adgangskode. Andre brugere på samme computer, fx på en terminalserver, kan derfor hente en brugers skemaer, hvis de kender porten og skemaets adresse.
- Login kræver Microsoft Edge WebView2 Runtime (se [Installation](#windows)).
- Klassisk Outlook skal have lov til at abonnere på internetkalendere. Det må ikke være slået fra med en gruppepolitik.
