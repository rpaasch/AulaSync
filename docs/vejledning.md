# Vejledning til AulaSync

AulaSync henter skemaer fra Aula og holder dem opdateret som kalendere i dit kalenderprogram. Du vælger selv, hvilke skemaer du vil have: dit eget, kollegers, klassers og lokalers.

## Før du starter

- **AulaSync og kalenderprogrammet skal køre på samme computer.** Kalenderprogrammet henter skemaerne fra AulaSync på din egen computer (`http://localhost:9876`, eller en af de næste porte, hvis 9876 var optaget ved første start). Derfor ser du skemaerne på denne computer, ikke på telefonen eller i en kalender på nettet. Undtagelsen er import i ny Outlook, Outlook til Mac og Outlook på nettet: Så ligger skemaet i din Outlook-konto hos Microsoft og kan også ses på telefonen, men det opdateres ikke af sig selv (se [Kalenderprogrammer](#kalenderprogrammer)).
- **Kalenderne opdateres kun, mens AulaSync kører.** AulaSync starter selv, når du logger ind på computeren, og kører i menulinjen (Mac) eller systembakken (Windows).
- Du skal kunne logge ind på Aula som medarbejder.
- Windows 10 eller 11 (64-bit), eller macOS 15 eller nyere.

## Installation

### Windows

1. Hent `AulaSync.exe` fra [den nyeste udgave](https://github.com/rpaasch/AulaSync/releases/latest). Filen står under **Assets**. Advarer browseren om, at filen ikke hentes ofte, så vælg at beholde den.
2. Flyt filen fra **Overførsler** til en mappe, du beholder, fx **Dokumenter**, og dobbeltklik på den dér. Det kræver ingen installation og ingen administratorrettigheder.
3. Viser Windows "Windows beskyttede din pc", så klik **Flere oplysninger** › **Kør alligevel**. Det sker, fordi programmet ikke er signeret, så Windows ikke kan se, hvem der har udgivet det. Gør det kun med en `AulaSync.exe`, du selv har hentet fra udgivelsessiden.

Mangler **Kør alligevel**, eller siger Windows, at programmet er blokeret, har skolens IT lukket for programmer, de ikke har godkendt. Så spørg IT.

AulaSync starter fra den mappe, filen ligger i, når du logger ind. Flytter eller sletter du filen, bliver kalenderne ikke længere opdateret. Har du flyttet den, så afslut AulaSync først: højreklik på AulaSync-ikonet i systembakken ved uret, og vælg **Afslut AulaSync**. Start den så fra den nye mappe, og slå **Start AulaSync, når jeg logger ind** fra og til igen i Indstillinger. Kommer der en ny udgave, så afslut AulaSync, og erstat filen i samme mappe.

Når AulaSync har kørt første gang, ligger den i Start-menuen, så du kan finde den ved at skrive *AulaSync* i Start. Flytter du filen, retter AulaSync selv genvejen, når du starter AulaSync fra den nye mappe.

**Med winget:** Højreklik på Start-knappen, vælg **Terminal** (på Windows 10: **Windows PowerShell**), og skriv `winget install rpaasch.AulaSync`. Spørger winget, om du accepterer vilkårene, så svar ja. winget laver ingen genvej i Start-menuen, så første gang starter du AulaSync ved at skrive `AulaSync` i et nyt vindue (Terminal eller PowerShell). Derefter ligger AulaSync i Start-menuen. En ny udgave kommer først i winget, når Microsoft har godkendt den. Når der kommer en ny udgave, så højreklik på AulaSync-ikonet i systembakken ved uret, og vælg **Afslut AulaSync**. Skriv så `winget upgrade rpaasch.AulaSync` i et nyt vindue (Terminal eller PowerShell), og start AulaSync igen ved at skrive `AulaSync`.

Login bruger Microsoft Edge WebView2 Runtime, som følger med Windows 11 og findes på de fleste computere med Windows 10. Mangler den, står der i login-vinduet, hvor du henter den.

### Mac

1. Hent den rigtige `.dmg`-fil fra [den nyeste udgave](https://github.com/rpaasch/AulaSync/releases/latest). Filerne står under **Assets**: `AulaSync-<version>-arm64.dmg` er til Mac med Apple-chip (M1 og nyere), og `AulaSync-<version>-x64.dmg` er til Mac med Intel-processor. Er du i tvivl, så klik på æblet øverst til venstre på skærmen, og vælg **Om denne Mac**. Står der **Chip** og **Apple M…**, skal du bruge `arm64`. Står der **Processor** og **Intel**, skal du bruge `x64`.
2. Åbn filen, og træk AulaSync over i **Programmer**.
3. Åbn AulaSync fra **Programmer**, ikke fra `.dmg`-filen. Ellers kan AulaSync ikke starte af sig selv, når du logger ind. Første gang siger macOS, at AulaSync ikke blev åbnet, fordi Apple ikke har kunnet kontrollere appen. Klik **Udført** (ikke **Flyt til papirkurv**).
4. Åbn **Systemindstillinger › Anonymitet og sikkerhed**, rul ned, og klik **Åbn alligevel** ved AulaSync. Knappen står der kun i cirka en time, efter at du har prøvet at åbne appen. Er den væk, så gentag trin 3. Spørger macOS igen, så klik **Åbn alligevel**, og skriv adgangskoden til din Mac (eller brug Touch ID).

Giv kun lov, hvis du selv har hentet `.dmg`-filen fra udgivelsessiden. Beder macOS om en administrators navn og adgangskode, når du trækker AulaSync over i **Programmer**, eller må du ikke godkende apps, fx på en skole-Mac, som IT styrer, så bed IT om hjælp.

Kommer der en ny udgave, så vælg **Afslut AulaSync** i ikonets menu. Hent den nye `.dmg`-fil, og træk AulaSync over i **Programmer**, så den erstatter den gamle. Åbn den fra **Programmer**, og giv lov igen som i trin 3 og 4, hvis macOS spørger.

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
- **Vis opdateringstid i kalenderen:** Slået fra som standard. Slår du den til, får hver af AulaSyncs kalendere en privat aftale mandag kl. 5.45-6.00, fx "AulaSync opdateret ons. 7. okt. 14:32". Så kan du i kalenderen se, hvornår skemaet sidst blev hentet fra Aula. Lørdag og søndag står aftalen i den kommende uge. Den optager ikke tid i kalenderen. Ændringen kommer med ved næste opdatering, også når du slår den fra (eller klik **Opdatér nu**). Kalenderprogrammet viser den, næste gang det henter fra AulaSync; i Apple Kalender kan du trykke ⌘R. Et importeret skema får aftalen først, når du importerer det igen, og den viser så, hvornår filen blev hentet før importen. Aftalen tæller ikke med, når AulaSync ser efter, om skemaet er ændret siden importen.
- **Start AulaSync, når jeg logger ind:** slået til efter første start.
- **Logget ind som …:** dit navn og din institution. **Log ud…** glemmer dit login og sletter dine valgte skemaer, og AulaSync begynder forfra med [første start](#første-start). Kalenderfilerne bliver liggende, så kalenderne viser stadig de seneste skemaer, men de bliver ikke opdateret. Vælger du de samme skemaer igen, bliver de opdateret igen, uden at du skal tilføje dem i kalenderprogrammet på ny; klik bare **Færdig** i trin 5. Ved skemaerne står **Tilføj til Kalender** eller **Tilføj til Outlook**, indtil kalenderprogrammet henter skemaet næste gang. Klik ikke på knappen, ellers står lektionerne to gange; rækken skifter selv til **✓ Tilføjet**. Har du importeret skemaer i Outlook, siger AulaSync ikke længere til, når de ændrer sig. Slet kalenderen i Outlook, og importér skemaet igen med **Importér…**. Vil du fjerne skemaerne helt, så slet filerne i kalendermappen (**Åbn kalendermappe**) og kalenderne i dit kalenderprogram.
- **Hjælp:** **Vejledning** åbner denne vejledning. **Kom i gang igen…** viser trinene fra [første start](#første-start) igen.
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
- Kører den gamle AulaSync stadig, viser AulaSync 3 en besked. Højreklik på det gamle ikon i systembakken, og vælg **Afslut**. Åbner der i stedet vinduet "AulaSync - Log ind", så luk det; det afslutter også den gamle AulaSync. Når du har gennemført første start, starter kun AulaSync 3, når du logger ind.
- Den gamle AulaSync lavede genvejen **AulaSync** i Start-menuen. AulaSync 3.1.1 og nyere retter den, når den starter, så genvejen starter den nye AulaSync. Har du en ældre AulaSync 3, så hent den nyeste.
- Den gamle mappe `%USERPROFILE%\.aulasync` bruges ikke længere. Slet den, når den gamle AulaSync er afsluttet. Den rummer stadig data fra 2.x: dit gamle Aula-login, loggen og de gamle skemafiler.

## Fejlfinding

### Apple Kalender: "Anmodningen om at opdatere … mislykkedes"

Kalender kunne ikke nå AulaSync, da den ville opdatere kalenderen. Det sker typisk, når AulaSync ikke kører. Start AulaSync, og tryk ⌘R i Kalender, eller vent til næste automatiske opdatering. Kører AulaSync allerede, så se, om hovedvinduet viser "Kalender-server kunne ikke starte" (se nedenfor). Brug ikke **Tilføj igen**: så abonnerer Kalender en gang til, og lektionerne står to gange.

### "Kalender hentede ikke skemaet" eller "Outlook hentede ikke skemaet"

Kalenderprogrammet hentede ikke skemaet inden for 2 minutter efter klikket. Det sker fx, hvis du ikke klikkede **Abonner** i Kalender eller **Ja** i Outlook. Tjek, at kalenderprogrammet kører på samme computer som AulaSync. Klik så knappen igen, og bekræft abonnementet.

Åbnede adressen ny Outlook i stedet for klassisk Outlook, så klik ikke knappen igen. Vælg i stedet **Kopiér adresse** under **⋯** ved skemaet, og tilføj adressen i klassisk Outlook med **Tilføj kalender › Fra internettet** (se [Outlook (klassisk)](#outlook-klassisk)). Henter programmet skemaet senere, skifter rækken selv til **✓ Tilføjet**.

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

`AulaSync.exe --silent` starter AulaSync i systembakken uden vindue. Har brugeren ikke gennemført første start, eller kan AulaSync ikke logge ind af sig selv, viser ikonet en prik, og der kommer en besked: "Klik her for at logge ind på Aula, så AulaSync kan holde dine skemaer opdateret." Et klik åbner første start, hvis den ikke er gennemført, og ellers login.

- **Start ved login** sættes først op, når brugeren klikker **Færdig** i første start eller slår **Start AulaSync, når jeg logger ind** til i Indstillinger. AulaSync skriver så værdien `AulaSync` med indholdet `"<sti til AulaSync.exe>" --silent` under `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`. Skal AulaSync starte ved login før første start, kan I selv skrive den samme værdi.
- Start ikke også AulaSync på anden vis, fx fra `HKLM` eller mappen Start. Så starter AulaSync to gange, og ved den anden start viser AulaSync et vindue: første start, hvis brugeren ikke har gennemført den, og ellers hovedvinduet.
- Startværdien peger på den sti, `AulaSync.exe` lå på, da brugeren klikkede **Færdig**. Læg filen et fast sted.
- AulaSync lægger genvejen `AulaSync.lnk` i brugerens egen Start-menu, `%APPDATA%\Microsoft\Windows\Start Menu\Programs`. Ved hver start laver AulaSync den igen, hvis den mangler, og retter den, hvis den peger på en anden fil. Det kan ikke slås fra. I behøver ikke selv lave en genvej.
- Alt, også login, gemmes i `%LOCALAPPDATA%\AulaSync`, som ikke følger med en roaming-profil. På en ny computer skal brugeren igennem første start igen.
- Hver bruger får sin egen port ved første start: 9876 eller den første ledige op til 9899. Porten reserveres ikke. Var to brugere ikke logget ind samtidig, da de startede AulaSync første gang, kan de få samme port, og så kan kun den ene starte kalender-serveren, når begge er logget ind.
- Kalender-serveren svarer kun på computeren selv, men den har ingen adgangskode. Andre brugere på samme computer, fx på en terminalserver, kan derfor hente en brugers skemaer, hvis de kender porten og skemaets adresse.
- Login kræver Microsoft Edge WebView2 Runtime (se [Installation](#windows)).
- Klassisk Outlook skal have lov til at abonnere på internetkalendere. Det må ikke være slået fra med en gruppepolitik.
