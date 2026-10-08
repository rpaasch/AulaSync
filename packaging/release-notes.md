AulaSync 3 er en ny udgave med ét formål: skemaer fra Aula, der holder sig opdateret i din kalender, på Windows og Mac.

**Nyt i 3.1**
- Vælg selv, hvor tit skemaerne hentes fra Aula: hver halve time, hver time, hver 2., 4. eller 8. time. Standard er nu hver 4. time (før hver 6.). Sjældnere giver færre forespørgsler til Aula. Se *Indstillinger*.
- *Vis opdateringstid i kalenderen*: en privat aftale mandag kl. 5.45 i hver kalender viser, hvornår skemaet sidst blev hentet. Slået fra som standard.
- Kalenderfilerne har institutionens nummer og initialer eller navn i navnet, fx `123456-AE-medarbejder-1001.ics`, så de er lette at finde i Stifinder og Finder, også når du importerer i Outlook. Filerne får det nye navn ved næste opdatering; adressen i kalenderprogrammet er den samme, så abonnementer virker som før.
- *Hjælp* i Indstillinger: *Vejledning* åbner vejledningen, og *Kom i gang igen…* viser trinene fra første start igen.

**Vigtigt, hvis du kommer fra 2.x**
- Beskeder synkroniseres ikke længere til Outlook.
- Indstillinger og valgte kalendere fra 2.x overføres ikke. Log ind, og vælg dine skemaer igen.
- Slet de gamle kalendere fra 2.x i Outlook (i kalendergruppen med skolens navn); de bliver ikke opdateret længere.
- Kører den gamle AulaSync stadig, siger AulaSync 3 til. Højreklik på det gamle ikon i systembakken, og vælg *Afslut*.
- Den gamle AulaSync lavede genvejen *AulaSync* i Start-menuen. Ligger AulaSync 3 i en anden mappe end den gamle, starter genvejen stadig den gamle. Slet så genvejen *AulaSync* i mappen `%APPDATA%\Microsoft\Windows\Start Menu\Programs` (skriv adressen i Stifinder).

**Sådan virker det**
- Vælg medarbejdere, klasser og lokaler. Hvert skema bliver sin egen kalender.
- Apple Kalender, klassisk Outlook og andre kalenderprogrammer på samme computer holdes opdateret automatisk, så længe AulaSync kører.
- Ny Outlook, Outlook til Mac og Outlook på nettet henter kalendere gennem Microsofts servere, som ikke kan nå din computer. Her importerer du en fil, og AulaSync siger til, når skemaet er ændret.

**Installation**
- Windows 10 og 11 (64-bit): hent `AulaSync.exe` herunder, læg den i en fast mappe, fx *Dokumenter*, og dobbeltklik på den dér. Flyt den ikke bagefter, for AulaSync starter derfra, når du logger ind. Viser Windows "Windows beskyttede din pc", så vælg *Flere oplysninger* › *Kør alligevel*. Du kan også bruge `winget install rpaasch.AulaSync`, men en ny udgave kommer først i winget, når Microsoft har godkendt den. Installerer winget version 2.x, så hent `AulaSync.exe` herunder i stedet.
- Mac (macOS 15 eller nyere): hent `-arm64.dmg` til Mac med Apple-chip (M1 og nyere) eller `-x64.dmg` til Mac med Intel, og træk AulaSync over i Programmer. Appen er ikke signeret af Apple, så første gang skal du give lov: åbn AulaSync fra Programmer (ikke fra `.dmg`-filen), og klik *Udført*, når macOS siger, at appen ikke blev åbnet. Åbn så *Systemindstillinger › Anonymitet og sikkerhed*, klik *Åbn alligevel* ved AulaSync, og bekræft med adgangskoden til din Mac. Når Kalender spørger, så vælg Placering: *På min Mac*.

AulaSync er ikke tilknyttet eller godkendt af Aula, KOMBIT, Netcompany eller KMD.
