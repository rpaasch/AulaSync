AulaSync 3 er en ny udgave med ét formål: skemaer fra Aula, der holder sig opdateret i din kalender, på Windows og Mac.

**Nyt i 3.2.1**
- Mac: AulaSync er nu signeret af udvikleren og undersøgt af Apple (notariseret). Første gang spørger macOS kun, om du vil åbne en app, der er hentet fra internettet; klik *Åbn*. Du skal ikke længere vælge *Åbn alligevel* i Systemindstillinger › Anonymitet & sikkerhed.
- Mac: AulaSync kan installeres med Homebrew: `brew install --cask rpaasch/tap/aulasync`, og en ny udgave får du med `brew upgrade --cask rpaasch/tap/aulasync`. macOS spørger kun første gang, ikke ved de næste opgraderinger.
- Mac: *Afinstallér AulaSync…* fjerner også Homebrews optegnelse af AulaSync, så Homebrew ikke længere regner den for installeret.

**Nyt i 3.2.0**
- Windows: AulaSync har fået et installationsprogram, `AulaSync-Setup.exe`. Det installerer AulaSync for dig uden administratorrettigheder, lægger den i Start-menuen og åbner den. `winget install rpaasch.AulaSync` gør det samme. Opdatering med `AulaSync-Setup.exe` lukker en kørende AulaSync og starter den igen.
- AulaSync har fået sit eget ikon på Windows og Mac, også i systembakken og menulinjen. Har du fastgjort en ældre AulaSync til proceslinjen, så frigør den, og fastgør AulaSync igen fra Start-menuen.
- *Vis opdateringstid i kalenderen*: aftalen mandag kl. 5.45 viser nu kort, hvornår skemaet blev hentet, fx "Opd. 091026@07:30", og beskrivelsen siger det samme.
- Har du en ældre AulaSync uden installation, så installér `AulaSync-Setup.exe`. Start ved login og genvejen i Start-menuen flytter med, og bagefter kan du slette den gamle `AulaSync.exe`. Har du AulaSync fra winget, så vælg *Afslut AulaSync*, skriv `winget upgrade rpaasch.AulaSync`, og start AulaSync fra Start-menuen.
- *Afinstallér AulaSync…* i Indstillinger fjerner alt, AulaSync har gemt på computeren: programmet, start ved login, dit login, dine indstillinger, loggen og kalenderfilerne. Det samme sker på Windows fra *Indstillinger › Apps* og med `winget uninstall rpaasch.AulaSync`. Har du AulaSync fra winget fra før 3.2, sletter de to kun programmet; brug *Afinstallér AulaSync…*. Slet bagefter AulaSync-kalenderne i dit kalenderprogram.
- Har du slettet en kalender i dit kalenderprogram, opdager AulaSync det: henter programmet dine andre skemaer, men ikke dette, står der fx "Ikke hentet af Outlook siden 5. okt.", og ved Outlook kommer knappen igen.

**Nyt i 3.1.1**
- Windows: AulaSync ligger i Start-menuen, når den har kørt første gang, så du kan finde den ved at skrive *AulaSync* i Start. Det gælder også, når du har installeret den med winget. En genvej fra den gamle AulaSync 2 bliver rettet, så den starter den nye.

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

**Sådan virker det**
- Vælg medarbejdere, klasser og lokaler. Hvert skema bliver sin egen kalender.
- Apple Kalender, klassisk Outlook og andre kalenderprogrammer på samme computer holdes opdateret automatisk, så længe AulaSync kører.
- Ny Outlook, Outlook til Mac og Outlook på nettet henter kalendere gennem Microsofts servere, som ikke kan nå din computer. Her importerer du en fil, og AulaSync siger til, når skemaet er ændret.

**Installation**
- Windows 10 og 11 (64-bit): hent `AulaSync-Setup.exe` herunder, og åbn den. Viser Windows "Windows beskyttede din pc", så vælg *Flere oplysninger* › *Kør alligevel*. Du kan også bruge `winget install rpaasch.AulaSync`; en ny udgave kommer først i winget, når Microsoft har godkendt den. `AulaSync.exe` er udgaven uden installation, til ældre installationer fra winget og til IT.
- Mac (macOS 15 eller nyere): hent `-arm64.dmg` til Mac med Apple-chip (M1 og nyere) eller `-x64.dmg` til Mac med Intel, og træk AulaSync over i Programmer. Åbn AulaSync fra Programmer (ikke fra `.dmg`-filen), og klik *Åbn*, når macOS spørger, om du vil åbne en app, der er hentet fra internettet. Med Homebrew: `brew install --cask rpaasch/tap/aulasync`. Når Kalender spørger, så vælg Placering: *På min Mac*.

AulaSync er ikke tilknyttet eller godkendt af Aula, KOMBIT, Netcompany eller KMD.
