# AulaSync

Synkroniserer beskeder og skemaer fra [Aula](https://www.aula.dk) til Microsoft Outlook. Virker med alle kommuners login (UniLogin, Azure AD, MitID m.fl.).

## Features

### Beskeder
- **Korrekte tidsstempler** - beskeder vises med Aulas dato, ikke importtidspunktet
- **Korrekte afsendere** - viser den rigtige afsender i Outlook
- **Read-only** - importerede beskeder kan ikke redigeres ved et uheld
- **Automatisk synkronisering** - konfigurerbart interval (2 min - 1 time)

### Skemaer
- **Medarbejderskemaer** - se kollegers ugeskema direkte i Outlook
- **Klasseskemaer** - se en klasses ugeskema
- **Lokaleskemaer** - se hvad der sker i et lokale
- Format tilpasset kontekst: `FAG; Lokale; Klasse` / `FAG; Lokale; INIT` / `FAG; Klasse; INIT`
- 3 måneder frem + 3 måneder tilbage
- Automatisk opdatering hver time

### Generelt
- **Browser-baseret login** via WebView2 - understøtter alle kommuners IdP
- **System tray** - kører diskret i baggrunden
- **Autostart** - kan starte automatisk ved Windows-login

## Installation

1. Download `AulaSync.exe` fra [Releases](https://github.com/rpaasch/AulaSync/releases)
2. Kør filen - ingen installation nødvendig

Eller via winget (når godkendt):
```
winget install rpaasch.AulaSync
```

## Vejledning

### Første gang

1. Start `AulaSync.exe`
2. Et login-vindue åbner med Aula - log ind som du plejer
3. Vælg den rigtige rolle (medarbejder/forælder) og institution
4. Vinduet lukker automatisk når login er færdigt
5. Beskeder importeres til mappen **Indbakke → Aula** i Outlook
6. AulaSync kører i system tray (ikonet ved uret)

### Tray-menu

| Menupunkt | Beskrivelse |
|-----------|-------------|
| **Synkronisér beskeder** | Hent nye beskeder med det samme |
| **Medarbejderskemaer til Outlook** | Vælg en kollega og se deres skema |
| **Klasseskemaer til Outlook** | Vælg en klasse og se ugeskemaet |
| **Lokaleskemaer til Outlook** | Vælg et lokale og se bookinger |
| **Indstillinger** | Synkinterval, autostart, log, vejledning m.m. |
| **Log ud** | Log ud og vis login igen |
| **Afslut** | Luk programmet |

### Beskeder i Outlook

- Beskeder placeres i mappen **Indbakke → Aula**
- Viser korrekt dato og afsender
- Klik "Åbn i Aula" i beskeden for at se den direkte i Aula

### Skemaer i Outlook

- Skemaer placeres i kalendergruppen **Skemaer {Institution}**
- Viser fag, lokale, klasse og lærer direkte i kalendervisningen
- Opdateres automatisk hver time

## Teknisk

- **.NET 8** WinForms app med WebView2
- **Outlook COM** til oprettelse af beskeder og kalenderaftaler
- **Direkte MAPI** (`mapi32.dll` P/Invoke) til korrekte tidsstempler
- Aktiv profil hentes via `profiles.getProfileContext` (korrekt multi-institution)
- Kalenderdata hentes i bidder af max 42 dage (API-begrænsning)
- Retry-logik med exponential backoff på alle HTTP-kald
- Atomisk fil-skrivning for data-integritet
- Log-rotation ved 1 MB

## Krav

- Windows 10/11
- Microsoft Outlook (klassisk, ikke ny Outlook)
- Edge WebView2 Runtime (følger med Windows 11)

## Byg selv

```bash
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -o publish
```

## Data

Konfiguration gemmes i `%USERPROFILE%\.aulasync\`:

| Fil | Indhold |
|-----|---------|
| `config.json` | Synkinterval m.m. |
| `aulasync.log` | Logfil (max 1 MB, roteres) |
| `seen_threads.json` | Importerede beskeder |
| `subscribed_calendars.json` | Abonnerede skemaer |
| `webview2/` | Browser-session (slettet ved logout) |
