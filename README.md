# AulaSync

Synkroniserer beskeder fra [Aula](https://www.aula.dk) til Microsoft Outlook. Virker med alle kommuners login (UniLogin, Azure AD, MitID m.fl.).

## Features

- **Browser-baseret login** via WebView2 - understøtter alle kommuners IdP
- **Korrekte tidsstempler** - beskeder vises med Aulas dato, ikke importtidspunktet (via direkte MAPI)
- **Korrekte afsendere** - viser den rigtige afsender i Outlook
- **Automatisk synkronisering** - konfigurerbart interval (2 min - 1 time)
- **System tray** - kører diskret i baggrunden
- **Autostart** - kan starte automatisk ved Windows-login
- **Read-only beskeder** - importerede beskeder kan ikke redigeres ved et uheld

## Installation

1. Download `AulaSync.exe` fra [Releases](https://github.com/rpaasch/AulaSync/releases)
2. Kør filen

## Brug

1. Start AulaSync - et login-vindue åbner
2. Log ind på Aula som du plejer (vælg den rigtige rolle/kommune)
3. Vinduet lukker automatisk når login er færdigt
4. Beskeder importeres til mappen **Indbakke → Aula** i Outlook
5. AulaSync kører i system tray og synkroniserer automatisk

## Tray-menu

| Menupunkt | Beskrivelse |
|-----------|-------------|
| **Hent nu** | Synkronisér med det samme |
| **Indstillinger** | Synkinterval, autostart, nulstil historik, log, vejledning |
| **Log ud** | Log ud og vis login igen |
| **Afslut** | Luk programmet |

## Teknisk

- **.NET 8** WinForms app med WebView2
- **Outlook COM** til oprettelse af beskeder
- **Direkte MAPI** (`mapi32.dll` P/Invoke) til korrekte tidsstempler og afsendere
- Beskeder oprettes som PostItems (read-only) i en "Aula"-undermappe
- Session-cookies hentes fra WebView2 og overføres til HttpClient
- Deduplicering via `seen_threads.json`

## Krav

- Windows 10/11
- Microsoft Outlook (klassisk, ikke ny Outlook)
- Edge WebView2 Runtime (følger med Windows 11)

## Byg selv

```bash
dotnet publish -c Release -r win-x64 --self-contained -o publish
```

## Data

Konfiguration gemmes i `%USERPROFILE%\.aulasync\`:

| Fil | Indhold |
|-----|---------|
| `config.json` | Synkinterval, autostart |
| `aulasync.log` | Logfil |
| `seen_threads.json` | Importerede beskeder |
| `webview2/` | Browser-session |
