using AulaSync.Core;

namespace AulaSync.App;

// Hovedknappen pr. skema afhænger af kalenderprogrammet (spec §3.2, tabellen "Kalenderprogram").
public sealed class MainActions(SyncService sync, ConfigStore config, IPlatform platform, IDialogs dialogs, TimeProvider? time = null)
    : IMainActions
{
    public async Task RunAsync(ScheduleRef schedule)
    {
        var app = CalendarChoice.Current(config, platform.IsMac);
        var port = config.Load().ServerPort;
        switch (CalendarApps.ActionFor(app).Kind)
        {
            case MainActionKind.Subscribe:
                // Markér først: Kalender henter filen, så snart brugeren klikker "Abonnér", og hentningen skal tælle.
                await sync.MarkAddedAsync(schedule);
                // Apple Kalender får http-adressen (webcal:// bliver til https, som AulaSync ikke taler); Outlook får webcal://.
                // Outlook får en ny adresse ved hvert klik (tidspunktet efter "?tilfoejet="; serveren ser bort fra det):
                // Outlook husker abonnementer pr. adresse, og blev en kalender slettet i Outlook og tilføjet igen med samme
                // adresse, fik den lektionerne, men intet navn (brugeren, 2026-10-13). Med en ny adresse laver Outlook et nyt
                // abonnement med navnet fra filen.
                if (app == CalendarApp.AppleCalendar) platform.OpenInCalendar(IcsServer.UrlFor(schedule, port));
                else platform.Open(IcsServer.WebcalFor(schedule, port, (time ?? TimeProvider.System).GetUtcNow().ToUnixTimeSeconds()));
                // webcal:// på Windows kan ende i ny Outlook; derfor tilbydes vejen "Fra internettet".
                if (app == CalendarApp.OutlookClassic) await dialogs.ShowOutlookFallbackAsync(schedule);
                break;
            case MainActionKind.Import:
                if (await dialogs.ShowImportGuideAsync(schedule)) await sync.MarkImportedAsync(schedule);
                break;
            case MainActionKind.CopyAddress:
                if (await platform.CopyTextAsync(IcsServer.UrlFor(schedule, port))) await sync.MarkAddedAsync(schedule);
                break;
        }
    }

    public Task CopyAddressAsync(ScheduleRef schedule) => platform.CopyTextAsync(IcsServer.UrlFor(schedule, config.Load().ServerPort));

    public void RevealFile(ScheduleRef schedule) => platform.RevealFile(sync.FilePath(schedule));
}
