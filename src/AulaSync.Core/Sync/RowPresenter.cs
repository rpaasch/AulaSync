namespace AulaSync.Core;

// Primary: hovedknappens tekst (null = ingen knap). Done: dæmpet status. Note: mærke ved siden af, fx "Ændret siden import".
// Waiting: Done er "Venter på …" og vises gråt, ikke grønt som "✓ Tilføjet".
public sealed record RowButtons(string? Primary, string? Done, string? Note, string Again, bool Waiting = false);

public static class RowPresenter
{
    // Så længe venter en række på, at kalenderprogrammet henter skemaet efter "Tilføj til …".
    public static readonly TimeSpan FetchWait = TimeSpan.FromMinutes(2);

    // fetch: henter kalenderprogrammet stadig skemaet (SyncService.CheckFetch); null: ved det ikke (fx kører serveren ikke).
    public static RowButtons Buttons(CalendarApp app, Subscription subscription, bool changedSinceImport, DateTimeOffset now, TimeZoneInfo tz,
        FetchCheck? fetch = null)
    {
        var action = CalendarApps.ActionFor(app);
        if (action.Kind == MainActionKind.Import)
        {
            if (subscription.ImportedAt is not { } at) return new(action.Label, null, null, "Importér igen…");
            var done = $"Importeret {Day(at, now, tz)}";
            return changedSinceImport ? new("Importér igen…", done, "Ændret siden import", "Importér igen…") : new(null, done, null, "Importér igen…");
        }
        // "✓ Tilføjet" først, når kalenderprogrammet har hentet filen fra AulaSync, og kun så længe det bliver ved med det.
        // Henter Outlook dine andre skemaer, men ikke dette, er kalenderen nok slettet: så kommer knappen igen. Andre
        // programmer kan hente hver kalender for sig sjældent (Apple Kalender fx én gang om ugen), så dér kun mærket; et
        // klik på knappen, mens kalenderen findes, ville give lektionerne to gange.
        if (subscription.Fetched)
        {
            if (fetch is not { Health: not FetchHealth.Ok, LastFetched: { } last }) return new(null, "✓ Tilføjet", null, "Tilføj igen");
            var by = fetch.Program is { } name ? $" af {name}" : "";
            var note = $"Ikke hentet{by} siden {Since(last, now, tz)}";
            return fetch is { Health: FetchHealth.Missing, Program: "Outlook" }
                ? new(action.Label, null, note, "Tilføj igen")
                : new(null, null, note, "Tilføj igen");
        }
        if (subscription.AddedAt is not { } added) return new(action.Label, null, null, "Tilføj igen");
        // Ved "Kopiér adresse" ved AulaSync ikke, hvornår adressen bliver sat ind, så der er ingen frist.
        if (action.Kind == MainActionKind.CopyAddress) return new(null, "✓ Kopieret", null, "Tilføj igen");
        var program = app == CalendarApp.OutlookClassic ? "Outlook" : "Kalender";
        return now - added < FetchWait
            ? new(null, $"Venter på {program}…", null, "Tilføj igen", Waiting: true)
            : new(action.Label, null, $"{program} hentede ikke skemaet", "Tilføj igen");
    }

    // Hvornår rækken skifter af sig selv (fra "Venter på …" til "… hentede ikke skemaet", eller til "Ikke hentet …");
    // null, hvis den ikke gør det.
    public static DateTimeOffset? ChangesAt(CalendarApp app, Subscription subscription, DateTimeOffset now, FetchCheck? fetch = null)
    {
        var kind = CalendarApps.ActionFor(app).Kind;
        if (kind == MainActionKind.Import) return null;
        if (subscription.Fetched) return fetch?.ChangesAt;
        return kind == MainActionKind.Subscribe && subscription.AddedAt is { } added && now - added < FetchWait ? added + FetchWait : null;
    }

    // "kl. 08:12" i dag, "i går kl. 21:21", ellers "5. okt.".
    static string Since(DateTimeOffset at, DateTimeOffset now, TimeZoneInfo tz)
    {
        var day = TimeZoneInfo.ConvertTime(at, tz).Date;
        var today = TimeZoneInfo.ConvertTime(now, tz).Date;
        var time = TimeZoneInfo.ConvertTime(at, tz).ToString("HH:mm", System.Globalization.CultureInfo.InvariantCulture);
        return day == today ? $"kl. {time}" : day == today.AddDays(-1) ? $"i går kl. {time}" : StatusText.ShortDate(at, tz);
    }

    static string Day(DateTimeOffset at, DateTimeOffset now, TimeZoneInfo tz) =>
        TimeZoneInfo.ConvertTime(at, tz).Date == TimeZoneInfo.ConvertTime(now, tz).Date ? "i dag" : StatusText.ShortDate(at, tz);
}
