namespace AulaSync.Core;

// Primary: hovedknappens tekst (null = ingen knap). Done: dæmpet status. Note: mærke ved siden af, fx "Ændret siden import".
// Waiting: Done er "Venter på …" og vises gråt, ikke grønt som "✓ Tilføjet".
public sealed record RowButtons(string? Primary, string? Done, string? Note, string Again, bool Waiting = false);

public static class RowPresenter
{
    // Så længe venter en række på, at kalenderprogrammet henter skemaet efter "Tilføj til …".
    public static readonly TimeSpan FetchWait = TimeSpan.FromMinutes(2);

    public static RowButtons Buttons(CalendarApp app, Subscription subscription, bool changedSinceImport, DateTimeOffset now, TimeZoneInfo tz)
    {
        var action = CalendarApps.ActionFor(app);
        if (action.Kind == MainActionKind.Import)
        {
            if (subscription.ImportedAt is not { } at) return new(action.Label, null, null, "Importér igen…");
            var done = $"Importeret {Day(at, now, tz)}";
            return changedSinceImport ? new("Importér igen…", done, "Ændret siden import", "Importér igen…") : new(null, done, null, "Importér igen…");
        }
        // "✓ Tilføjet" først, når kalenderprogrammet har hentet filen fra AulaSync.
        if (subscription.Fetched) return new(null, "✓ Tilføjet", null, "Tilføj igen");
        if (subscription.AddedAt is not { } added) return new(action.Label, null, null, "Tilføj igen");
        // Ved "Kopiér adresse" ved AulaSync ikke, hvornår adressen bliver sat ind, så der er ingen frist.
        if (action.Kind == MainActionKind.CopyAddress) return new(null, "✓ Kopieret", null, "Tilføj igen");
        var program = app == CalendarApp.OutlookClassic ? "Outlook" : "Kalender";
        return now - added < FetchWait
            ? new(null, $"Venter på {program}…", null, "Tilføj igen", Waiting: true)
            : new(action.Label, null, $"{program} hentede ikke skemaet", "Tilføj igen");
    }

    // Hvornår rækken skifter af sig selv (fra "Venter på …" til "… hentede ikke skemaet"); null, hvis den ikke venter.
    public static DateTimeOffset? ChangesAt(CalendarApp app, Subscription subscription, DateTimeOffset now) =>
        CalendarApps.ActionFor(app).Kind == MainActionKind.Subscribe && !subscription.Fetched
        && subscription.AddedAt is { } added && now - added < FetchWait
            ? added + FetchWait
            : null;

    static string Day(DateTimeOffset at, DateTimeOffset now, TimeZoneInfo tz) =>
        TimeZoneInfo.ConvertTime(at, tz).Date == TimeZoneInfo.ConvertTime(now, tz).Date ? "i dag" : StatusText.ShortDate(at, tz);
}
