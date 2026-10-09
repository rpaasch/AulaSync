namespace AulaSync.Core;

// Ok: intet at sige. Missing: kalenderprogrammet henter dine andre skemaer, men ikke dette, så kalenderen er nok slettet.
// Quiet: skemaet er ikke hentet i lang tid, fx fordi kalenderprogrammet ikke har været åbent.
public enum FetchHealth { Ok, Missing, Quiet }

// Program: hvem der hentede skemaet sidst ("Outlook", "Kalender", "Thunderbird"; null: et andet program).
// ChangesAt: hvornår svaret kan skifte af sig selv, uden at der hentes noget.
public sealed record FetchCheck(FetchHealth Health, string? Program, DateTimeOffset? LastFetched, DateTimeOffset? ChangesAt)
{
    public static readonly FetchCheck None = new(FetchHealth.Ok, null, null, null);
}

// Holder øje med, hvornår kalenderprogrammerne henter skemaerne fra AulaSync (uden COM eller Graph). Har et program hentet
// uden ophold i mindst ActiveFor efter et skemas seneste hentning uden at hente det igen, er kalenderen nok slettet i
// programmet (Outlook henter hver kalender omtrent hver time, mens Outlook er åben). Et ophold er mere end Gap uden
// hentninger fra programmet, fx fordi det er lukket. Det huskes kun, mens AulaSync kører, og vurderes ved hver hentning
// (SyncService.MarkFetchedAsync), ikke kun når hovedvinduet viser rækkerne.
public sealed class FetchWatch(DateTimeOffset since)
{
    public static readonly TimeSpan Gap = TimeSpan.FromHours(3);
    public static readonly TimeSpan ActiveFor = TimeSpan.FromHours(4);
    public static readonly TimeSpan MissingAfter = TimeSpan.FromHours(6);
    // Apple Kalender kan sættes til at hente én gang om ugen pr. kalender.
    public static readonly TimeSpan MissingAfterCalendar = TimeSpan.FromDays(8);
    public static readonly TimeSpan QuietAfter = TimeSpan.FromDays(3);
    // Så længe skal AulaSync have kørt, før "ikke hentet i lang tid" siges (kalenderprogrammet kan lige være startet).
    public static readonly TimeSpan QuietWatch = TimeSpan.FromHours(2);

    readonly object _lock = new();
    readonly Dictionary<string, (DateTimeOffset Start, DateTimeOffset Seen)> _periods = new(); // program → hentninger uden ophold
    readonly Dictionary<string, DateTimeOffset> _last = new();   // filnavn → seneste hentning, også når den ikke er gemt
    readonly Dictionary<string, DateTimeOffset> _proven = new(); // filnavn → seneste hentning, da det viste sig at mangle

    // Hvornår AulaSync begyndte at holde øje (start).
    public DateTimeOffset Since { get; } = since;

    // Programmet ud fra User-Agent; null, hvis det er et andet program eller en browser.
    public static string? ProgramOf(string userAgent) =>
        userAgent.Contains("Outlook", StringComparison.OrdinalIgnoreCase) || userAgent.Contains("Microsoft Office", StringComparison.OrdinalIgnoreCase) ? "Outlook"
        : userAgent.Contains("CalendarAgent", StringComparison.OrdinalIgnoreCase) || userAgent.Contains("iCal/", StringComparison.OrdinalIgnoreCase) ? "Kalender"
        : userAgent.Contains("Thunderbird", StringComparison.OrdinalIgnoreCase) ? "Thunderbird"
        : null;

    // Programmet for det valgte kalenderprogram (til skemaer, hvor programmet ikke er gemt, fx fra 3.1).
    public static string? ProgramOf(CalendarApp? app) => app switch
    {
        CalendarApp.OutlookClassic => "Outlook",
        CalendarApp.AppleCalendar => "Kalender",
        _ => null,
    };

    public void Record(string fileName, string? program, DateTimeOffset at)
    {
        lock (_lock)
        {
            _last[fileName] = at;
            var key = program ?? "";
            _periods[key] = _periods.TryGetValue(key, out var p) && at - p.Seen <= Gap ? (p.Start, at) : (at, at);
        }
    }

    // Fra før 3.2 er kun første hentning gemt (FetchedAt); SyncService giver dem en seneste hentning ved første start.
    public DateTimeOffset? LastFetched(Subscription s)
    {
        DateTimeOffset? saved = s.LastFetchedAt ?? s.FetchedAt;
        lock (_lock)
            return _last.TryGetValue(s.Schedule.FileName, out var seen) && (saved is not { } t || seen > t) ? seen : saved;
    }

    public FetchCheck Check(Subscription s, DateTimeOffset now)
    {
        if (!s.Fetched || LastFetched(s) is not { } last) return FetchCheck.None;
        if (last > now) last = now; // uret har været stillet forkert
        var program = s.FetchedBy;
        var file = s.Schedule.FileName;
        bool proven;
        lock (_lock)
        {
            proven = _proven.TryGetValue(file, out var at) && at == last;
            if (!proven && _periods.TryGetValue(program ?? "", out var p) && p.Seen - Max(p.Start, last) >= ActiveFor)
            {
                _proven[file] = last;
                proven = true;
            }
        }
        var patience = program == "Kalender" ? MissingAfterCalendar : MissingAfter;
        var missingAt = last + patience;
        if (proven && now >= missingAt) return new(FetchHealth.Missing, program, last, null);
        var quietAt = Max(last + (patience > QuietAfter ? patience : QuietAfter), Since + QuietWatch);
        if (now >= quietAt) return new(FetchHealth.Quiet, program, last, null);
        return new(FetchHealth.Ok, program, last, proven ? Min(missingAt, quietAt) : quietAt);
    }

    static DateTimeOffset Max(DateTimeOffset a, DateTimeOffset b) => a > b ? a : b;
    static DateTimeOffset Min(DateTimeOffset a, DateTimeOffset b) => a < b ? a : b;
}
