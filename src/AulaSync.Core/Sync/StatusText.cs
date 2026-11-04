namespace AulaSync.Core;

public enum StatusLevel { Ok, Warning, Error }

public enum BannerKind { LoggedOut, PortBusy, Offline }

public sealed record Banner(BannerKind Kind, string Text, string? ActionLabel);

public sealed record RowText(string Text, bool IsError);

public static class StatusText
{
    static readonly string[] Months = EventFormatter.Months;

    public static string Menu(SyncStatus s, bool serverRunning, TimeZoneInfo tz)
    {
        if (!s.LoggedIn) return "Logget ud af Aula";
        if (!serverRunning) return "Kalender-server kunne ikke starte";
        if (s.Connection == ConnectionState.Offline) return "Ingen forbindelse til Aula";
        if (s.Progress is { } p) return $"Henter {p.Done} af {p.Total}…";
        var count = s.Schedules == 1 ? "1 skema" : $"{s.Schedules} skemaer";
        return s.LastSuccess is { } at ? $"{count} · opdateret {Time(at, tz)}" : count;
    }

    // Logget ud af Aula: hvornår skemaerne sidst blev hentet, uden "næste" (der hentes ikke, før man logger ind).
    public static string BottomBar(SyncStatus s, DateTimeOffset now, TimeZoneInfo tz)
    {
        if (s.Progress is { } p) return $"Henter {p.Done} af {p.Total}…";
        if (!s.LoggedIn) return s.LastSuccess is { } last ? $"Ikke opdateret siden {When(last, now, tz)}" : "Logget ud af Aula";
        var next = s.NextSync is { } n ? $" · næste {Time(n, tz)}" : "";
        if (s.LastSuccess is not { } at) return "Ikke opdateret endnu" + next;
        return $"Opdateret {When(at, now, tz)}{next}";
    }

    // "i dag 07:30", "i går 07:30" eller "3. nov. 07:30".
    static string When(DateTimeOffset at, DateTimeOffset now, TimeZoneInfo tz)
    {
        var day = DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(at, tz).DateTime);
        var today = DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(now, tz).DateTime);
        var when = day == today ? "i dag" : day == today.AddDays(-1) ? "i går" : ShortDate(at, tz);
        return $"{when} {Time(at, tz)}";
    }

    public static StatusLevel Level(SyncStatus s, bool serverRunning) =>
        !s.LoggedIn || !serverRunning ? StatusLevel.Error
        : s.Connection == ConnectionState.Offline || s.LastError is not null ? StatusLevel.Warning
        : StatusLevel.Ok;

    // Ét banner ad gangen, vigtigste først.
    public static Banner? Banner(SyncStatus s, bool serverRunning, int port)
    {
        if (!s.LoggedIn) return new(BannerKind.LoggedOut, "Du er logget ud af Aula. Kalenderne viser stadig de seneste skemaer.", "Log ind igen");
        if (!serverRunning) return new(BannerKind.PortBusy, $"Kalender-server kunne ikke starte: port {port} er optaget af et andet program.", "Prøv igen");
        if (s.Connection == ConnectionState.Offline) return new(BannerKind.Offline, "Ingen forbindelse til Aula.", null);
        return null;
    }

    // loggedOut: logget ud af Aula; der hentes ikke og prøves ikke igen, så rækken siger, hvornår skemaet sidst blev hentet
    // (now: i dag eller i går), eller at det hentes efter et login.
    public static RowText Row(ScheduleRef schedule, ScheduleState state, DateTimeOffset? nextSync, TimeZoneInfo tz,
        bool loggedOut = false, DateTimeOffset now = default)
    {
        if (loggedOut)
            return state.UpdatedAt is { } fetched
                ? new($"{KindName(schedule)} · ikke opdateret siden {When(fetched, now, tz)}", false)
                : new($"{KindName(schedule)} · hentes, når du logger ind", false);
        if (state.Error is not null && state.ErrorAt is { } failed)
        {
            var retry = nextSync is { } n ? $" · prøver igen {Time(n, tz)}" : "";
            return new($"Kunne ikke hentes kl. {Time(failed, tz)}{retry}", true);
        }
        var kind = KindName(schedule);
        return state.Lessons is { } count
            ? new($"{kind} · {count} {(count == 1 ? "lektion" : "lektioner")}", false)
            : new($"{kind} · henter…", false);
    }

    // Medarbejdere med deres rolle fra Aula; klasser og lokaler med typen.
    public static string KindName(ScheduleRef schedule) =>
        schedule.Kind == ScheduleKind.Employee ? RoleName(schedule.Role) : KindName(schedule.Kind);

    public static string RoleName(string role) => role switch
    {
        "teacher" => "Lærer",
        "preschool-teacher" => "Pædagog",
        "leader" => "Leder",
        _ => "Medarbejder",
    };

    public static string KindName(ScheduleKind kind) => kind switch
    {
        ScheduleKind.Employee => "Medarbejder",
        ScheduleKind.Group => "Klasse",
        _ => "Lokale",
    };

    // Til beskeden og menuen, når importerede skemaer er ændret: navnet på ét skema, ellers antallet.
    public static string ImportChanged(IReadOnlyList<ScheduleRef> changed) =>
        changed.Count == 1 ? $"{changed[0].CalendarName} er ændret siden import" : $"{changed.Count} skemaer er ændret siden import";

    public static string ShortDate(DateTimeOffset at, TimeZoneInfo tz)
    {
        var local = TimeZoneInfo.ConvertTime(at, tz);
        return $"{local.Day}. {Months[local.Month - 1]}";
    }

    // InvariantCulture: dansk kultur kan bruge "." som tidsseparator.
    static string Time(DateTimeOffset at, TimeZoneInfo tz) => TimeZoneInfo.ConvertTime(at, tz).ToString("HH:mm", System.Globalization.CultureInfo.InvariantCulture);
}
