using System.Globalization;
using System.Text;

namespace AulaSync.Core;

// Refresh: hvor tit kalenderprogrammet bør hente filen igen (UpdateIntervals.Refresh). StatusZone: med en tidszone får
// kalenderen en privat begivenhed mandag morgen, der viser, hvornår skemaet blev hentet (Indstillinger); null: ingen.
public sealed record IcsOptions(TimeSpan Refresh, TimeZoneInfo? StatusZone = null)
{
    public static readonly IcsOptions Default = new(TimeSpan.FromHours(1));
}

public static class IcsWriter
{
    // Statusbegivenhedens UID; IcsInspect lader den ude af antal og hash.
    public const string StatusUidPrefix = "aulasync-status-";

    public static string Write(ScheduleRef schedule, IEnumerable<AulaEvent> events, DateTimeOffset now, IcsOptions? options = null)
    {
        options ??= IcsOptions.Default;
        var sb = new StringBuilder();
        void Line(string content) => sb.Append(IcsText.Fold(content));

        Line("BEGIN:VCALENDAR");
        Line("VERSION:2.0");
        Line("PRODID:-//AulaSync//Skema 3.0//DA");
        Line("CALSCALE:GREGORIAN");
        Line("METHOD:PUBLISH");
        Line($"X-WR-CALNAME:{IcsText.Escape(schedule.CalendarName)}");
        Line($"REFRESH-INTERVAL;VALUE=DURATION:{Duration(options.Refresh)}");
        Line($"X-PUBLISHED-TTL:{Duration(options.Refresh)}");

        var stamp = Utc(now);
        foreach (var e in events.OrderBy(e => e.Start).ThenBy(e => e.Id, StringComparer.Ordinal))
        {
            Line("BEGIN:VEVENT");
            // Starttiden er med, fordi gentagne begivenheder deler id i Aula (webappen bruger id + start som nøgle).
            Line($"UID:aula-{e.Id}-{Utc(e.Start)}-{schedule.Kind.Slug()}-{schedule.Id}@aulasync");
            Line($"DTSTAMP:{stamp}");
            Line($"DTSTART:{Utc(e.Start)}");
            if (e.End > e.Start) Line($"DTEND:{Utc(e.End)}"); // RFC 5545: DTEND skal ligge efter DTSTART
            Line($"SUMMARY:{IcsText.Escape(EventFormatter.Summary(schedule.Kind, e))}");
            if (e.Location != "") Line($"LOCATION:{IcsText.Escape(e.Location)}");
            var description = EventFormatter.Description(e);
            if (description != "") Line($"DESCRIPTION:{IcsText.Escape(description)}");
            Line("TRANSP:TRANSPARENT");
            Line("END:VEVENT");
        }

        if (options.StatusZone is { } zone)
        {
            // Mandag kl. 5.45-6.00 i denne uge (lørdag og søndag: den kommende), så den står samme sted hver uge. Samme UID
            // hver gang, så kalenderen flytter den i stedet for at lave en ny. Privat og uden at optage tid. Titel og
            // beskrivelse er det samme, fx "Opd. 091026@07:30".
            var start = StatusStart(now, zone);
            var text = IcsText.Escape(EventFormatter.StatusText(now, zone));
            Line("BEGIN:VEVENT");
            Line($"UID:{StatusUidPrefix}{schedule.Kind.Slug()}-{schedule.Id}@aulasync");
            Line($"DTSTAMP:{stamp}");
            Line($"DTSTART:{Utc(start)}");
            Line($"DTEND:{Utc(start.AddMinutes(15))}");
            Line($"SUMMARY:{text}");
            Line($"DESCRIPTION:{text}");
            Line("CLASS:PRIVATE");
            Line("TRANSP:TRANSPARENT");
            Line("END:VEVENT");
        }

        Line("END:VCALENDAR");
        return sb.ToString();
    }

    static DateTimeOffset StatusStart(DateTimeOffset now, TimeZoneInfo zone)
    {
        var day = DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(now, zone).DateTime);
        var monday = day.AddDays(-(((int)day.DayOfWeek + 6) % 7));
        if (day.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday) monday = monday.AddDays(7);
        var local = monday.ToDateTime(new TimeOnly(5, 45));
        while (zone.IsInvalidTime(local)) local = local.AddMinutes(30); // findes klokkeslættet ikke (sommertid), lidt senere
        return new DateTimeOffset(local, zone.GetUtcOffset(local));
    }

    // RFC 5545-varighed: PT30M, PT1H.
    static string Duration(TimeSpan t) =>
        t.Minutes == 0 ? $"PT{(int)t.TotalHours}H" : $"PT{(int)t.TotalMinutes}M";

    static string Utc(DateTimeOffset t) => t.UtcDateTime.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture);
}
