using System.Globalization;
using System.Text;

namespace AulaSync.Core;

// Refresh: hvor tit kalenderprogrammet bør hente filen igen (UpdateIntervals.Refresh). StatusZone: med en tidszone får
// kalenderen en privat begivenhed mandag morgen, der viser, hvornår skemaet blev hentet (Indstillinger); null: ingen.
// StatusTime: begivenhedens klokkeslæt hver mandag (null: kl. 5.45). StatusStyle: en aftale på klokkeslættet, en
// heldagsbegivenhed, eller tidspunktet nederst i hver lektions beskrivelse i stedet for en begivenhed.
public sealed record IcsOptions(TimeSpan Refresh, TimeZoneInfo? StatusZone = null, TimeOnly? StatusTime = null,
    StatusStyle StatusStyle = StatusStyle.Timed)
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
        var status = options.StatusZone is { } statusZone ? EventFormatter.StatusText(now, statusZone) : null;
        var inDescription = status is not null && options.StatusStyle == StatusStyle.InDescription;
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
            // Tidspunktet står sidst efter en tom linje; IcsInspect ser bort fra det (StatusInDescription).
            if (inDescription) description = description == "" ? status! : $"{description}\n\n{status}";
            if (description != "") Line($"DESCRIPTION:{IcsText.Escape(description)}");
            Line("TRANSP:TRANSPARENT");
            Line("END:VEVENT");
        }

        if (options.StatusZone is { } zone && !inDescription)
        {
            // Hver mandag i skemaets periode (90 dage tilbage og frem), så den ses i hver uge: kl. 5.45-6.00 eller det
            // klokkeslæt, der er valgt, eller som heldagsbegivenhed. Én begivenhed pr. mandag med datoen i UID'en (ikke en
            // gentagelse), så klokkeslættet passer på begge sider af skift til og fra sommertid, og kalenderen opdaterer dem
            // i stedet for at lave nye. Privat og uden at optage tid. Titel og beskrivelse er det samme, fx "Opd. 091026@07:30".
            var text = IcsText.Escape(status!);
            var time = options.StatusTime ?? AppConfig.DefaultStatusTime;
            foreach (var monday in StatusMondays(now, zone))
            {
                Line("BEGIN:VEVENT");
                Line($"UID:{StatusUidPrefix}{schedule.Kind.Slug()}-{schedule.Id}-{Date(monday)}@aulasync");
                Line($"DTSTAMP:{stamp}");
                if (options.StatusStyle == StatusStyle.AllDay)
                {
                    Line($"DTSTART;VALUE=DATE:{Date(monday)}");
                    Line($"DTEND;VALUE=DATE:{Date(monday.AddDays(1))}");
                }
                else
                {
                    var start = MondayAt(monday, time, zone);
                    Line($"DTSTART:{Utc(start)}");
                    Line($"DTEND:{Utc(start.AddMinutes(15))}");
                }
                Line($"SUMMARY:{text}");
                Line($"DESCRIPTION:{text}");
                Line("CLASS:PRIVATE");
                Line("TRANSP:TRANSPARENT");
                Line("END:VEVENT");
            }
        }

        Line("END:VCALENDAR");
        return sb.ToString();
    }

    // Mandagene i skemaets periode (ScheduleFetcher: 90 dage tilbage og frem fra i dag).
    static IEnumerable<DateOnly> StatusMondays(DateTimeOffset now, TimeZoneInfo zone)
    {
        var today = DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(now, zone).DateTime);
        var first = today.AddDays(-ScheduleFetcher.DaysBack);
        first = first.AddDays((8 - (int)first.DayOfWeek) % 7); // første mandag fra og med
        for (var monday = first; monday <= today.AddDays(ScheduleFetcher.DaysForward); monday = monday.AddDays(7)) yield return monday;
    }

    static DateTimeOffset MondayAt(DateOnly monday, TimeOnly time, TimeZoneInfo zone)
    {
        var local = monday.ToDateTime(time);
        while (zone.IsInvalidTime(local)) local = local.AddMinutes(30); // findes klokkeslættet ikke (sommertid), lidt senere
        return new DateTimeOffset(local, zone.GetUtcOffset(local));
    }

    // RFC 5545-varighed: PT30M, PT1H.
    static string Duration(TimeSpan t) =>
        t.Minutes == 0 ? $"PT{(int)t.TotalHours}H" : $"PT{(int)t.TotalMinutes}M";

    static string Utc(DateTimeOffset t) => t.UtcDateTime.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture);

    // Altid den gregorianske kalender: med fx thailandsk eller persisk format i Windows ville året ellers blive 2569 eller 1405.
    static string Date(DateOnly d) => d.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
}
