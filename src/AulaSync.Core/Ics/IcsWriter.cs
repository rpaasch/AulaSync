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

    // Hvornår skemaet blev hentet fra Aula (filens hoved); Rewrite og IcsInspect læser det, fordi DTSTAMP er tidspunktet for
    // den seneste skrivning.
    public const string FetchedProperty = "X-AULASYNC-FETCHED";

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
        Line($"{FetchedProperty}:{Utc(now)}");

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

        if (options.StatusZone is { } zone && !inDescription) StatusEvents(Line, schedule, stamp, status!, options, zone, now);

        Line("END:VCALENDAR");
        return sb.ToString();
    }

    // Hver mandag i skemaets periode (90 dage tilbage og frem), så den ses i hver uge: kl. 5.45-6.00 eller det klokkeslæt,
    // der er valgt, eller som heldagsbegivenhed. Én begivenhed pr. mandag med datoen i UID'en (ikke en gentagelse), så
    // klokkeslættet passer på begge sider af skift til og fra sommertid, og kalenderen opdaterer dem i stedet for at lave nye.
    // Privat og uden at optage tid. Titel og beskrivelse er det samme, fx "Opd. 091026@07:30".
    static void StatusEvents(Action<string> line, ScheduleRef schedule, string stamp, string status, IcsOptions options,
        TimeZoneInfo zone, DateTimeOffset now)
    {
        var text = IcsText.Escape(status);
        var time = options.StatusTime ?? AppConfig.DefaultStatusTime;
        foreach (var monday in StatusMondays(now, zone))
        {
            line("BEGIN:VEVENT");
            line($"UID:{StatusUidPrefix}{schedule.Kind.Slug()}-{schedule.Id}-{Date(monday)}@aulasync");
            line($"DTSTAMP:{stamp}");
            if (options.StatusStyle == StatusStyle.AllDay)
            {
                line($"DTSTART;VALUE=DATE:{Date(monday)}");
                line($"DTEND;VALUE=DATE:{Date(monday.AddDays(1))}");
            }
            else
            {
                var start = MondayAt(monday, time, zone);
                line($"DTSTART:{Utc(start)}");
                line($"DTEND:{Utc(start.AddMinutes(15))}");
            }
            line($"SUMMARY:{text}");
            line($"DESCRIPTION:{text}");
            line("CLASS:PRIVATE");
            line("TRANSP:TRANSPARENT");
            line("END:VEVENT");
        }
    }

    // Skriver opdateringstiden i en fil, AulaSync har skrevet, om uden at hente skemaet: lektionerne bliver, som de er, og
    // tidspunktet for hentningen (FetchedProperty; i en fil fra før 3.4 lektionernes DTSTAMP, ellers fetchedFallback) vises
    // på den måde, options siger, og med loggedOut som "⚠ Log ind i AulaSync · Opd. …". Bruges, når AulaSync er logget ud af
    // Aula, og når opdateringstiden igen skal være normal for et skema, der ikke kunne hentes. DTSTAMP bliver now, så
    // kalenderprogrammet ser ændringen. null: intet at ændre (kun DTSTAMP ville blive anderledes), eller ikke en fil fra
    // AulaSync.
    public static string? Rewrite(string ics, ScheduleRef schedule, DateTimeOffset now, IcsOptions options, bool loggedOut,
        DateTimeOffset fetchedFallback)
    {
        var lines = IcsInspect.Unfold(ics);
        if (lines.Length == 0 || lines[0] != "BEGIN:VCALENDAR" || !lines.Contains("END:VCALENDAR")) return null;
        var fetched = IcsInspect.Read(ics).Fetched ?? fetchedFallback;
        var hasFetched = lines.Any(l => l.StartsWith(FetchedProperty + ":", StringComparison.Ordinal));
        var stamp = Utc(now);
        var status = options.StatusZone is { } statusZone ? EventFormatter.StatusText(fetched, statusZone, loggedOut) : null;
        var inDescription = status is not null && options.StatusStyle == StatusStyle.InDescription;

        var sb = new StringBuilder();
        void Line(string content) => sb.Append(IcsText.Fold(content));
        var header = true;
        List<string>? lesson = null;
        var skip = false;
        foreach (var line in lines)
        {
            if (line == "") continue;
            if (header && (line == "BEGIN:VEVENT" || line == "END:VCALENDAR"))
            {
                header = false;
                if (!hasFetched) Line($"{FetchedProperty}:{Utc(fetched)}"); // en fil fra før 3.4
            }
            if (line == "BEGIN:VEVENT") { lesson = [line]; skip = false; continue; }
            if (lesson is not null)
            {
                if (line.StartsWith($"UID:{StatusUidPrefix}", StringComparison.Ordinal)) skip = true;
                if (line != "END:VEVENT") { lesson.Add(line); continue; }
                lesson.Add(line);
                if (!skip) foreach (var l in Lesson(lesson, stamp, inDescription ? status : null)) Line(l);
                lesson = null;
                continue;
            }
            if (line == "END:VCALENDAR" && status is not null && !inDescription)
                StatusEvents(Line, schedule, stamp, status, options, options.StatusZone!, now);
            Line(line);
        }
        var result = sb.ToString();
        return WithoutStamps(result) == WithoutStamps(ics) ? null : result;
    }

    // En lektion med nyt DTSTAMP og uden et tidligere tidspunkt sidst i beskrivelsen; status: tidspunktet sidst i
    // beskrivelsen (StatusStyle.InDescription), som Write skriver det.
    static IEnumerable<string> Lesson(List<string> lines, string stamp, string? status)
    {
        var hasDescription = false;
        foreach (var line in lines)
        {
            if (line.StartsWith("DTSTAMP:", StringComparison.Ordinal)) { yield return $"DTSTAMP:{stamp}"; continue; }
            if (line.StartsWith("DESCRIPTION:", StringComparison.Ordinal))
            {
                var description = IcsInspect.StatusInDescription().Match(line) is { Success: true } m ? m.Groups["rest"].Value : line[12..];
                hasDescription = true;
                if (status is not null) description = description == "" ? IcsText.Escape(status) : $"{description}\\n\\n{IcsText.Escape(status)}";
                if (description != "") yield return $"DESCRIPTION:{description}";
                continue;
            }
            // Uden beskrivelse står tidspunktet alene, før TRANSP, som i Write.
            if (line == "TRANSP:TRANSPARENT" && !hasDescription && status is not null) yield return $"DESCRIPTION:{IcsText.Escape(status)}";
            yield return line;
        }
    }

    static string WithoutStamps(string ics) =>
        string.Join("\n", IcsInspect.Unfold(ics).Where(l => !l.StartsWith("DTSTAMP:", StringComparison.Ordinal)));

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
