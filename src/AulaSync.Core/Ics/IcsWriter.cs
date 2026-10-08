using System.Globalization;
using System.Text;

namespace AulaSync.Core;

public static class IcsWriter
{
    public static string Write(ScheduleRef schedule, IEnumerable<AulaEvent> events, DateTimeOffset now)
    {
        var sb = new StringBuilder();
        void Line(string content) => sb.Append(IcsText.Fold(content));

        Line("BEGIN:VCALENDAR");
        Line("VERSION:2.0");
        Line("PRODID:-//AulaSync//Skema 3.0//DA");
        Line("CALSCALE:GREGORIAN");
        Line("METHOD:PUBLISH");
        Line($"X-WR-CALNAME:{IcsText.Escape(schedule.CalendarName)}");
        Line("REFRESH-INTERVAL;VALUE=DURATION:PT6H");
        Line("X-PUBLISHED-TTL:PT6H");

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

        Line("END:VCALENDAR");
        return sb.ToString();
    }

    static string Utc(DateTimeOffset t) => t.UtcDateTime.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture);
}
