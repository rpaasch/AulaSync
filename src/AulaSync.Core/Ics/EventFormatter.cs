namespace AulaSync.Core;

public static class EventFormatter
{
    internal static readonly string[] Months = ["jan.", "feb.", "mar.", "apr.", "maj", "jun.", "jul.", "aug.", "sep.", "okt.", "nov.", "dec."];
    static readonly string[] LongMonths = ["januar", "februar", "marts", "april", "maj", "juni", "juli", "august", "september", "oktober", "november", "december"];
    // DayOfWeek-rækkefølge: søndag først.
    static readonly string[] Weekdays = ["søn.", "man.", "tirs.", "ons.", "tors.", "fre.", "lør."];
    static readonly string[] LongWeekdays = ["søndag", "mandag", "tirsdag", "onsdag", "torsdag", "fredag", "lørdag"];

    public static string Summary(ScheduleKind kind, AulaEvent e)
    {
        var groups = string.Join(", ", e.Groups);
        var initials = string.Join(", ", e.Teachers.Select(t => t.Initials).Where(i => i != ""));
        var parts = new List<string>();
        if (e.IsSubstitute) parts.Add("Vikar");
        parts.Add(string.IsNullOrWhiteSpace(e.Title) ? "Uden titel" : e.Title);
        switch (kind)
        {
            case ScheduleKind.Employee: parts.Add(e.Location); parts.Add(groups); break;
            case ScheduleKind.Group: parts.Add(e.Location); parts.Add(initials); break;
            case ScheduleKind.Resource: parts.Add(groups); parts.Add(initials); break;
        }
        return string.Join(" | ", parts.Where(p => !string.IsNullOrWhiteSpace(p)));
    }

    public static string Description(AulaEvent e)
    {
        var lines = new List<string>();
        if (e.SubstituteFor is { } absent) lines.Add($"Vikar for: {Format(absent)}");
        if (e.Teachers.Count > 0) lines.Add(string.Join(", ", e.Teachers.Select(Format)));
        if (e.Groups.Count > 0) lines.Add($"Klasse: {string.Join(", ", e.Groups)}");
        if (e.Location != "") lines.Add($"Lokale: {e.Location}");
        return string.Join("\n", lines);
    }

    static string Format(Participant p) => p.Initials == "" ? p.Name : $"{p.Name} ({p.Initials})";

    // Statusbegivenheden (IcsWriter, Indstillinger › Vis opdateringstid i kalenderen): titel og beskrivelse.
    public static string StatusSummary(DateTimeOffset at, TimeZoneInfo tz)
    {
        var local = TimeZoneInfo.ConvertTime(at, tz);
        return $"AulaSync opdateret {Weekdays[(int)local.DayOfWeek]} {local.Day}. {Months[local.Month - 1]} {Time(local)}";
    }

    public static string StatusDescription(DateTimeOffset at, TimeZoneInfo tz)
    {
        var local = TimeZoneInfo.ConvertTime(at, tz);
        return $"AulaSync hentede skemaet fra Aula {LongWeekdays[(int)local.DayOfWeek]} {local.Day}. {LongMonths[local.Month - 1]} {local.Year} kl. {Time(local)}.\n"
            + "Denne private aftale kan slås fra i AulaSync under Indstillinger.";
    }

    // InvariantCulture: dansk kultur kan bruge "." som tidsseparator.
    static string Time(DateTimeOffset local) => local.ToString("HH:mm", System.Globalization.CultureInfo.InvariantCulture);
}
