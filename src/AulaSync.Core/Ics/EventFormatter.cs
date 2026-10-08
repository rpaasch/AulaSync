namespace AulaSync.Core;

public static class EventFormatter
{
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
}
