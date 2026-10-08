namespace AulaSync.Core;

public enum ScheduleKind { Employee, Group, Resource }

public static class ScheduleKindExtensions
{
    public static string Slug(this ScheduleKind kind) => kind switch
    {
        ScheduleKind.Employee => "medarbejder",
        ScheduleKind.Group => "klasse",
        ScheduleKind.Resource => "lokale",
        _ => throw new ArgumentOutOfRangeException(nameof(kind)),
    };
}

// Role: Aulas institutionsrolle for en medarbejder (teacher, preschool-teacher, leader, other); "" hvis ukendt.
public sealed record ScheduleRef(ScheduleKind Kind, string Id, string Name, string Initials = "", string Role = "")
{
    public string Id { get; } = IsValidId(Id) ? Id : throw new ArgumentException($"Ugyldigt Aula-id: '{Id}'", nameof(Id));

    // Navnet i kalender-adressen (http://localhost:9876/klasse-88231.ics). Filen på disken har også institution og
    // initialer eller navn i navnet (CalendarFiles).
    public string FileName => $"{Kind.Slug()}-{Id}.ics";

    public string CalendarName => Kind == ScheduleKind.Employee && Initials != "" ? $"{Initials} {Name}" : Name;

    public string Key => $"{Kind.Slug()}-{Id}";

    // Kun ASCII-bogstaver, cifre og bindestreg: id'et indgår i filnavne og URL'er.
    public static bool IsValidId(string id) =>
        id.Length is > 0 and <= 40 && id.All(c => char.IsAsciiLetterOrDigit(c) || c == '-');
}

public sealed record Participant(string Name, string Initials);

public sealed record AulaEvent(
    string Id,
    string Title,
    DateTimeOffset Start,
    DateTimeOffset End,
    string Location,
    IReadOnlyList<string> Groups,
    IReadOnlyList<Participant> Teachers,
    bool IsSubstitute,
    Participant? SubstituteFor);

public sealed record Profile(string Id, string Name, string Initials, string Role, string InstitutionCode, string InstitutionName);

public sealed record Employee(string Id, string Name, string Initials, string Role);

public sealed record NamedItem(string Id, string Name);
