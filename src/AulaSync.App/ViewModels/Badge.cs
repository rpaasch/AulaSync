using AulaSync.Core;

namespace AulaSync.App;

// Typemærket foran hvert skema: initialer, klassenavn eller lokalenummer (højst 3 tegn), farvet efter type.
public static class Badge
{
    public static string TextFor(ScheduleRef schedule)
    {
        var text = schedule.Kind switch
        {
            ScheduleKind.Employee => schedule.Initials != "" ? schedule.Initials : InitialsOf(schedule.Name),
            ScheduleKind.Resource => FirstNumber(schedule.Name) ?? schedule.Name,
            _ => schedule.Name.Replace(" ", ""),
        };
        text = text.Trim();
        return text.Length <= 3 ? text : text[..3];
    }

    public static string ClassFor(ScheduleKind kind) => kind switch
    {
        ScheduleKind.Employee => "staff",
        ScheduleKind.Group => "class",
        _ => "room",
    };

    static string InitialsOf(string name)
    {
        var words = name.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return words.Length == 0 ? "?" : string.Concat(words[0][0], words.Length > 1 ? words[^1][0].ToString() : "").ToUpperInvariant();
    }

    static string? FirstNumber(string name)
    {
        var digits = new string(name.SkipWhile(c => !char.IsAsciiDigit(c)).TakeWhile(char.IsAsciiDigit).ToArray());
        return digits == "" ? null : digits;
    }
}
