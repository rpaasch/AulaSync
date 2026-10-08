namespace AulaSync.Core;

public static class ScheduleOrdering
{
    public static IEnumerable<ScheduleRef> Sort(IEnumerable<ScheduleRef> items) => items
        .OrderBy(s => s.Kind)
        .ThenBy(s => s.Kind == ScheduleKind.Employee && s.Initials != "" ? s.Initials : s.Name, NaturalComparer.Danish)
        .ThenBy(s => s.Name, NaturalComparer.Danish)
        .ThenBy(s => s.Id, StringComparer.Ordinal);
}
