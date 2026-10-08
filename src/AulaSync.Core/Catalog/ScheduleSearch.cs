using System.Globalization;

namespace AulaSync.Core;

public enum CatalogFilter { All, Employees, Groups, Resources }

public abstract record CatalogRow;
public sealed record CatalogHeader(string Title) : CatalogRow;
public sealed record CatalogEntry(ScheduleRef Schedule, bool Subscribed, bool Suggested) : CatalogRow;

public static class ScheduleSearch
{
    static readonly CompareInfo Danish = CultureInfo.GetCultureInfo("da-DK").CompareInfo;

    static readonly (ScheduleKind Kind, CatalogFilter Filter, string Title)[] Groups =
    [
        (ScheduleKind.Employee, CatalogFilter.Employees, "Medarbejdere"),
        (ScheduleKind.Group, CatalogFilter.Groups, "Klasser"),
        (ScheduleKind.Resource, CatalogFilter.Resources, "Lokaler"),
    ];

    // Alle ord i søgningen skal findes i kalendernavnet (som indeholder initialerne for medarbejdere).
    public static bool Matches(ScheduleRef schedule, string query) =>
        query.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .All(word => Danish.IndexOf(schedule.CalendarName, word, CompareOptions.IgnoreCase) >= 0);

    public static IReadOnlyList<CatalogRow> BuildRows(IReadOnlyList<ScheduleRef> all, string query, CatalogFilter filter,
        ScheduleRef? suggestion, IReadOnlySet<string> subscribedKeys)
    {
        query = query.Trim();
        var rows = new List<CatalogRow>();
        if (query == "" && filter is CatalogFilter.All or CatalogFilter.Employees
            && suggestion is { } own && !subscribedKeys.Contains(own.Key))
        {
            rows.Add(new CatalogHeader("Forslag"));
            rows.Add(new CatalogEntry(own, false, true));
        }

        foreach (var (kind, groupFilter, title) in Groups)
        {
            if (filter != CatalogFilter.All && filter != groupFilter) continue;
            var matches = all.Where(s => s.Kind == kind && Matches(s, query)).ToList();
            if (matches.Count == 0) continue;
            rows.Add(new CatalogHeader(title));
            rows.AddRange(matches.Select(s => new CatalogEntry(s, subscribedKeys.Contains(s.Key), false)));
        }
        return rows;
    }
}
