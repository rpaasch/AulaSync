namespace AulaSync.Core;

// Henter medarbejdere, klasser og lokaler én gang pr. login og cacher dem. Ved fejl prøves igen næste gang.
// loaded får listen, når den er hentet (så rollerne på valgte skemaer kan følge Aula).
public sealed class ScheduleCatalog(IAulaDirectory directory, Action<IReadOnlyList<ScheduleRef>>? loaded = null)
{
    readonly object _lock = new();
    Task<IReadOnlyList<ScheduleRef>>? _all;

    public Task<IReadOnlyList<ScheduleRef>> GetAllAsync()
    {
        lock (_lock)
        {
            // En fejl, der kastes allerede før første await, når ikke at rydde _all i catch — derfor tjekkes også IsFaulted.
            if (_all is null || _all.IsFaulted || _all.IsCanceled) _all = LoadAsync();
            return _all;
        }
    }

    async Task<IReadOnlyList<ScheduleRef>> LoadAsync()
    {
        try
        {
            // Ingen CancellationToken: en lukket dialog må ikke afbryde en hentning, som næste dialog kan genbruge.
            var employees = await directory.GetEmployeesAsync(CancellationToken.None);
            var groups = await directory.GetGroupsAsync(CancellationToken.None);
            var resources = await directory.GetResourcesAsync(CancellationToken.None);
            var all = employees.Select(e => new ScheduleRef(ScheduleKind.Employee, e.Id, e.Name, e.Initials, e.Role))
                .Concat(groups.Select(g => new ScheduleRef(ScheduleKind.Group, g.Id, g.Name)))
                .Concat(resources.Select(r => new ScheduleRef(ScheduleKind.Resource, r.Id, r.Name)));
            IReadOnlyList<ScheduleRef> sorted = ScheduleOrdering.Sort(all).ToList();
            loaded?.Invoke(sorted);
            return sorted;
        }
        catch
        {
            lock (_lock) _all = null;
            throw;
        }
    }
}
