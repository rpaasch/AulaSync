namespace AulaSync.Core;

public sealed class ScheduleFetcher(IAulaClient client, Func<TimeSpan, CancellationToken, Task>? delay = null)
{
    public const int DaysBack = 90;
    public const int DaysForward = 90;
    public const int MaxChunkDays = 42;

    readonly Func<TimeSpan, CancellationToken, Task> _delay = delay ?? Task.Delay;

    // Dagintervaller (begge inklusive) uden overlap, hver på højst maxDays dage.
    public static IReadOnlyList<(DateOnly From, DateOnly To)> Chunks(DateOnly from, DateOnly to, int maxDays = MaxChunkDays)
    {
        var chunks = new List<(DateOnly, DateOnly)>();
        for (var start = from; start <= to; start = start.AddDays(maxDays))
        {
            var end = start.AddDays(maxDays - 1);
            chunks.Add((start, end > to ? to : end));
        }
        return chunks;
    }

    // Kaster videre, hvis blot én bid fejler — så overskrives en god fil aldrig med et halvt skema.
    public async Task<IReadOnlyList<AulaEvent>> FetchAsync(ScheduleRef schedule, DateOnly today, CancellationToken ct)
    {
        // Nøgle: id + start. Gentagne begivenheder deler id i Aula, så id alene ville smide forekomster væk.
        var unique = new Dictionary<(string Id, DateTimeOffset Start), AulaEvent>();
        var chunks = Chunks(today.AddDays(-DaysBack), today.AddDays(DaysForward));
        for (int i = 0; i < chunks.Count; i++)
        {
            if (i > 0) await _delay(TimeSpan.FromMilliseconds(250), ct);
            foreach (var e in await client.GetEventsAsync(schedule, chunks[i].From, chunks[i].To, ct))
                unique.TryAdd((e.Id, e.Start), e);
        }
        return unique.Values.ToList();
    }
}
