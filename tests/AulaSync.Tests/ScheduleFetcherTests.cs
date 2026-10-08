using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.Tests;

public class ScheduleFetcherTests
{
    static readonly Func<TimeSpan, CancellationToken, Task> NoDelay = (_, _) => Task.CompletedTask;
    static readonly ScheduleRef Anna = new(ScheduleKind.Employee, "1001", "Anna");

    [Fact]
    public void Chunks_cover_period_without_gaps_or_overlap()
    {
        var from = new DateOnly(2026, 1, 6);
        var to = new DateOnly(2026, 7, 5);
        var chunks = ScheduleFetcher.Chunks(from, to);
        Assert.Equal(from, chunks[0].From);
        Assert.Equal(to, chunks[^1].To);
        for (int i = 1; i < chunks.Count; i++) Assert.Equal(chunks[i - 1].To.AddDays(1), chunks[i].From);
        Assert.All(chunks, c => Assert.InRange(c.To.DayNumber - c.From.DayNumber + 1, 1, 42));
    }

    [Fact]
    public void Single_day_is_one_chunk() =>
        Assert.Single(ScheduleFetcher.Chunks(new DateOnly(2026, 4, 13), new DateOnly(2026, 4, 13)));

    [Fact]
    public void Exactly_84_days_is_two_chunks() =>
        Assert.Equal(2, ScheduleFetcher.Chunks(new DateOnly(2026, 1, 1), new DateOnly(2026, 1, 1).AddDays(83)).Count);

    [Fact]
    public async Task Fetch_uses_90_days_back_and_forward()
    {
        var calls = new List<(DateOnly, DateOnly)>();
        var fake = new FakeAulaClient { Events = (_, f, t) => { calls.Add((f, t)); return []; } };
        await new ScheduleFetcher(fake, NoDelay).FetchAsync(Anna, new DateOnly(2026, 4, 13), default);
        Assert.Equal(new DateOnly(2026, 4, 13).AddDays(-90), calls[0].Item1);
        Assert.Equal(new DateOnly(2026, 4, 13).AddDays(90), calls[^1].Item2);
    }

    [Fact]
    public async Task Events_repeated_across_chunks_are_deduplicated()
    {
        var fake = new FakeAulaClient { Events = (_, _, _) => [Ev(id: "7"), Ev(id: "8")] };
        var events = await new ScheduleFetcher(fake, NoDelay).FetchAsync(Anna, new DateOnly(2026, 4, 13), default);
        Assert.Equal(["7", "8"], events.Select(e => e.Id).Order());
        Assert.True(fake.EventCalls > 1);
    }

    // Gentagne begivenheder deler id i Aula: dubletter fjernes på id + starttid, så forekomsterne bevares.
    [Fact]
    public async Task Recurring_events_with_same_id_are_kept()
    {
        var fake = new FakeAulaClient
        {
            Events = (_, _, _) =>
            [
                Ev(id: "900", start: "2026-04-13T14:00:00+02:00", end: "2026-04-13T15:00:00+02:00"),
                Ev(id: "900", start: "2026-04-20T14:00:00+02:00", end: "2026-04-20T15:00:00+02:00"),
            ],
        };
        var events = await new ScheduleFetcher(fake, NoDelay).FetchAsync(Anna, new DateOnly(2026, 4, 13), default);
        Assert.Equal(2, events.Count);
        Assert.True(fake.EventCalls > 1); // samme forekomster fra flere bidder giver stadig kun to
    }

    [Fact]
    public async Task Failure_in_one_chunk_fails_the_whole_schedule()
    {
        int n = 0;
        var fake = new FakeAulaClient { Events = (_, _, _) => ++n == 3 ? throw new AulaException("boom") : [Ev()] };
        await Assert.ThrowsAsync<AulaException>(() => new ScheduleFetcher(fake, NoDelay).FetchAsync(Anna, new DateOnly(2026, 4, 13), default));
    }
}
