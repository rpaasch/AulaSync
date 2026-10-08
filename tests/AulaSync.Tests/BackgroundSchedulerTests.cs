using AulaSync.Core;
using Microsoft.Extensions.Time.Testing;
using static AulaSync.Tests.TestData;

namespace AulaSync.Tests;

public class BackgroundSchedulerTests : IDisposable
{
    static readonly DateTimeOffset Start = new(2026, 4, 13, 8, 0, 0, TimeSpan.Zero);

    readonly TempDir _dir = new();
    readonly FakeTimeProvider _time = new(Start);
    readonly FileLog _log;
    readonly SyncService _sync;
    readonly FakeAulaClient _client = new();
    readonly List<TimeSpan> _syncs = [];
    readonly List<TimeSpan> _pings = [];
    readonly int _chunksPerSync = ScheduleFetcher.Chunks(new DateOnly(2026, 1, 1), new DateOnly(2026, 1, 1).AddDays(180)).Count;

    public BackgroundSchedulerTests()
    {
        var store = new SubscriptionStore(_dir.File("abonnementer.json"));
        store.Save([new Subscription(new ScheduleRef(ScheduleKind.Resource, "412", "53"))]);
        _log = new FileLog(_dir.File("log.txt"));
        _sync = new SyncService(store, _dir.File("kalendere"), _log, _time, (_, _) => Task.CompletedTask);
        _client.OnPing = () => _pings.Add(_time.GetUtcNow() - Start);
        _sync.SetClient(_client);
    }

    public void Dispose() => _dir.Dispose();

    // Hentningen tager 3 minutter (uret flyttes ved første forespørgsel i hver opdatering).
    void SlowFetch() => _client.Events = (_, _, _) =>
    {
        if (_client.EventCalls % _chunksPerSync == 1)
        {
            _syncs.Add(_time.GetUtcNow() - Start);
            _time.Advance(TimeSpan.FromMinutes(3));
        }
        return [];
    };

    [Fact]
    public async Task Syncs_at_start_then_every_interval_and_pings_every_10_minutes_between()
    {
        SlowFetch();
        using var cts = new CancellationTokenSource();
        var waits = new List<TimeSpan>();
        // Ventetiden går med det samme; efter 8 timer og 5 minutter stoppes planen.
        Task Delay(TimeSpan wait, CancellationToken ct)
        {
            waits.Add(wait);
            _time.Advance(wait);
            if (_time.GetUtcNow() >= Start.AddHours(8).AddMinutes(5)) cts.Cancel();
            return Task.CompletedTask;
        }

        await new BackgroundScheduler(_sync, _time, _log, () => TimeSpan.FromHours(4), Delay).RunAsync(cts.Token); // stopper uden exception

        // Intervallet regnes fra starten af hver hentning, så de 3 minutter ikke lægges til.
        Assert.Equal([TimeSpan.Zero, TimeSpan.FromHours(4), TimeSpan.FromHours(8)], _syncs);
        Assert.Equal(3 * _chunksPerSync, _client.EventCalls);
        // Ping 10 minutter efter seneste forespørgsel, også efter en hentning; aldrig samtidig med en hentning.
        var expected = new List<TimeSpan>();
        foreach (var sync in _syncs.Take(2))
            for (var t = sync + TimeSpan.FromMinutes(13); t < sync + TimeSpan.FromHours(4); t += TimeSpan.FromMinutes(10)) expected.Add(t);
        Assert.Equal(expected, _pings);
        Assert.All(waits, w => Assert.InRange(w, TimeSpan.Zero, BackgroundScheduler.PingInterval)); // højst 10 min ad gangen (Mac i dvale)
        Assert.Equal(Start.AddHours(12), _sync.Status.NextSync);
    }

    [Fact]
    public async Task Changed_interval_takes_effect_at_once()
    {
        var interval = TimeSpan.FromHours(4);
        var waiting = 0;
        using var cts = new CancellationTokenSource();
        // Uret står stille; planen venter, til den bliver vækket.
        Task Delay(TimeSpan wait, CancellationToken ct)
        {
            Interlocked.Increment(ref waiting);
            return Task.Delay(Timeout.Infinite, ct);
        }
        var scheduler = new BackgroundScheduler(_sync, _time, _log, () => interval, Delay);
        var run = scheduler.RunAsync(cts.Token);
        await WaitUntil(() => waiting == 1);
        Assert.Equal(_chunksPerSync, _client.EventCalls);
        Assert.Equal(Start.AddHours(4), _sync.Status.NextSync);

        // Kortere end tiden siden sidst: hent nu.
        _time.Advance(TimeSpan.FromHours(1));
        interval = TimeSpan.FromMinutes(30);
        scheduler.Reschedule();
        await WaitUntil(() => _client.EventCalls == 2 * _chunksPerSync);
        await WaitUntil(() => _sync.Status.NextSync == Start.AddHours(1).AddMinutes(30));

        // Længere: kun et nyt tidspunkt.
        interval = TimeSpan.FromHours(8);
        scheduler.Reschedule();
        await WaitUntil(() => _sync.Status.NextSync == Start.AddHours(9));
        Assert.Equal(2 * _chunksPerSync, _client.EventCalls);
        Assert.Equal(0, _client.PingCalls);

        cts.Cancel();
        await run;
    }

    // Login og Opdatér nu henter også alle skemaer; næste planlagte opdatering regnes fra dem, så Aula ikke spørges igen
    // få minutter efter.
    [Fact]
    public async Task Sync_from_outside_moves_the_next_scheduled_sync()
    {
        var waiting = 0;
        using var cts = new CancellationTokenSource();
        Task Delay(TimeSpan wait, CancellationToken ct)
        {
            Interlocked.Increment(ref waiting);
            return Task.Delay(Timeout.Infinite, ct);
        }
        var run = new BackgroundScheduler(_sync, _time, _log, () => TimeSpan.FromHours(4), Delay).RunAsync(cts.Token);
        await WaitUntil(() => waiting == 1);

        _time.Advance(TimeSpan.FromMinutes(235));
        await _sync.SyncAllAsync(default); // fx Opdatér nu kl. 3:55 efter start
        await WaitUntil(() => _sync.Status.NextSync == Start.AddMinutes(235).AddHours(4));
        Assert.Equal(2 * _chunksPerSync, _client.EventCalls);

        cts.Cancel();
        await run;
    }

    // Kan config.json ikke læses et øjeblik (fx låst af antivirus), bruges standarden, og planen kører videre.
    [Fact]
    public async Task Unreadable_interval_falls_back_to_the_default()
    {
        SlowFetch();
        var reads = 0;
        TimeSpan Interval() => ++reads is 2 or 3 ? throw new IOException("Testfejl") : TimeSpan.FromHours(4);
        using var cts = new CancellationTokenSource();
        Task Delay(TimeSpan wait, CancellationToken ct)
        {
            _time.Advance(wait);
            if (_time.GetUtcNow() >= Start.AddHours(8).AddMinutes(5)) cts.Cancel();
            return Task.CompletedTask;
        }

        await new BackgroundScheduler(_sync, _time, _log, Interval, Delay).RunAsync(cts.Token);

        Assert.Equal([TimeSpan.Zero, TimeSpan.FromHours(4), TimeSpan.FromHours(8)], _syncs);
        Assert.Contains("Testfejl", File.ReadAllText(_dir.File("log.txt")));
    }

    // Stilles uret tilbage (fx af tidssynkronisering), ventes der stadig højst 10 minutter, og ping fortsætter.
    [Fact]
    public async Task Clock_set_back_does_not_stop_the_pings()
    {
        var clock = new ManualClock(Start);
        var store = new SubscriptionStore(_dir.File("abonnementer.json"));
        var sync = new SyncService(store, _dir.File("kalendere"), _log, clock, (_, _) => Task.CompletedTask);
        var client = new FakeAulaClient();
        sync.SetClient(client);
        var waits = new List<TimeSpan>();
        using var cts = new CancellationTokenSource();
        Task Delay(TimeSpan wait, CancellationToken ct)
        {
            waits.Add(wait);
            clock.Now += wait;
            if (waits.Count == 1) clock.Now -= TimeSpan.FromDays(60);
            if (waits.Count == 20) cts.Cancel();
            return Task.CompletedTask;
        }

        await new BackgroundScheduler(sync, clock, _log, () => TimeSpan.FromHours(4), Delay).RunAsync(cts.Token);

        Assert.Equal(20, waits.Count);
        Assert.All(waits, w => Assert.InRange(w, TimeSpan.Zero, BackgroundScheduler.PingInterval));
        Assert.True(client.PingCalls >= 18, $"{client.PingCalls} ping");
    }

    sealed class ManualClock(DateTimeOffset start) : TimeProvider
    {
        public DateTimeOffset Now { get; set; } = start;
        public override DateTimeOffset GetUtcNow() => Now;
        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }

    [Theory]
    [InlineData(null, 240)]
    [InlineData(30, 30)]
    [InlineData(60, 60)]
    [InlineData(120, 120)]
    [InlineData(480, 480)]
    [InlineData(45, 240)] // ukendt værdi (rettet i config.json): standard
    [InlineData(0, 240)]
    public void Interval_from_config(int? minutes, int expected) =>
        Assert.Equal(TimeSpan.FromMinutes(expected), new AppConfig(UpdateMinutes: minutes).UpdateInterval);

    [Fact]
    public void Interval_labels() =>
        Assert.Equal(["Hver halve time", "Hver time", "Hver 2. time", "Hver 4. time (standard)", "Hver 8. time"],
            UpdateIntervals.Minutes.Select(UpdateIntervals.Label));

    // Kalenderprogrammet henter filen fra AulaSync mindst hver time; det spørger ikke Aula.
    [Theory]
    [InlineData(30, 30)]
    [InlineData(60, 60)]
    [InlineData(240, 60)]
    public void Calendar_refresh_is_at_most_an_hour(int interval, int expected) =>
        Assert.Equal(TimeSpan.FromMinutes(expected), UpdateIntervals.Refresh(TimeSpan.FromMinutes(interval)));
}
