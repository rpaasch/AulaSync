using AulaSync.Core;
using Microsoft.Extensions.Time.Testing;
using static AulaSync.Tests.TestData;

namespace AulaSync.Tests;

public class BackgroundSchedulerTests
{
    [Fact]
    public async Task Syncs_at_start_pings_every_10_minutes_and_syncs_every_6_hours()
    {
        using var dir = new TempDir();
        var start = new DateTimeOffset(2026, 4, 13, 8, 0, 0, TimeSpan.Zero);
        var time = new FakeTimeProvider(start);
        var store = new SubscriptionStore(dir.File("abonnementer.json"));
        store.Save([new Subscription(new ScheduleRef(ScheduleKind.Resource, "412", "53"))]);
        var log = new FileLog(dir.File("log.txt"));
        var sync = new SyncService(store, dir.File("kalendere"), log, time, (_, _) => Task.CompletedTask);
        var client = new FakeAulaClient();
        sync.SetClient(client);
        using var cts = new CancellationTokenSource();

        var run = new BackgroundScheduler(sync, time, log).RunAsync(cts.Token);
        var chunksPerSync = ScheduleFetcher.Chunks(new DateOnly(2026, 1, 1), new DateOnly(2026, 1, 1).AddDays(180)).Count;

        await WaitUntil(() => client.EventCalls == chunksPerSync);
        await WaitUntil(() => sync.Status.NextSync == start.AddHours(6));
        for (int i = 1; i <= 35; i++)
        {
            time.Advance(BackgroundScheduler.PingInterval);
            await WaitUntil(() => client.PingCalls == i);
        }
        Assert.Equal(chunksPerSync, client.EventCalls);

        time.Advance(BackgroundScheduler.PingInterval); // 6 timer
        await WaitUntil(() => client.EventCalls == 2 * chunksPerSync);
        Assert.Equal(35, client.PingCalls);

        cts.Cancel();
        await run; // stopper uden exception
    }
}
