namespace AulaSync.Core;

public sealed class BackgroundScheduler(SyncService sync, TimeProvider time, FileLog log)
{
    public static readonly TimeSpan SyncInterval = TimeSpan.FromHours(6);
    public static readonly TimeSpan PingInterval = TimeSpan.FromMinutes(10);

    // Hvert 10. minut: sync hvis der er gået 6 timer, ellers keep-alive-ping.
    // Bruger vægur, så en Mac der har sovet, synkroniserer ved første tick efter opvågning.
    public async Task RunAsync(CancellationToken ct)
    {
        var nextSync = time.GetUtcNow();
        using var timer = new PeriodicTimer(PingInterval, time);
        try
        {
            do
            {
                try
                {
                    if (time.GetUtcNow() >= nextSync)
                    {
                        await sync.SyncAllAsync(ct);
                        nextSync = time.GetUtcNow() + SyncInterval;
                        sync.ReportNextSync(nextSync);
                    }
                    else
                    {
                        await sync.PingAsync(ct);
                    }
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                catch (Exception ex) { log.Error("Baggrundsopdatering fejlede", ex); }
            }
            while (await timer.WaitForNextTickAsync(ct));
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
    }
}
