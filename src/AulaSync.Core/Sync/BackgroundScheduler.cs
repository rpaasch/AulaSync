namespace AulaSync.Core;

// Henter alle skemaer ved start og derefter med intervallet fra Indstillinger (standard hver 4. time), regnet fra starten
// af seneste opdatering, også en, brugeren har startet (login, Opdatér nu). Imellem holdes Aula-sessionen i live med et
// ping, når der er gået 10 minutter siden seneste forespørgsel. Tidspunkterne regnes på vægur, og der ventes højst 10
// minutter ad gangen, så en Mac der har sovet, opdaterer ved første tjek efter opvågning, og et ur, der stilles tilbage,
// ikke stopper ping.
public sealed class BackgroundScheduler
{
    public static readonly TimeSpan PingInterval = TimeSpan.FromMinutes(10);

    readonly SyncService _sync;
    readonly TimeProvider _time;
    readonly FileLog _log;
    readonly Func<TimeSpan> _interval;
    readonly Func<TimeSpan, CancellationToken, Task> _delay;
    readonly SemaphoreSlim _wake = new(0, 1);
    volatile bool _syncing; // planen opdaterer selv; dens egen start skal ikke vække den

    public BackgroundScheduler(SyncService sync, TimeProvider time, FileLog log, Func<TimeSpan> interval,
        Func<TimeSpan, CancellationToken, Task>? delay = null)
    {
        _sync = sync;
        _time = time;
        _log = log;
        _interval = interval;
        _delay = delay ?? ((wait, ct) => Task.Delay(wait, time, ct));
        sync.SyncStarted += () => { if (!_syncing) Reschedule(); };
    }

    // Intervallet er ændret i Indstillinger, eller en opdatering er startet udefra: næste opdatering regnes om, og den
    // sker med det samme, hvis den nu er forfalden.
    public void Reschedule()
    {
        try { _wake.Release(); }
        catch (SemaphoreFullException) { } // allerede vækket
    }

    public async Task RunAsync(CancellationToken ct)
    {
        var lastSync = DateTimeOffset.MinValue; // starten af planens seneste opdatering; MinValue: opdatér nu
        var nextPing = _time.GetUtcNow() + PingInterval;
        DateTimeOffset? reported = null;
        try
        {
            while (true)
            {
                var now = _time.GetUtcNow();
                // Er uret stillet tilbage, regnes fra nu, så der hverken går uger til næste opdatering eller næste ping.
                if (lastSync > now) lastSync = now;
                if (nextPing > now + PingInterval) nextPing = now + PingInterval;
                var interval = Interval();
                var asked = false;
                try
                {
                    if (now >= Latest(lastSync) + interval)
                    {
                        lastSync = now; // fra starten, så hentetiden ikke lægges oven i intervallet
                        asked = true;
                        _syncing = true;
                        try { await _sync.SyncAllAsync(ct); }
                        finally { _syncing = false; }
                    }
                    else if (now >= nextPing)
                    {
                        asked = true;
                        await _sync.PingAsync(ct);
                    }
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                catch (Exception ex) { _log.Error("Baggrundsopdatering fejlede", ex); }
                if (asked) nextPing = _time.GetUtcNow() + PingInterval;

                var due = Latest(lastSync) + interval;
                if (due != reported) _sync.ReportNextSync((reported = due).Value);
                var wait = (due < nextPing ? due : nextPing) - _time.GetUtcNow();
                if (wait > PingInterval) wait = PingInterval;
                if (wait > TimeSpan.Zero) await WaitAsync(wait, ct);
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { }
    }

    // Seneste start af en opdatering: planens egen eller en, brugeren har startet.
    DateTimeOffset Latest(DateTimeOffset lastSync) =>
        _sync.LastSyncStarted is { } started && started > lastSync ? started : lastSync;

    // Kan intervallet ikke læses (fx config.json låst et øjeblik), bruges standarden.
    TimeSpan Interval()
    {
        try { return _interval(); }
        catch (Exception ex)
        {
            _log.Error("Kunne ikke læse opdateringsintervallet", ex);
            return UpdateIntervals.Of(null);
        }
    }

    // Venter, til tiden er gået, eller planen bliver vækket.
    async Task WaitAsync(TimeSpan wait, CancellationToken ct)
    {
        using var cycle = CancellationTokenSource.CreateLinkedTokenSource(ct);
        var delay = _delay(wait, cycle.Token);
        var woken = _wake.WaitAsync(cycle.Token);
        await Task.WhenAny(delay, woken);
        cycle.Cancel(); // den anden venten stopper; et ubrugt væk bliver liggende til næste gang
        ct.ThrowIfCancellationRequested();
    }
}
