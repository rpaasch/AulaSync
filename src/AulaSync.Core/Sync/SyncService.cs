namespace AulaSync.Core;

public enum ConnectionState { LoggedOut, Online, Offline }

public sealed record SyncProgress(int Done, int Total);

public sealed record SyncStatus(ConnectionState Connection, int Schedules, DateTimeOffset? LastSuccess,
    DateTimeOffset? NextSync, string? LastError, SyncProgress? Progress)
{
    public bool LoggedIn => Connection != ConnectionState.LoggedOut;

    public static SyncStatus Initial(int schedules) => new(ConnectionState.LoggedOut, schedules, null, null, null, null);
}

public sealed record ScheduleState(int? Lessons, DateTimeOffset? UpdatedAt, string? Error, DateTimeOffset? ErrorAt);

public sealed class SyncService
{
    enum Outcome { Ok, Failed, Offline, SessionExpired }

    const string OfflineMessage = "Ingen forbindelse til Aula";

    readonly SubscriptionStore _store;
    readonly string _calendarDir;
    readonly FileLog _log;
    readonly TimeProvider _time;
    readonly Func<TimeSpan, CancellationToken, Task>? _delay;
    readonly SemaphoreSlim _syncGate = new(1, 1);   // én hentning ad gangen
    readonly SemaphoreSlim _storeGate = new(1, 1);  // læs-ændr-skriv af abonnementer.json
    readonly object _lock = new();
    readonly Dictionary<string, ScheduleState> _states = new();
    readonly HashSet<string> _reportedImports; // importerede skemaer, der er meldt som ændret (eller var det ved start)
    IAulaClient? _client;

    public SyncService(SubscriptionStore store, string calendarDir, FileLog log, TimeProvider time,
        Func<TimeSpan, CancellationToken, Task>? delay = null)
    {
        _store = store;
        _calendarDir = calendarDir;
        _log = log;
        _time = time;
        _delay = delay;
        Status = SyncStatus.Initial(store.Load().Count);
        _reportedImports = ChangedImports().Select(s => s.Key).ToHashSet();
    }

    public SyncStatus Status { get; private set; }
    public event Action<SyncStatus>? StatusChanged;
    public event Action? SessionExpired;
    // Efter en opdatering: importerede skemaer, der lige er blevet ændret siden importen (hvert skema én gang pr. import).
    public event Action<IReadOnlyList<ScheduleRef>>? ImportsChanged;

    public IReadOnlyList<Subscription> Subscriptions => _store.Load();

    public string FilePath(ScheduleRef schedule) => Path.Combine(_calendarDir, schedule.FileName);

    // Kendt tilstand fra denne kørsel; ellers læses filen på disken (fx lige efter opstart).
    public ScheduleState StateOf(ScheduleRef schedule)
    {
        lock (_lock)
            if (_states.TryGetValue(schedule.Key, out var known)) return known;
        var path = FilePath(schedule);
        return IcsInspect.ReadFile(path) is { } summary
            ? new ScheduleState(summary.Events, new DateTimeOffset(File.GetLastWriteTimeUtc(path), TimeSpan.Zero), null, null)
            : new ScheduleState(null, null, null, null);
    }

    // Skemaet dækker 90 dage tilbage og frem fra i dag, så perioden rykker hver dag. Derfor sammenlignes kun lektionerne i
    // det tidsrum, importen dækker (fra importen og 90 dage frem). Er filen opdateret, efter tidsrummet er gået, tæller
    // importen også som ændret: den dækker så ikke længere fremad. Filens opdatering bruges (ikke uret), så udløbet kommer
    // ved en opdatering og bliver meldt, også hvis AulaSync var lukket, da tidsrummet gik. Uden ImportUntil (gemt før):
    // hele filen.
    public bool ChangedSinceImport(Subscription subscription) =>
        subscription.ImportHash is { } hash
        && IcsInspect.ReadFile(FilePath(subscription.Schedule),
            subscription.ImportUntil is null ? null : subscription.ImportedAt, subscription.ImportUntil) is { } now
        && (now.ContentHash != hash || subscription.ImportUntil <= StateOf(subscription.Schedule).UpdatedAt);

    // Importerede skemaer, hvis indhold er ændret siden importen.
    public IReadOnlyList<ScheduleRef> ChangedImports() => _store.Load().Where(ChangedSinceImport).Select(s => s.Schedule).ToList();

    public void SetClient(IAulaClient? client)
    {
        Volatile.Write(ref _client, client);
        Update(s => s with { Connection = client is null ? ConnectionState.LoggedOut : ConnectionState.Online, LastError = null });
    }

    public void ReportNextSync(DateTimeOffset at) => Update(s => s with { NextSync = at });

    public async Task SyncAllAsync(CancellationToken ct)
    {
        await _syncGate.WaitAsync(ct);
        try
        {
            var client = Volatile.Read(ref _client);
            if (client is null) return;
            var subscriptions = _store.Load();
            string? error = null;
            var offline = false;
            var anyOk = false; // "Opdateret" følger de skemaer, der lykkedes; kun hvis alle fejler, står tidspunktet
            for (int i = 0; i < subscriptions.Count; i++)
            {
                var progress = new SyncProgress(i + 1, subscriptions.Count);
                Update(s => s with { Progress = progress });
                var (outcome, message) = await SyncOneAsync(client, subscriptions[i].Schedule, ct);
                if (outcome == Outcome.SessionExpired)
                {
                    if (anyOk) { var at = _time.GetUtcNow(); Update(s => s with { LastSuccess = at }); }
                    ReportChangedImports();
                    HandleExpired(client);
                    return;
                }
                if (outcome == Outcome.Offline) { offline = true; error = message; break; }
                anyOk |= outcome == Outcome.Ok;
                error ??= message;
            }
            var now = _time.GetUtcNow();
            Update(s => s with
            {
                Schedules = subscriptions.Count,
                Connection = !s.LoggedIn ? s.Connection : offline ? ConnectionState.Offline : ConnectionState.Online,
                LastSuccess = anyOk || error is null ? now : s.LastSuccess,
                LastError = error,
            });
            ReportChangedImports();
        }
        finally
        {
            Update(s => s.Progress is null ? s : s with { Progress = null });
            _syncGate.Release();
        }
    }

    public async Task AddAsync(ScheduleRef schedule, CancellationToken ct)
    {
        await ChangeSubscriptionsAsync(list =>
        {
            if (list.Any(s => s.Key == schedule.Key)) return list;
            _log.Info($"Skema tilføjet: {schedule.FileName} ({schedule.CalendarName})");
            return [.. list, new Subscription(schedule)];
        });

        await _syncGate.WaitAsync(ct);
        try
        {
            var client = Volatile.Read(ref _client);
            if (client is null) return;
            var (outcome, message) = await SyncOneAsync(client, schedule, ct);
            if (outcome == Outcome.SessionExpired) HandleExpired(client);
            else if (outcome == Outcome.Offline) Update(s => s with { Connection = ConnectionState.Offline, LastError = message });
        }
        finally { _syncGate.Release(); }
    }

    public async Task RemoveAsync(ScheduleRef schedule, CancellationToken ct)
    {
        await ChangeSubscriptionsAsync(list => list.Where(s => s.Key != schedule.Key).ToList());
        lock (_lock) _states.Remove(schedule.Key);
        DeleteFile(schedule);
        _log.Info($"Skema fjernet: {schedule.FileName}");
        Update(s => s);
    }

    public Task MarkAddedAsync(ScheduleRef schedule)
    {
        var now = _time.GetUtcNow();
        return ChangeSubscriptionsAsync(list => list.Select(s => s.Key == schedule.Key ? s with { AddedAt = now } : s).ToList());
    }

    public Task MarkImportedAsync(ScheduleRef schedule)
    {
        lock (_lock) _reportedImports.Remove(schedule.Key); // næste ændring meldes igen
        var now = _time.GetUtcNow();
        // Importen dækker lektionerne fra nu og de næste 90 dage, også dage uden lektioner endnu (se ChangedSinceImport).
        // Dagene lægges til på den lokale kalender, så tidsrummet ender på filens sidste dag, også over skift til sommertid.
        var ahead = _time.GetLocalNow().DateTime.AddDays(ScheduleFetcher.DaysForward);
        var until = new DateTimeOffset(ahead, _time.LocalTimeZone.GetUtcOffset(ahead)).ToUniversalTime();
        var hash = IcsInspect.ReadFile(FilePath(schedule), now, until)?.ContentHash;
        return ChangeSubscriptionsAsync(list => list.Select(s => s.Key == schedule.Key
            ? s with { AddedAt = now, ImportedAt = now, ImportHash = hash, ImportUntil = until } : s).ToList());
    }

    // Serveren har udleveret filen (200 eller 304), så kalenderprogrammet har skemaet. Kun første hentning efter et klik
    // gemmes, så en kalender, der henter hvert 5. minut, ikke skriver abonnementer.json hver gang.
    public Task MarkFetchedAsync(string fileName)
    {
        var now = _time.GetUtcNow();
        bool Waiting(Subscription s) => s.Schedule.FileName == fileName && !s.Fetched;
        return ChangeSubscriptionsAsync(list => !list.Any(Waiting)
            ? list
            : list.Select(s => Waiting(s) ? s with { FetchedAt = now } : s).ToList());
    }

    // Rollen (Lærer, Pædagog, Leder) på valgte medarbejderskemaer følger Aula: et skema valgt uden rolle, eller hvis rolle
    // er ændret, får den kendte rolle. En ukendt rolle ("") ændrer intet, og uden ændringer skrives filen ikke.
    public Task UpdateRolesAsync(IEnumerable<ScheduleRef> known)
    {
        var roles = new Dictionary<string, string>();
        foreach (var s in known)
            if (s.Kind == ScheduleKind.Employee && s.Role != "") roles.TryAdd(s.Key, s.Role);
        bool Differs(Subscription s, out string role) => roles.TryGetValue(s.Key, out role!) && role != s.Schedule.Role;
        return ChangeSubscriptionsAsync(list => !list.Any(s => Differs(s, out _))
            ? list
            : list.Select(s => Differs(s, out var role) ? s with { Schedule = s.Schedule with { Role = role } } : s).ToList());
    }

    // Log ud: valgene ryddes, .ics-filerne bliver liggende.
    public async Task ClearAllAsync()
    {
        await ChangeSubscriptionsAsync(_ => []);
        lock (_lock) _states.Clear();
        Update(s => s with { LastSuccess = null, LastError = null });
    }

    public async Task PingAsync(CancellationToken ct)
    {
        var client = Volatile.Read(ref _client);
        if (client is null) return;
        try
        {
            await client.PingAsync(ct);
            Update(s => s.Connection == ConnectionState.Offline ? s with { Connection = ConnectionState.Online, LastError = null } : s);
        }
        catch (SessionExpiredException) { HandleExpired(client); }
        catch (Exception ex) when (IsOffline(ex, ct))
        {
            Update(s => s.LoggedIn ? s with { Connection = ConnectionState.Offline } : s);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (Exception ex) { _log.Error("Keep-alive mod Aula fejlede", ex); }
    }

    async Task ChangeSubscriptionsAsync(Func<IReadOnlyList<Subscription>, IReadOnlyList<Subscription>> change)
    {
        await _storeGate.WaitAsync();
        try
        {
            var before = _store.Load();
            var after = change(before);
            if (!ReferenceEquals(before, after)) _store.Save(after);
            Update(s => s with { Schedules = after.Count });
        }
        finally { _storeGate.Release(); }
    }

    async Task<(Outcome, string?)> SyncOneAsync(IAulaClient client, ScheduleRef schedule, CancellationToken ct)
    {
        var now = _time.GetUtcNow();
        try
        {
            var today = DateOnly.FromDateTime(_time.GetLocalNow().DateTime);
            var events = await new ScheduleFetcher(client, _delay).FetchAsync(schedule, today, ct);
            AtomicFile.WriteAllText(FilePath(schedule), IcsWriter.Write(schedule, events, now));
            // Blev skemaet fjernet, mens det blev hentet, så fjern også filen igen.
            if (!_store.Load().Any(s => s.Key == schedule.Key)) { DeleteFile(schedule); return (Outcome.Ok, null); }
            SetState(schedule, new ScheduleState(events.Count, now, null, null));
            _log.Info($"Skrev {schedule.FileName} ({events.Count} begivenheder)");
            return (Outcome.Ok, null);
        }
        catch (SessionExpiredException) { return (Outcome.SessionExpired, null); }
        catch (Exception ex) when (IsOffline(ex, ct))
        {
            _log.Info($"Ingen forbindelse til Aula ved {schedule.FileName}: {ex.Message}");
            SetState(schedule, StateOf(schedule) with { Error = OfflineMessage, ErrorAt = now });
            return (Outcome.Offline, OfflineMessage);
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch (Exception ex)
        {
            _log.Error($"Opdatering af {schedule.FileName} fejlede — den gamle fil beholdes", ex);
            SetState(schedule, StateOf(schedule) with { Error = ex.Message, ErrorAt = now });
            return (Outcome.Failed, $"{schedule.CalendarName}: {ex.Message}");
        }
    }

    // Netværksfejl (ingen HTTP-status) og timeouts betyder "ingen forbindelse" — ikke en fejl i skemaet.
    static bool IsOffline(Exception ex, CancellationToken ct) =>
        ex is HttpRequestException { StatusCode: null } || (ex is TaskCanceledException && !ct.IsCancellationRequested);

    void DeleteFile(ScheduleRef schedule)
    {
        var file = FilePath(schedule);
        try { File.Delete(file); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { _log.Error($"Kunne ikke slette {file}", ex); }
    }

    void ReportChangedImports()
    {
        var changed = ChangedImports();
        List<ScheduleRef> fresh;
        lock (_lock)
        {
            fresh = changed.Where(s => !_reportedImports.Contains(s.Key)).ToList();
            _reportedImports.Clear();
            _reportedImports.UnionWith(changed.Select(s => s.Key));
        }
        if (fresh.Count > 0) ImportsChanged?.Invoke(fresh);
    }

    void SetState(ScheduleRef schedule, ScheduleState state)
    {
        lock (_lock) _states[schedule.Key] = state;
    }

    void HandleExpired(IAulaClient client)
    {
        if (Interlocked.CompareExchange(ref _client, null, client) != client) return;
        _log.Info("Aula-sessionen er udløbet");
        Update(s => s with { Connection = ConnectionState.LoggedOut });
        SessionExpired?.Invoke();
    }

    void Update(Func<SyncStatus, SyncStatus> change)
    {
        SyncStatus next;
        lock (_lock) next = Status = change(Status);
        StatusChanged?.Invoke(next);
    }
}
