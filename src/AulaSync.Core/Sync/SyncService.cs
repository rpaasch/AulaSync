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
    readonly Func<AppConfig> _config;
    readonly SemaphoreSlim _syncGate = new(1, 1);   // én hentning ad gangen
    readonly SemaphoreSlim _storeGate = new(1, 1);  // læs-ændr-skriv af abonnementer.json
    readonly object _lock = new();
    readonly Dictionary<string, ScheduleState> _states = new();
    readonly HashSet<string> _reportedImports; // importerede skemaer, der er meldt som ændret (eller var det ved start)
    readonly FetchWatch _fetches;
    IAulaClient? _client;
    string _institution = ""; // institutionens nummer i filnavnene (CalendarFiles)
    DateTimeOffset? _lastSyncStarted;

    // config: interval og statusbegivenhed (Indstillinger) til kalenderfilerne; læses ved hver opdatering.
    public SyncService(SubscriptionStore store, string calendarDir, FileLog log, TimeProvider time,
        Func<TimeSpan, CancellationToken, Task>? delay = null, Func<AppConfig>? config = null)
    {
        _store = store;
        _calendarDir = calendarDir;
        _log = log;
        _time = time;
        _delay = delay;
        _config = config ?? (() => new AppConfig());
        Status = SyncStatus.Initial(store.Load().Count);
        _fetches = new FetchWatch(time.GetUtcNow());
        StampLastFetched();
        _reportedImports = ChangedImports().Select(s => s.Key).ToHashSet();
    }

    public SyncStatus Status { get; private set; }
    public event Action<SyncStatus>? StatusChanged;
    public event Action? SessionExpired;
    // Efter en opdatering: importerede skemaer, der lige er blevet ændret siden importen (hvert skema én gang pr. import).
    public event Action<IReadOnlyList<ScheduleRef>>? ImportsChanged;
    // En opdatering af alle skemaer er startet med forbindelse til Aula (baggrundsplanen regner næste opdatering derfra).
    public event Action? SyncStarted;

    public DateTimeOffset? LastSyncStarted { get { lock (_lock) return _lastSyncStarted; } }

    public IReadOnlyList<Subscription> Subscriptions => _store.Load();

    // Skemaets fil: den, der ligger på disken (navnet kan være fra en tidligere udgave), ellers det navn, den får.
    public string FilePath(ScheduleRef schedule) => CalendarFiles.Find(_calendarDir, schedule.Key).FirstOrDefault() ?? NewPath(schedule);

    string NewPath(ScheduleRef schedule) => Path.Combine(_calendarDir, CalendarFiles.Name(Volatile.Read(ref _institution), schedule));

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

    // institution: den indloggede brugers institution (Profile.InstitutionCode); den står forrest i filnavnene. Ved log
    // ud beholdes den, så en opdatering, der stadig kører, ikke giver filerne et andet navn.
    public void SetClient(IAulaClient? client, string institution = "")
    {
        if (client is not null) Volatile.Write(ref _institution, institution);
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
            lock (_lock) _lastSyncStarted = _time.GetUtcNow();
            SyncStarted?.Invoke();
            var subscriptions = _store.Load();
            string? error = null;
            var offline = false;
            var anyOk = false; // "Opdateret" følger de skemaer, der lykkedes; kun hvis alle fejler, står tidspunktet
            for (int i = 0; i < subscriptions.Count; i++)
            {
                if (Volatile.Read(ref _client) != client) return; // logget ud undervejs: valgene er ryddet, filerne bliver
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

    // Seneste hentning gemmes højst så tit (se MarkFetchedAsync); FetchWatch husker den præcise tid, mens AulaSync kører.
    public static readonly TimeSpan SaveFetchEvery = TimeSpan.FromHours(1);

    // Serveren har udleveret filen (200 eller 304), så kalenderprogrammet har skemaet. Første hentning efter et klik gemmes
    // straks; derefter gemmes seneste hentning højst en gang i timen, så en kalender, der henter hvert 5. minut, ikke skriver
    // abonnementer.json hver gang. Programmet gemmes, når det kendes; en browser eller et ukendt program overskriver det ikke.
    public async Task MarkFetchedAsync(string fileName, string userAgent = "")
    {
        var now = _time.GetUtcNow();
        var program = FetchWatch.ProgramOf(userAgent);
        _fetches.Record(fileName, program, now);
        bool Save(Subscription s) => s.Schedule.FileName == fileName
            && (!s.Fetched || s.LastFetchedAt is not { } last || now - last >= SaveFetchEvery || last > now || (s.FetchedBy is null && program is not null));
        await ChangeSubscriptionsAsync(list => !list.Any(Save)
            ? list
            : list.Select(s => Save(s) ? s with { FetchedAt = s.Fetched ? s.FetchedAt : now, LastFetchedAt = now, FetchedBy = program ?? s.FetchedBy } : s).ToList());
        // Hver hentning kan vise, at et andet skema ikke hentes længere (FetchWatch husker det).
        foreach (var s in Subscriptions) _fetches.Check(WithProgram(s), now);
    }

    // Skemaer fra før 3.2 har kun første hentning (FetchedAt). Første gang får de seneste hentning = nu, så "ikke hentet i
    // 3 dage" regnes fra opgraderingen og ikke rykker ved hver start.
    void StampLastFetched()
    {
        var list = _store.Load();
        if (!list.Any(s => s.Fetched && s.LastFetchedAt is null)) return;
        var now = _time.GetUtcNow();
        try { _store.Save(list.Select(s => s.Fetched && s.LastFetchedAt is null ? s with { LastFetchedAt = now } : s).ToList()); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { _log.Error("Kunne ikke gemme seneste hentning", ex); }
    }

    // Henter kalenderprogrammet stadig skemaet (FetchWatch)? Uden kendt program (fx fra 3.1) bruges det valgte
    // kalenderprogram.
    public FetchCheck CheckFetch(Subscription subscription) => _fetches.Check(WithProgram(subscription), _time.GetUtcNow());

    Subscription WithProgram(Subscription s) =>
        s.FetchedBy is not null ? s : s with { FetchedBy = FetchWatch.ProgramOf(_config().CalendarApp) };

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
            var config = _config();
            var options = new IcsOptions(UpdateIntervals.Refresh(config.UpdateInterval), config.StatusEvent ? _time.LocalTimeZone : null);
            var path = NewPath(schedule);
            AtomicFile.WriteAllText(path, IcsWriter.Write(schedule, events, now, options));
            // Blev skemaet fjernet, mens det blev hentet, så fjern også filen igen. Ved log ud bliver filerne liggende.
            if (!_store.Load().Any(s => s.Key == schedule.Key))
            {
                if (Volatile.Read(ref _client) == client) DeleteFile(schedule);
                return (Outcome.Ok, null);
            }
            DeleteFile(schedule, except: path); // filen under et tidligere navn
            SetState(schedule, new ScheduleState(events.Count, now, null, null));
            _log.Info($"Skrev {Path.GetFileName(path)} ({events.Count} begivenheder)");
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

    // Alle skemaets filer, undtagen except (CalendarFiles.SameName: den nye fil må ikke slettes, fordi filsystemet
    // giver navnet tilbage med en anden stavemåde).
    void DeleteFile(ScheduleRef schedule, string? except = null)
    {
        foreach (var file in CalendarFiles.Find(_calendarDir, schedule.Key))
        {
            if (CalendarFiles.SameName(file, except)) continue;
            try { File.Delete(file); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { _log.Error($"Kunne ikke slette {file}", ex); }
        }
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
