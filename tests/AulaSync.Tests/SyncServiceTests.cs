using AulaSync.Core;
using Microsoft.Extensions.Time.Testing;
using static AulaSync.Tests.TestData;

namespace AulaSync.Tests;

public class SyncServiceTests : IDisposable
{
    static readonly Func<TimeSpan, CancellationToken, Task> NoDelay = (_, _) => Task.CompletedTask;
    static readonly ScheduleRef Anna = new(ScheduleKind.Employee, "1001", "Anna Eksempel", "AE");
    static readonly ScheduleRef SevenA = new(ScheduleKind.Group, "88231", "7A");
    static readonly DateTimeOffset Start = new(2026, 4, 13, 8, 0, 0, TimeSpan.Zero);

    readonly TempDir _dir = new();
    readonly FakeTimeProvider _time = new(Start);
    readonly SubscriptionStore _store;
    readonly SyncService _sync;
    readonly FakeAulaClient _client = new();

    public SyncServiceTests()
    {
        _store = new SubscriptionStore(_dir.File("abonnementer.json"));
        _sync = new SyncService(_store, _dir.File("kalendere"), new FileLog(_dir.File("log.txt")), _time, NoDelay);
    }

    public void Dispose() => _dir.Dispose();

    string IcsPath(ScheduleRef s) => _sync.FilePath(s);

    void Subscribe(params ScheduleRef[] items) => _store.Save(items.Select(s => new Subscription(s)));

    // Skemaer valgt før rollen kendtes (eller med en ændret rolle) får rollen fra Aula, når den kendes.
    [Fact]
    public async Task Roles_are_updated_on_existing_subscriptions()
    {
        Subscribe(Anna, SevenA);
        var changes = 0;
        _sync.StatusChanged += _ => changes++;

        await _sync.UpdateRolesAsync([Anna with { Role = "leader" }, new ScheduleRef(ScheduleKind.Employee, "1002", "Bo", "BT", "teacher"), Anna with { Role = "" }]);

        Assert.Equal(["leader", ""], _store.Load().Select(s => s.Schedule.Role));
        Assert.Equal(1, changes);
        var written = File.GetLastWriteTimeUtc(_dir.File("abonnementer.json"));
        await _sync.UpdateRolesAsync([Anna with { Role = "leader" }]);
        Assert.Equal(written, File.GetLastWriteTimeUtc(_dir.File("abonnementer.json")));
    }

    // Serveren har udleveret filen, så kalenderprogrammet har skemaet. Kun første hentning efter et klik gemmes,
    // så en kalender, der henter hvert 5. minut, ikke skriver abonnementer.json hver gang.
    [Fact]
    public async Task Fetch_after_click_is_recorded_once()
    {
        Subscribe(Anna, SevenA);
        await _sync.MarkAddedAsync(SevenA);
        _time.Advance(TimeSpan.FromSeconds(20));
        await _sync.MarkFetchedAsync(SevenA.FileName);
        await _sync.MarkFetchedAsync("klasse-999.ics"); // ukendt fil ændrer intet
        Assert.Equal([null, Start.AddSeconds(20)], _store.Load().Select(s => s.FetchedAt));

        var written = File.GetLastWriteTimeUtc(_dir.File("abonnementer.json"));
        _time.Advance(TimeSpan.FromMinutes(5));
        await _sync.MarkFetchedAsync(SevenA.FileName);
        Assert.Equal(written, File.GetLastWriteTimeUtc(_dir.File("abonnementer.json")));

        await _sync.MarkAddedAsync(SevenA); // "Tilføj igen"
        Assert.False(_store.Load()[1].Fetched);
        await _sync.MarkFetchedAsync(SevenA.FileName);
        Assert.True(_store.Load()[1].Fetched);
    }

    // Seneste hentning gemmes højst en gang i timen. Programmet gemmes, når det kendes; en browser overskriver det ikke, og
    // "Tilføj igen" glemmer det ikke.
    [Fact]
    public async Task Last_fetch_is_saved_at_most_hourly()
    {
        const string outlook = "Microsoft Office/16.0 (Windows NT 10.0; Microsoft Outlook 16.0; Pro)";
        const string browser = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/141.0 Safari/537.36";
        Subscribe(SevenA);
        await _sync.MarkFetchedAsync(SevenA.FileName, outlook);
        var s = _store.Load().Single();
        Assert.Equal((Start, Start, "Outlook"), (s.FetchedAt, s.LastFetchedAt, s.FetchedBy));

        _time.Advance(TimeSpan.FromMinutes(59));
        await _sync.MarkFetchedAsync(SevenA.FileName, "macOS/15.6 CalendarAgent/1000");
        Assert.Equal((Start, "Outlook"), (_store.Load().Single().LastFetchedAt, _store.Load().Single().FetchedBy));
        Assert.Equal(Start.AddMinutes(59), _sync.CheckFetch(_store.Load().Single()).LastFetched); // husket, ikke gemt

        _time.Advance(TimeSpan.FromMinutes(1));
        await _sync.MarkFetchedAsync(SevenA.FileName, browser);
        s = _store.Load().Single();
        Assert.Equal((Start, Start.AddHours(1), "Outlook"), (s.FetchedAt, s.LastFetchedAt, s.FetchedBy));

        _time.Advance(TimeSpan.FromHours(1));
        await _sync.MarkFetchedAsync(SevenA.FileName, "macOS/15.6 CalendarAgent/1000");
        Assert.Equal("Kalender", _store.Load().Single().FetchedBy);

        await _sync.MarkAddedAsync(SevenA);
        Assert.Equal("Kalender", _store.Load().Single().FetchedBy);
    }

    // At et skema ikke hentes længere, ses ved hentningerne, også når hovedvinduet ikke er åbent: Outlook henter kun 7A i
    // fem timer, lukkes og åbnes næste morgen; Annas skema mangler stadig.
    [Fact]
    public async Task Missing_schedule_is_noticed_without_the_main_window()
    {
        const string outlook = "Microsoft Office/16.0 (Windows NT 10.0; Microsoft Outlook 16.0; Pro)";
        Subscribe(Anna, SevenA);
        await _sync.MarkFetchedAsync(Anna.FileName, outlook);
        await _sync.MarkFetchedAsync(SevenA.FileName, outlook);
        for (var i = 0; i < 10; i++) // Annas kalender er slettet i Outlook
        {
            _time.Advance(TimeSpan.FromMinutes(30));
            await _sync.MarkFetchedAsync(SevenA.FileName, outlook);
        }
        _time.Advance(TimeSpan.FromHours(15)); // Outlook lukket om natten, åbnet igen om morgenen
        await _sync.MarkFetchedAsync(SevenA.FileName, outlook);
        Assert.Equal(FetchHealth.Missing, _sync.CheckFetch(_store.Load().Single(s => s.Key == Anna.Key)).Health);
        Assert.Equal(FetchHealth.Ok, _sync.CheckFetch(_store.Load().Single(s => s.Key == SevenA.Key)).Health);
    }

    // Skemaer fra 3.1 har kun første hentning: ved første start i 3.2 får de seneste hentning = nu, og uden gemt program
    // bruges det valgte kalenderprogram.
    [Fact]
    public void Schedules_from_3_1_get_a_last_fetch_and_the_chosen_program()
    {
        _store.Save([new Subscription(SevenA, FetchedAt: Start.AddDays(-20)), new Subscription(Anna)]);
        var config = new AppConfig(CalendarApp.OutlookClassic);
        var sync = new SyncService(_store, _dir.File("kalendere"), new FileLog(_dir.File("log.txt")), _time, NoDelay, () => config);
        Assert.Equal([Start, null], _store.Load().Select(s => s.LastFetchedAt));
        Assert.Equal("Outlook", sync.CheckFetch(_store.Load()[0]).Program);
    }

    [Fact]
    public async Task Without_client_nothing_happens()
    {
        Subscribe(Anna);
        await _sync.SyncAllAsync(default);
        Assert.False(File.Exists(IcsPath(Anna)));
        Assert.Equal(ConnectionState.LoggedOut, _sync.Status.Connection);
    }

    [Fact]
    public async Task Writes_one_file_per_subscription_and_records_lesson_counts()
    {
        Subscribe(Anna, SevenA);
        _client.Events = (s, _, _) => s.Kind == ScheduleKind.Group ? [Ev(id: "1"), Ev(id: "2")] : [Ev(id: "5001")];
        _sync.SetClient(_client);
        await _sync.SyncAllAsync(default);
        Assert.Contains("X-WR-CALNAME:AE Anna Eksempel", File.ReadAllText(IcsPath(Anna)));
        Assert.Equal(new ScheduleState(2, Start, null, null), _sync.StateOf(SevenA));
        Assert.Equal(new SyncStatus(ConnectionState.Online, 2, Start, null, null, null), _sync.Status);
    }

    [Fact]
    public void State_falls_back_to_file_on_disk()
    {
        Directory.CreateDirectory(_dir.File("kalendere"));
        File.WriteAllText(IcsPath(Anna), IcsWriter.Write(Anna, [Ev(id: "1"), Ev(id: "2"), Ev(id: "3")], Start));
        Assert.Equal(3, _sync.StateOf(Anna).Lessons);
        Assert.Null(_sync.StateOf(SevenA).Lessons);
    }

    [Fact]
    public async Task Failure_keeps_existing_file_and_marks_row()
    {
        Subscribe(Anna);
        _client.Events = (_, _, _) => [Ev(id: "5001")];
        _sync.SetClient(_client);
        await _sync.SyncAllAsync(default);
        var before = File.ReadAllText(IcsPath(Anna));

        _client.Events = (_, _, _) => throw new AulaException("Aula er nede");
        _time.Advance(TimeSpan.FromHours(6));
        await _sync.SyncAllAsync(default);

        Assert.Equal(before, File.ReadAllText(IcsPath(Anna)));
        Assert.Contains("Aula er nede", _sync.Status.LastError);
        Assert.Equal(Start, _sync.Status.LastSuccess);
        var state = _sync.StateOf(Anna);
        Assert.Equal(1, state.Lessons);
        Assert.Equal(Start.AddHours(6), state.ErrorAt);
        Assert.Contains("Aula er nede", state.Error);
    }

    [Fact]
    public async Task One_failing_schedule_does_not_stop_the_others()
    {
        Subscribe(Anna, SevenA);
        _client.Events = (s, _, _) => s.Kind == ScheduleKind.Employee ? throw new AulaException("x") : [Ev()];
        _sync.SetClient(_client);
        await _sync.SyncAllAsync(default);
        Assert.True(File.Exists(IcsPath(SevenA)));
    }

    // "Opdateret" følger de skemaer, der lykkedes: kun en kørsel, hvor alle skemaer fejler, lader tidspunktet stå.
    [Fact]
    public async Task Partial_failure_still_updates_last_success()
    {
        Subscribe(Anna, SevenA);
        _sync.SetClient(_client);
        await _sync.SyncAllAsync(default);

        _client.Events = (s, _, _) => s.Kind == ScheduleKind.Group ? throw new AulaException("Aula er nede") : [Ev()];
        _time.Advance(TimeSpan.FromHours(6));
        await _sync.SyncAllAsync(default);

        Assert.Equal(Start.AddHours(6), _sync.Status.LastSuccess);
        Assert.Contains("Aula er nede", _sync.Status.LastError);
        Assert.Equal(Start.AddHours(6), _sync.StateOf(SevenA).ErrorAt);
    }

    [Fact]
    public async Task Offline_after_some_schedules_still_updates_last_success()
    {
        Subscribe(Anna, SevenA);
        _client.Events = (s, _, _) => s.Kind == ScheduleKind.Group ? throw new HttpRequestException("netværk") : [Ev()];
        _sync.SetClient(_client);
        await _sync.SyncAllAsync(default);
        Assert.Equal(ConnectionState.Offline, _sync.Status.Connection);
        Assert.Equal(Start, _sync.Status.LastSuccess);
    }

    [Fact]
    public async Task Session_expiry_after_some_schedules_still_updates_last_success()
    {
        Subscribe(Anna, SevenA);
        _client.Events = (s, _, _) => s.Kind == ScheduleKind.Group ? throw new SessionExpiredException() : [Ev()];
        _sync.SetClient(_client);
        await _sync.SyncAllAsync(default);
        Assert.Equal(ConnectionState.LoggedOut, _sync.Status.Connection);
        Assert.Equal(Start, _sync.Status.LastSuccess);
    }

    // Ingen adgang til ét skema (fx ikke medlem af klassen) er en fejl på den række — ikke et udløbet login.
    [Fact]
    public async Task No_access_marks_the_row_and_the_others_continue()
    {
        Subscribe(Anna, SevenA);
        _client.Events = (s, _, _) => s.Kind == ScheduleKind.Employee ? throw new ForbiddenException() : [Ev()];
        _sync.SetClient(_client);
        int expired = 0;
        _sync.SessionExpired += () => expired++;

        await _sync.SyncAllAsync(default);

        Assert.Equal(0, expired);
        Assert.Equal(ConnectionState.Online, _sync.Status.Connection);
        Assert.True(File.Exists(IcsPath(SevenA)));
        Assert.Equal("Ingen adgang i Aula", _sync.StateOf(Anna).Error);
        Assert.Contains("Ingen adgang i Aula", _sync.Status.LastError);
    }

    [Fact]
    public async Task Network_error_means_offline_until_next_success()
    {
        Subscribe(Anna, SevenA);
        _client.Events = (_, _, _) => throw new HttpRequestException("netværk");
        _sync.SetClient(_client);
        await _sync.SyncAllAsync(default);
        Assert.Equal(ConnectionState.Offline, _sync.Status.Connection);
        Assert.Equal(1, _client.EventCalls); // stopper ved første netværksfejl

        _client.Events = (_, _, _) => [];
        await _sync.SyncAllAsync(default);
        Assert.Equal(ConnectionState.Online, _sync.Status.Connection);
        Assert.Null(_sync.Status.LastError);
    }

    [Fact]
    public async Task Progress_is_reported_and_cleared()
    {
        Subscribe(Anna, SevenA);
        _sync.SetClient(_client);
        var seen = new List<SyncProgress?>();
        _sync.StatusChanged += s => seen.Add(s.Progress);
        await _sync.SyncAllAsync(default);
        Assert.Contains(new SyncProgress(1, 2), seen);
        Assert.Contains(new SyncProgress(2, 2), seen);
        Assert.Null(_sync.Status.Progress);
    }

    [Fact]
    public async Task Ok_empty_result_writes_empty_calendar()
    {
        Subscribe(Anna);
        _sync.SetClient(_client);
        await _sync.SyncAllAsync(default);
        var ics = File.ReadAllText(IcsPath(Anna));
        Assert.Contains("BEGIN:VCALENDAR", ics);
        Assert.DoesNotContain("BEGIN:VEVENT", ics);
        Assert.Equal(0, _sync.StateOf(Anna).Lessons);
    }

    [Fact]
    public async Task Session_expiry_stops_sync_and_raises_event()
    {
        Subscribe(Anna, SevenA);
        _client.Events = (_, _, _) => throw new SessionExpiredException();
        _sync.SetClient(_client);
        int raised = 0;
        _sync.SessionExpired += () => raised++;

        await _sync.SyncAllAsync(default);

        Assert.Equal(1, raised);
        Assert.Equal(1, _client.EventCalls);
        Assert.Equal(ConnectionState.LoggedOut, _sync.Status.Connection);
        Assert.Null(_sync.Status.Progress);
        await _sync.SyncAllAsync(default);
        Assert.Equal(1, _client.EventCalls); // klienten er sluppet
    }

    [Fact]
    public async Task Ping_expiry_offline_and_recovery()
    {
        _sync.SetClient(_client);
        int raised = 0;
        _sync.SessionExpired += () => raised++;

        _client.OnPing = () => throw new HttpRequestException("netværk");
        await _sync.PingAsync(default);
        Assert.Equal(ConnectionState.Offline, _sync.Status.Connection);

        _client.OnPing = () => { };
        await _sync.PingAsync(default);
        Assert.Equal(ConnectionState.Online, _sync.Status.Connection);

        _client.OnPing = () => throw new SessionExpiredException();
        await _sync.PingAsync(default);
        Assert.Equal(1, raised);
        Assert.Equal(ConnectionState.LoggedOut, _sync.Status.Connection);
    }

    [Fact]
    public async Task Add_saves_and_writes_file_immediately()
    {
        _sync.SetClient(_client);
        await _sync.AddAsync(Anna, default);
        await _sync.AddAsync(Anna, default); // dublet ignoreres
        Assert.Equal([new Subscription(Anna)], _store.Load());
        Assert.True(File.Exists(IcsPath(Anna)));
        Assert.Equal(1, _sync.Status.Schedules);
    }

    // Ved første opdatering får filen institutionens nummer og initialer eller navn foran nøglen, og filen fra 3.0.0
    // (bare nøglen) slettes. Adressen er den samme, og en import er stadig uændret.
    [Fact]
    public async Task File_gets_readable_name_and_old_name_is_removed()
    {
        var dir = _dir.File("kalendere");
        Directory.CreateDirectory(dir);
        var old = Path.Combine(dir, "klasse-88231.ics");
        File.WriteAllText(old, IcsWriter.Write(SevenA, [Ahead("1")], Start));
        Subscribe(SevenA);
        Assert.Equal(old, _sync.FilePath(SevenA)); // før opdateringen: den gamle fil
        Assert.Equal(1, _sync.StateOf(SevenA).Lessons);
        await _sync.MarkImportedAsync(SevenA);

        _sync.SetClient(_client, "123456");
        _client.Events = (_, _, _) => [Ahead("1")];
        await _sync.SyncAllAsync(default);
        var renamed = Path.Combine(dir, "123456-7A-klasse-88231.ics");
        Assert.Equal([renamed], Directory.GetFiles(dir));
        Assert.Equal(renamed, _sync.FilePath(SevenA));
        Assert.False(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));
        Assert.Contains("Skrev 123456-7A-klasse-88231.ics (1 begivenheder)", File.ReadAllText(_dir.File("log.txt")));

        await _sync.RemoveAsync(SevenA, default);
        Assert.Empty(Directory.GetFiles(dir));
    }

    // Log ud midt i en opdatering: valgene ryddes, men kalenderfilerne bliver liggende (også dem, der ikke er hentet endnu).
    [Fact]
    public async Task Log_out_during_sync_keeps_the_files()
    {
        Subscribe(Anna, SevenA);
        _sync.SetClient(_client, "123456");
        _client.Events = (_, _, _) => [Ahead("1")];
        await _sync.SyncAllAsync(default);
        var dir = _dir.File("kalendere");
        string[] files = [Path.Combine(dir, "123456-7A-klasse-88231.ics"), Path.Combine(dir, "123456-AE-medarbejder-1001.ics")];
        Assert.Equal(files, Directory.GetFiles(dir).Order());

        var loggedOut = false;
        _client.Events = (_, _, _) =>
        {
            if (!loggedOut)
            {
                loggedOut = true;
                _sync.SetClient(null);
                _sync.ClearAllAsync().GetAwaiter().GetResult();
            }
            return [Ahead("1")];
        };
        await _sync.SyncAllAsync(default);
        Assert.Equal(files, Directory.GetFiles(dir).Order());
        Assert.Empty(_store.Load());
    }

    [Fact]
    public async Task Remove_deletes_subscription_state_and_file()
    {
        _sync.SetClient(_client);
        await _sync.AddAsync(Anna, default);
        await _sync.RemoveAsync(Anna with { Name = "Andet navn" }, default);
        Assert.Empty(_store.Load());
        Assert.False(File.Exists(IcsPath(Anna)));
        Assert.Null(_sync.StateOf(Anna).Lessons);
        Assert.Equal(0, _sync.Status.Schedules);
    }

    [Fact]
    public async Task Marks_are_saved()
    {
        _sync.SetClient(_client);
        _client.Events = (_, _, _) => [Ahead("1")];
        await _sync.AddAsync(Anna, default);
        await _sync.AddAsync(SevenA, default);
        _time.Advance(TimeSpan.FromMinutes(5));
        await _sync.MarkAddedAsync(Anna);
        await _sync.MarkImportedAsync(SevenA);
        var subs = _store.Load();
        Assert.Equal(Start.AddMinutes(5), subs.Single(s => s.Key == Anna.Key).AddedAt);
        var imported = subs.Single(s => s.Key == SevenA.Key);
        Assert.Equal(Start.AddMinutes(5), imported.ImportedAt);
        Assert.Equal(Start.AddMinutes(5).AddDays(90), imported.ImportUntil);
        Assert.Equal(IcsInspect.ReadFile(IcsPath(SevenA), imported.ImportedAt, imported.ImportUntil)!.ContentHash, imported.ImportHash);
    }

    // Indstillingerne (interval og statusbegivenhed) læses ved hver opdatering. Statusbegivenheden ændres hver gang, men
    // er ikke en ændring i skemaet: et importeret skema er stadig uændret.
    [Fact]
    public async Task Settings_reach_the_calendar_file_without_changing_imports()
    {
        var config = new AppConfig();
        _time.SetLocalTimeZone(Copenhagen);
        var sync = new SyncService(_store, _dir.File("kalendere"), new FileLog(_dir.File("log.txt")), _time, NoDelay, () => config);
        sync.SetClient(_client);
        _client.Events = (_, _, _) => [Ahead("1")];
        await sync.AddAsync(SevenA, default);
        await sync.MarkImportedAsync(SevenA);
        var ics = File.ReadAllText(sync.FilePath(SevenA));
        Assert.Contains("REFRESH-INTERVAL;VALUE=DURATION:PT1H\r\n", ics);
        Assert.DoesNotContain("aulasync-status", ics);

        config = new AppConfig(UpdateMinutes: 30, StatusEvent: true);
        _time.Advance(TimeSpan.FromHours(1));
        await sync.SyncAllAsync(default);
        ics = File.ReadAllText(sync.FilePath(SevenA));
        Assert.Contains("REFRESH-INTERVAL;VALUE=DURATION:PT30M\r\n", ics);
        Assert.Contains("UID:aulasync-status-klasse-88231@aulasync\r\n", ics);
        Assert.Contains("SUMMARY:Opd. 130426@11:00\r\n", ics); // 9:00 UTC er 11:00 i København
        Assert.Equal(1, sync.StateOf(SevenA).Lessons);
        Assert.False(sync.ChangedSinceImport(sync.Subscriptions.Single()));
    }

    // Skemaet dækker altid 90 dage tilbage og 90 dage frem fra i dag, så perioden rykker en dag hver dag. Det er ikke en
    // ændring: kun lektionerne fra importen og frem til den sidste importerede lektion tæller.
    [Fact]
    public async Task Moving_window_is_not_a_change_since_import()
    {
        _sync.SetClient(_client);
        _client.Events = (_, from, to) => Daily(from, to);
        await _sync.AddAsync(SevenA, default);
        await _sync.MarkImportedAsync(SevenA);
        var imported = _sync.Subscriptions.Single();
        Assert.Equal(Start.AddDays(90), imported.ImportUntil); // importen dækker de næste 90 dage
        Assert.Equal(IcsInspect.ReadFile(IcsPath(SevenA), Start, imported.ImportUntil)!.ContentHash, imported.ImportHash);

        _time.Advance(TimeSpan.FromDays(1)); // næste dag: den ældste dag falder ud, en ny kommer til
        await _sync.SyncAllAsync(default);
        Assert.False(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));
        Assert.Empty(_sync.ChangedImports());

        _client.Events = (_, from, to) => Daily(from, to, moved: new DateOnly(2026, 4, 20)); // en lektion flytter
        await _sync.SyncAllAsync(default);
        Assert.True(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));
    }

    // Efter 90 dage er hele det importerede tidsrum gået, og de første lektioner falder ud af skemaet: så siger AulaSync
    // til, for importen dækker ikke længere fremad.
    [Fact]
    public async Task Import_that_has_run_out_counts_as_changed()
    {
        _sync.SetClient(_client);
        _client.Events = (_, from, to) => Daily(from, to);
        await _sync.AddAsync(SevenA, default);
        await _sync.MarkImportedAsync(SevenA);
        _time.Advance(TimeSpan.FromDays(89));
        await _sync.SyncAllAsync(default);
        Assert.False(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));

        _time.Advance(TimeSpan.FromDays(2));
        await _sync.SyncAllAsync(default);
        Assert.True(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));
    }

    // Importeret i en ferie (ingen lektioner fremad): perioden, der rykker, er stadig ikke en ændring, men nye lektioner er.
    [Fact]
    public async Task Import_without_lessons_ahead_counts_new_lessons_but_not_the_moving_window()
    {
        _sync.SetClient(_client);
        _client.Events = (_, from, to) => Daily(from, to, last: new DateOnly(2026, 4, 10));
        await _sync.AddAsync(SevenA, default);
        await _sync.MarkImportedAsync(SevenA);
        _time.Advance(TimeSpan.FromDays(1));
        await _sync.SyncAllAsync(default);
        Assert.False(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));

        _client.Events = (_, from, to) => [.. Daily(from, to, last: new DateOnly(2026, 4, 10)), Ev(id: "ny", start: "2026-05-01T08:00:00Z", end: "2026-05-01T08:45:00Z")];
        await _sync.SyncAllAsync(default);
        Assert.True(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));
    }

    // Kommer der lektioner efter den sidste importerede (fx det nye skoleår), men inden for de 90 dage, importen dækker,
    // er skemaet ændret.
    [Fact]
    public async Task New_lessons_after_the_last_imported_one_count_as_changed()
    {
        _sync.SetClient(_client);
        _client.Events = (_, from, to) => Daily(from, to, last: new DateOnly(2026, 5, 1));
        await _sync.AddAsync(SevenA, default);
        await _sync.MarkImportedAsync(SevenA);
        _time.Advance(TimeSpan.FromDays(7));
        await _sync.SyncAllAsync(default);
        Assert.False(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));

        _client.Events = (_, from, to) => Daily(from, to, last: new DateOnly(2026, 6, 15));
        await _sync.SyncAllAsync(default);
        Assert.True(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));
    }

    // En lektion i morgen kl. 8 (dansk tid), altså efter importen; hour flytter den.
    static AulaEvent Ahead(string id, int hour = 8) =>
        Ev(id: id, start: $"2026-04-14T{hour:00}:00:00+02:00", end: $"2026-04-14T{hour:00}:45:00+02:00");

    // Ingen lektioner i de 90 dage, importen dækker (fx en orlov): når tidsrummet er gået, tæller importen som ændret.
    [Fact]
    public async Task Import_without_lessons_in_its_period_runs_out_when_the_period_has_passed()
    {
        _sync.SetClient(_client);
        _client.Events = (_, from, to) => Daily(from, to, first: new DateOnly(2026, 7, 21));
        await _sync.AddAsync(SevenA, default);
        await _sync.MarkImportedAsync(SevenA);
        _time.Advance(TimeSpan.FromDays(89));
        await _sync.SyncAllAsync(default);
        Assert.False(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));

        _time.Advance(TimeSpan.FromDays(1));
        await _sync.SyncAllAsync(default);
        Assert.True(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));
    }

    // Går tidsrummet, mens AulaSync er lukket, meldes udløbet ved første opdatering efter start. Udløbet tæller først, når
    // filen er skrevet efter tidsrummet; ellers ville det tælle som "ændret før start" og kun få prik og mærke.
    [Fact]
    public async Task Run_out_while_closed_is_reported_after_the_next_update()
    {
        _sync.SetClient(_client);
        _client.Events = (_, from, to) => Daily(from, to, first: new DateOnly(2026, 7, 21));
        await _sync.AddAsync(SevenA, default);
        await _sync.MarkImportedAsync(SevenA);
        _time.Advance(TimeSpan.FromDays(89));
        await _sync.SyncAllAsync(default);
        File.SetLastWriteTimeUtc(IcsPath(SevenA), _time.GetUtcNow().UtcDateTime); // filen blev sidst skrevet dag 89
        _time.Advance(TimeSpan.FromDays(2)); // AulaSync var lukket, da tidsrummet gik

        var restarted = new SyncService(_store, _dir.File("kalendere"), new FileLog(_dir.File("log.txt")), _time, NoDelay);
        var reported = new List<string>();
        restarted.ImportsChanged += changed => reported.Add(string.Join(",", changed.Select(s => s.CalendarName)));
        Assert.Empty(restarted.ChangedImports());
        restarted.SetClient(_client);
        await restarted.SyncAllAsync(default);
        Assert.Equal(["7A"], reported);
    }

    // De 90 dage lægges til på den lokale kalender, så tidsrummet ikke rækker ud over filens sidste dag, når sommertiden
    // begynder imellem (her: import kl. 23.50 dansk tid den 10. januar).
    [Fact]
    public async Task Import_period_ends_on_the_local_day_90_days_ahead()
    {
        var time = new FakeTimeProvider(new DateTimeOffset(2027, 1, 10, 22, 50, 0, TimeSpan.Zero));
        time.SetLocalTimeZone(Copenhagen);
        var sync = new SyncService(_store, _dir.File("kalendere"), new FileLog(_dir.File("log.txt")), time, NoDelay);
        Subscribe(SevenA);
        await sync.MarkImportedAsync(SevenA);
        Assert.Equal(new DateTimeOffset(2027, 4, 10, 23, 50, 0, TimeSpan.FromHours(2)), _store.Load().Single().ImportUntil);
    }

    // Én lektion kl. 8 (UTC) hver dag i det tidsrum, AulaSync beder om, fra first til og med last; på dagen moved kl. 10.
    static IReadOnlyList<AulaEvent> Daily(DateOnly from, DateOnly to, DateOnly? moved = null, DateOnly? last = null, DateOnly? first = null)
    {
        var list = new List<AulaEvent>();
        for (var d = first is { } f && f > from ? f : from; d <= (last is { } l && l < to ? l : to); d = d.AddDays(1))
        {
            var start = new DateTimeOffset(d.ToDateTime(new TimeOnly(d == moved ? 10 : 8, 0)), TimeSpan.Zero);
            list.Add(Ev(id: $"{d:yyyyMMdd}", start: start.ToString("o"), end: start.AddMinutes(45).ToString("o")));
        }
        return list;
    }

    // Et importeret skema, der ændrer sig ved en opdatering, meldes én gang; efter ny import meldes næste ændring igen.
    [Fact]
    public async Task Imports_that_change_are_reported_once()
    {
        var reported = new List<string>();
        _sync.ImportsChanged += changed => reported.Add(string.Join(",", changed.Select(s => s.CalendarName)));
        _sync.SetClient(_client);
        _client.Events = (_, _, _) => [Ahead("1")];
        await _sync.AddAsync(SevenA, default);
        await _sync.MarkImportedAsync(SevenA);
        await _sync.SyncAllAsync(default);
        Assert.Empty(reported);
        Assert.Empty(_sync.ChangedImports());

        _client.Events = (_, _, _) => [Ahead("1"), Ahead("2")];
        await _sync.SyncAllAsync(default);
        await _sync.SyncAllAsync(default); // stadig ændret: ingen ny melding
        Assert.Equal(["7A"], reported);
        Assert.Equal([SevenA], _sync.ChangedImports());

        await _sync.MarkImportedAsync(SevenA); // importeret igen
        Assert.Empty(_sync.ChangedImports());
        _client.Events = (_, _, _) => [Ahead("3")];
        await _sync.SyncAllAsync(default);
        Assert.Equal(["7A", "7A"], reported);
    }

    // Var skemaet allerede ændret, da AulaSync startede, kommer der ingen ny melding (men det står stadig som ændret).
    [Fact]
    public async Task Imports_changed_before_start_are_not_reported_again()
    {
        _sync.SetClient(_client);
        _client.Events = (_, _, _) => [Ahead("1")];
        await _sync.AddAsync(SevenA, default);
        await _sync.MarkImportedAsync(SevenA);
        _client.Events = (_, _, _) => [Ahead("1"), Ahead("2")];
        await _sync.SyncAllAsync(default);

        var restarted = new SyncService(_store, _dir.File("kalendere"), new FileLog(_dir.File("log.txt")), _time, NoDelay);
        var reported = 0;
        restarted.ImportsChanged += _ => reported++;
        restarted.SetClient(_client);
        await restarted.SyncAllAsync(default);
        Assert.Equal(0, reported);
        Assert.Equal([SevenA], restarted.ChangedImports());
    }

    [Fact]
    public async Task Changed_since_import_only_when_lessons_change()
    {
        _sync.SetClient(_client);
        _client.Events = (_, _, _) => [Ahead("1")];
        await _sync.AddAsync(SevenA, default);
        await _sync.MarkImportedAsync(SevenA);
        Assert.False(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));

        _time.Advance(TimeSpan.FromHours(6)); // nyt DTSTAMP, samme lektioner
        await _sync.SyncAllAsync(default);
        Assert.False(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));

        _client.Events = (_, _, _) => [Ahead("1", hour: 9)];
        await _sync.SyncAllAsync(default);
        Assert.True(_sync.ChangedSinceImport(_sync.Subscriptions.Single()));
    }

    [Fact]
    public async Task Clear_all_keeps_files()
    {
        _sync.SetClient(_client);
        await _sync.AddAsync(Anna, default);
        await _sync.ClearAllAsync();
        Assert.Empty(_store.Load());
        Assert.True(File.Exists(IcsPath(Anna)));
        Assert.Equal(0, _sync.Status.Schedules);
    }

    [Fact]
    public void Next_sync_and_status_events()
    {
        SyncStatus? seen = null;
        _sync.StatusChanged += s => seen = s;
        _sync.SetClient(_client);
        Assert.Equal(ConnectionState.Online, seen?.Connection);
        _sync.ReportNextSync(Start.AddHours(6));
        Assert.Equal(Start.AddHours(6), seen?.NextSync);
    }
}
