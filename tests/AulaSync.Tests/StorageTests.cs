using AulaSync.Core;

namespace AulaSync.Tests;

public class AtomicFileTests
{
    [Fact]
    public void Writes_utf8_without_bom_and_leaves_no_temp_files()
    {
        using var dir = new TempDir();
        var path = dir.File("a.ics");
        AtomicFile.WriteAllText(path, "æøå");
        Assert.Equal([0xC3, 0xA6], File.ReadAllBytes(path)[..2]);
        Assert.Single(Directory.GetFiles(dir.Path));
    }

    [Fact]
    public void Overwrites_existing_file()
    {
        using var dir = new TempDir();
        var path = dir.File("a.ics");
        AtomicFile.WriteAllText(path, "gammel");
        AtomicFile.WriteAllText(path, "ny");
        Assert.Equal("ny", File.ReadAllText(path));
    }

    [Fact]
    public void Creates_missing_directory()
    {
        using var dir = new TempDir();
        var path = Path.Combine(dir.Path, "under", "a.ics");
        AtomicFile.WriteAllText(path, "x");
        Assert.True(File.Exists(path));
    }

    [Fact]
    public void Replaces_file_that_is_open_for_reading()
    {
        using var dir = new TempDir();
        var path = dir.File("a.ics");
        AtomicFile.WriteAllText(path, "gammel");
        using (var reader = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        {
            AtomicFile.WriteAllText(path, "ny");
            Assert.Equal((byte)'g', (byte)reader.ReadByte()); // læseren ser stadig den gamle fil
        }
        Assert.Equal("ny", File.ReadAllText(path));
    }
}

public class SubscriptionStoreTests
{
    [Fact]
    public void Role_round_trips_and_old_files_have_none()
    {
        using var dir = new TempDir();
        var store = new SubscriptionStore(dir.File("abonnementer.json"));
        store.Save([new Subscription(new ScheduleRef(ScheduleKind.Employee, "1", "Anna", "AE", "leader"))]);
        Assert.Equal("leader", store.Load().Single().Schedule.Role);
        File.WriteAllText(dir.File("abonnementer.json"), """[{"Kind":"Employee","Id":"1","Name":"Anna","Initials":"AE"}]""");
        Assert.Equal("", store.Load().Single().Schedule.Role);
    }

    static readonly DateTimeOffset At = new(2026, 10, 3, 9, 30, 0, TimeSpan.Zero);

    [Fact]
    public void Missing_file_is_empty()
    {
        using var dir = new TempDir();
        Assert.Empty(new SubscriptionStore(dir.File("abonnementer.json")).Load());
    }

    [Fact]
    public void Round_trip_with_marks()
    {
        using var dir = new TempDir();
        var store = new SubscriptionStore(dir.File("abonnementer.json"));
        Subscription[] items =
        [
            new(new ScheduleRef(ScheduleKind.Employee, "1001", "Anna Eksempel", "AE"), AddedAt: At, FetchedAt: At.AddMinutes(1)),
            new(new ScheduleRef(ScheduleKind.Group, "88231", "7A"), ImportedAt: At, ImportHash: "abc123", ImportUntil: At.AddDays(90)),
            new(new ScheduleRef(ScheduleKind.Resource, "412", "53")),
        ];
        store.Save(items);
        Assert.Equal(items, store.Load());
    }

    // Fetched: et kalenderprogram har hentet filen efter brugerens seneste klik på hovedknappen (eller helt uden klik).
    [Fact]
    public void Fetched_counts_only_after_the_last_click()
    {
        var s = new Subscription(new ScheduleRef(ScheduleKind.Group, "5", "7A"));
        Assert.False(s.Fetched);
        Assert.True((s with { FetchedAt = At }).Fetched);
        Assert.True((s with { AddedAt = At, FetchedAt = At }).Fetched);
        Assert.False((s with { AddedAt = At, FetchedAt = At.AddSeconds(-1) }).Fetched);
        Assert.False((s with { AddedAt = At }).Fetched);
    }

    [Fact]
    public void Key_is_schedule_key()
    {
        var s = new Subscription(new ScheduleRef(ScheduleKind.Group, "5", "7A"));
        Assert.Equal("klasse-5", s.Key);
    }

    [Fact]
    public void Corrupt_file_is_backed_up_and_treated_as_empty()
    {
        using var dir = new TempDir();
        var path = dir.File("abonnementer.json");
        File.WriteAllText(path, "{ ikke json");
        Assert.Empty(new SubscriptionStore(path).Load());
        Assert.Equal("{ ikke json", File.ReadAllText(path + ".beskadiget"));
    }

    [Fact]
    public void Invalid_entries_are_skipped()
    {
        using var dir = new TempDir();
        var path = dir.File("abonnementer.json");
        File.WriteAllText(path, """[{"Kind":"Group","Id":"../x","Name":"Ond"},{"Kind":"Ukendt","Id":"1","Name":"X"},{"Kind":"7","Id":"2","Name":"Tal"},{"Kind":"Resource","Id":"4","Name":"Sal"}]""");
        Assert.Equal([new Subscription(new ScheduleRef(ScheduleKind.Resource, "4", "Sal"))], new SubscriptionStore(path).Load());
    }
}

public class ConfigStoreTests
{
    [Fact]
    public void Missing_file_gives_defaults()
    {
        using var dir = new TempDir();
        Assert.Equal(new AppConfig(), new ConfigStore(dir.File("config.json")).Load());
    }

    [Fact]
    public void Round_trip_writes_enum_as_text()
    {
        using var dir = new TempDir();
        var store = new ConfigStore(dir.File("config.json"));
        var config = new AppConfig(CalendarApp.OutlookImport, FirstRunDone: true);
        store.Save(config);
        Assert.Equal(config, store.Load());
        Assert.Contains("\"OutlookImport\"", File.ReadAllText(dir.File("config.json")));
    }

    // Porten gemmes for brugeren; ServerPort (9876, til en port er valgt) skrives ikke i filen.
    [Fact]
    public void Port_is_saved_and_defaults_to_9876()
    {
        using var dir = new TempDir();
        var store = new ConfigStore(dir.File("config.json"));
        Assert.Equal(9876, store.Load().ServerPort);
        store.Save(new AppConfig(Port: 9877));
        Assert.Equal(9877, store.Load().ServerPort);
        Assert.DoesNotContain("ServerPort", File.ReadAllText(dir.File("config.json")));
    }

    [Fact]
    public void Corrupt_file_gives_defaults()
    {
        using var dir = new TempDir();
        File.WriteAllText(dir.File("config.json"), "{\"CalendarApp\": \"Lotus\"");
        Assert.Equal(new AppConfig(), new ConfigStore(dir.File("config.json")).Load());
    }
}

public class FileLogTests
{
    [Fact]
    public void Appends_lines()
    {
        using var dir = new TempDir();
        var log = new FileLog(dir.File("aulasync.log"));
        log.Info("første");
        log.Error("anden", new InvalidOperationException("boom"));
        var text = File.ReadAllText(dir.File("aulasync.log"));
        Assert.Contains("INFO første", text);
        Assert.Contains("FEJL anden", text);
        Assert.Contains("boom", text);
    }

    [Fact]
    public void Rotates_when_too_big()
    {
        using var dir = new TempDir();
        var log = new FileLog(dir.File("aulasync.log"), maxBytes: 100);
        for (int i = 0; i < 10; i++) log.Info(new string('x', 30));
        Assert.True(File.Exists(dir.File("aulasync.log.old")));
        Assert.True(new FileInfo(dir.File("aulasync.log")).Length < 200);
    }
}
