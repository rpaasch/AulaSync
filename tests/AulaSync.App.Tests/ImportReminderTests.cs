using AulaSync.Core;

namespace AulaSync.App.Tests;

// Et importeret skema er ændret: AulaSync siger til med en besked, der åbner hovedvinduet (spec §3.2).
public class ImportReminderTests : IDisposable
{
    static readonly ScheduleRef SevenA = new(ScheduleKind.Group, "88231", "7A");

    readonly TestHost _host = new();
    readonly FakeNotifier _notifier = new();
    int _opened;

    public void Dispose() => _host.Dispose();

    async Task ChangeImportedScheduleAsync()
    {
        _host.Sync.SetClient(_host.Client);
        _host.Client.Events = (_, _, _) => [TestHost.Ahead("1")];
        await _host.Sync.AddAsync(SevenA, default);
        await _host.Sync.MarkImportedAsync(SevenA);
        _host.Client.Events = (_, _, _) => [TestHost.Ahead("1"), TestHost.Ahead("2")];
        await _host.Sync.SyncAllAsync(default);
    }

    void Start(CalendarApp app)
    {
        _host.Config.Save(new AppConfig(app, FirstRunDone: true));
        new ImportReminder(_host.Sync, _host.Config, new FakePlatform(isMac: false), _notifier, () => _opened++, a => a()).Start();
    }

    [Fact]
    public async Task Changed_import_shows_a_message_that_opens_the_main_window()
    {
        Start(CalendarApp.OutlookImport);
        await ChangeImportedScheduleAsync();
        Assert.Equal(["7A er ændret siden import. Klik her for at importere igen."], _notifier.Shown);
        _notifier.Click();
        Assert.Equal(1, _opened);
    }

    // Med et kalenderprogram, der abonnerer, er en ændret fil ikke noget at gøre ved.
    [Fact]
    public async Task No_message_when_the_calendar_subscribes()
    {
        Start(CalendarApp.AppleCalendar);
        await ChangeImportedScheduleAsync();
        Assert.Empty(_notifier.Shown);
    }
}
