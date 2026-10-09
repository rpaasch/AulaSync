using Avalonia.Controls;
using Avalonia.Headless.XUnit;
using Avalonia.VisualTree;
using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.App.Tests;

public class MainViewModelTests : IDisposable
{
    static readonly ScheduleRef Anna = new(ScheduleKind.Employee, "1001", "Anna Eksempel", "AE");
    static readonly ScheduleRef SevenA = new(ScheduleKind.Group, "88231", "7A");
    static readonly ScheduleRef Room = new(ScheduleKind.Resource, "412", "Lokale 53");

    readonly TestHost _host = new();
    readonly FakeDialogs _dialogs = new();
    readonly FakeActions _actions = new();
    bool _serverRunning = true;
    int _retries;

    public void Dispose() => _host.Dispose();

    MainViewModel Create(bool isMac = true) =>
        new(_host.Sync, _host.Config, new FakePlatform(isMac), _dialogs, _actions, _host.Time, Copenhagen,
            () => _serverRunning, () => { _retries++; return _serverRunning = true; }, a => a());

    async Task AddAsync(params ScheduleRef[] schedules)
    {
        _host.Sync.SetClient(_host.Client);
        foreach (var s in schedules) await _host.Sync.AddAsync(s, default);
    }

    [Fact]
    public void Empty_list_shows_empty_state()
    {
        var vm = Create();
        Assert.True(vm.IsEmpty);
        Assert.Empty(vm.Rows);
    }

    [Fact]
    public async Task Rows_show_badge_counts_and_main_button_for_apple_calendar()
    {
        _host.Client.Events = (_, _, _) => [Ev(id: "1"), Ev(id: "2")];
        await AddAsync(Anna, SevenA, Room);
        var vm = Create();

        Assert.Equal(["AE", "7A", "53"], vm.Rows.Select(r => r.BadgeText));
        Assert.Equal(["AE Anna Eksempel", "7A", "Lokale 53"], vm.Rows.Select(r => r.Name));
        var anna = vm.Rows[0];
        Assert.Equal("Medarbejder · 2 lektioner", anna.SubText);
        Assert.Equal("Tilføj til Kalender", anna.PrimaryLabel);
        Assert.False(anna.HasDone);
        Assert.Equal("Flere handlinger for AE Anna Eksempel", anna.MoreLabel);
        Assert.Equal("Vis fil i Finder", anna.RevealLabel);
    }

    // "✓ Tilføjet" først, når Kalender har hentet skemaet; uden hentning kommer knappen igen efter FetchWait.
    [Fact]
    public async Task Clicked_row_waits_for_the_calendar()
    {
        await AddAsync(SevenA, Room);
        var vm = Create();
        await _host.Sync.MarkAddedAsync(SevenA);
        await _host.Sync.MarkAddedAsync(Room);
        var (sevenA, room) = (vm.Rows[0], vm.Rows[1]);
        Assert.Equal("Venter på Kalender…", sevenA.DoneText);
        Assert.True(sevenA.IsWaiting); // grå, ikke grøn som "✓ Tilføjet"
        Assert.False(sevenA.HasPrimary);
        Assert.False(sevenA.HasNote);

        await _host.Sync.MarkFetchedAsync(SevenA.FileName);
        Assert.Equal("✓ Tilføjet", sevenA.DoneText);
        Assert.False(sevenA.IsWaiting);

        _host.Time.Advance(RowPresenter.FetchWait); // Lokale 53 blev aldrig hentet
        Assert.Equal("Tilføj til Kalender", room.PrimaryLabel);
        Assert.Equal("Kalender hentede ikke skemaet", room.Note);
        Assert.False(room.HasDone);
        Assert.Equal("✓ Tilføjet", sevenA.DoneText);
    }

    // Slettes en kalender i Outlook, henter Outlook stadig de andre skemaer, men ikke dette: rækken får en note og knappen
    // igen. Hentes skemaet igen, står der "✓ Tilføjet".
    [Fact]
    public async Task Row_notices_when_outlook_stops_fetching_a_schedule()
    {
        const string outlook = "Microsoft Office/16.0 (Windows NT 10.0; Microsoft Outlook 16.0; Pro)";
        _host.Config.Save(new AppConfig(CalendarApp.OutlookClassic, FirstRunDone: true));
        await AddAsync(SevenA, Room);
        var vm = Create(isMac: false);
        await _host.Sync.MarkFetchedAsync(SevenA.FileName, outlook);
        await _host.Sync.MarkFetchedAsync(Room.FileName, outlook);
        var (sevenA, room) = (vm.Rows[0], vm.Rows[1]);
        for (var i = 0; i < 12; i++) // 7A er slettet i Outlook; Outlook henter Lokale 53 hver halve time
        {
            _host.Time.Advance(TimeSpan.FromMinutes(30));
            await _host.Sync.MarkFetchedAsync(Room.FileName, outlook);
        }
        Assert.Equal("Tilføj til Outlook", sevenA.PrimaryLabel);
        Assert.Equal("Ikke hentet af Outlook siden kl. 14:00", sevenA.Note);
        Assert.False(sevenA.HasDone);
        Assert.Equal("✓ Tilføjet", room.DoneText);

        await _host.Sync.MarkFetchedAsync(SevenA.FileName, outlook);
        Assert.Equal("✓ Tilføjet", sevenA.DoneText);
        Assert.False(sevenA.HasNote);
    }

    [Fact]
    public async Task Failed_schedule_shows_error_and_no_main_button()
    {
        _host.Client.Events = (_, _, _) => throw new AulaException("nede");
        await AddAsync(Room);
        var row = Create().Rows.Single();
        Assert.True(row.SubIsError);
        Assert.StartsWith("Kunne ikke hentes kl. 14:00", row.SubText);
        Assert.False(row.HasPrimary);
    }

    [Fact]
    public async Task Import_app_shows_import_button_and_changed_mark()
    {
        _host.Config.Save(new AppConfig(CalendarApp.OutlookImport, FirstRunDone: true));
        _host.Client.Events = (_, _, _) => [TestHost.Ahead("1")];
        await AddAsync(SevenA);
        var vm = Create(isMac: false);
        Assert.Equal("Importér…", vm.Rows.Single().PrimaryLabel);
        Assert.Equal("Vis fil i Stifinder", vm.Rows.Single().RevealLabel);

        await _host.Sync.MarkImportedAsync(SevenA);
        _host.Client.Events = (_, _, _) => [TestHost.Ahead("1"), TestHost.Ahead("2")];
        await _host.Sync.SyncAllAsync(default);

        var row = vm.Rows.Single();
        Assert.Equal("Importeret i dag", row.DoneText);
        Assert.Equal("Ændret siden import", row.Note);
        Assert.Equal("Importér igen…", row.PrimaryLabel);
    }

    [Fact]
    public async Task Changing_calendar_app_changes_buttons_on_reload()
    {
        await AddAsync(SevenA);
        var vm = Create();
        _host.Config.Save(new AppConfig(CalendarApp.Other));
        vm.Reload();
        Assert.Equal("Kopiér adresse", vm.Rows.Single().PrimaryLabel);
    }

    [Fact]
    public async Task Menu_commands_go_to_actions()
    {
        await AddAsync(SevenA);
        var row = Create().Rows.Single();
        await row.PrimaryCommand.ExecuteAsync(null);
        await row.CopyAddressCommand.ExecuteAsync(null);
        row.RevealCommand.Execute(null);
        Assert.Equal(["run klasse-88231", "copy klasse-88231", "reveal klasse-88231"], _actions.Calls);
    }

    [Fact]
    public async Task Remove_is_confirmed_in_the_row()
    {
        await AddAsync(SevenA, Room);
        var vm = Create();
        var row = vm.Rows[0];
        row.AskRemoveCommand.Execute(null);
        Assert.True(row.IsConfirmingRemove);
        Assert.Equal("Fjern 7A? Kalenderen holder op med at blive opdateret. Slet den også i dit kalenderprogram.", row.ConfirmText);

        vm.Reload(); // en statusændring lukker ikke bekræftelsen
        Assert.True(vm.Rows[0].IsConfirmingRemove);

        await row.ConfirmRemoveCommand.ExecuteAsync(null);
        Assert.Equal(["Lokale 53"], vm.Rows.Select(r => r.Name));
        Assert.False(File.Exists(_host.Sync.FilePath(SevenA)));
    }

    [Fact]
    public void Logged_out_banner_opens_login()
    {
        var vm = Create();
        Assert.Equal("Du er logget ud af Aula. Kalenderne viser stadig de seneste skemaer.", vm.BannerText);
        Assert.Equal("Log ind igen", vm.BannerAction);
        Assert.True(vm.IsError);
        vm.RunBannerActionCommand.Execute(null);
        Assert.Equal(1, _dialogs.Logins);
    }

    [Fact]
    public void Port_busy_banner_retries_server()
    {
        _host.Sync.SetClient(_host.Client);
        _serverRunning = false;
        var vm = Create();
        Assert.Equal("Prøv igen", vm.BannerAction);
        vm.RunBannerActionCommand.Execute(null);
        Assert.Equal(1, _retries);
        Assert.False(vm.HasBanner);
    }

    [Fact]
    public async Task Bottom_bar_and_progress()
    {
        await AddAsync(SevenA);
        await _host.Sync.SyncAllAsync(default);
        var vm = Create();
        _host.Sync.ReportNextSync(TestHost.Start.AddHours(6));
        Assert.Equal("Opdateret i dag 14:00 · næste 20:00", vm.BottomText);
        Assert.Equal("Opdatér nu", vm.RefreshLabel);

        var seen = new List<string>();
        vm.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(vm.RefreshLabel)) seen.Add(vm.RefreshLabel); };
        await vm.RefreshCommand.ExecuteAsync(null);
        Assert.Contains("Henter 1 af 1…", seen);
        Assert.Equal("Opdatér nu", vm.RefreshLabel);
    }

    [Fact]
    public async Task Add_button_opens_dialog()
    {
        await Create().AddScheduleCommand.ExecuteAsync(null);
        Assert.Equal(1, _dialogs.AddSchedules);
    }
}

public class MainWindowTests : IDisposable
{
    readonly TestHost _host = new();

    public void Dispose() => _host.Dispose();

    [AvaloniaFact]
    public async Task Shows_rows_and_accessible_menu_button()
    {
        _host.Sync.SetClient(_host.Client);
        await _host.Sync.AddAsync(new ScheduleRef(ScheduleKind.Group, "88231", "7A"), default);
        var vm = new MainViewModel(_host.Sync, _host.Config, new FakePlatform(), new FakeDialogs(), new FakeActions(),
            _host.Time, Copenhagen, () => true, () => true, a => a());
        var window = new MainWindow(vm);
        window.Show();

        var texts = window.GetVisualDescendants().OfType<TextBlock>().Select(t => t.Text).ToList();
        Assert.Contains("Mine skemaer", texts);
        Assert.Contains("7A", texts);
        Assert.Contains("Klasse · 0 lektioner", texts);
        var more = window.GetVisualDescendants().OfType<Button>().Single(b => (string?)b.Content == "⋯");
        Assert.Equal("Flere handlinger for 7A", Avalonia.Automation.AutomationProperties.GetName(more));
    }
}
