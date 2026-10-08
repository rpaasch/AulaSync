using Avalonia.Controls;
using Avalonia.Headless.XUnit;
using Avalonia.VisualTree;
using AulaSync.Core;

namespace AulaSync.App.Tests;

public class MainActionsTests : IDisposable
{
    static readonly ScheduleRef SevenA = new(ScheduleKind.Group, "88231", "7A");

    readonly TestHost _host = new();
    readonly FakeDialogs _dialogs = new();

    public MainActionsTests() => _host.Sync.SetClient(_host.Client);

    public void Dispose() => _host.Dispose();

    async Task<(MainActions Actions, FakePlatform Platform)> CreateAsync(CalendarApp? app, bool isMac = true, int? port = null)
    {
        _host.Config.Save(new AppConfig(app, Port: port));
        await _host.Sync.AddAsync(SevenA, default);
        var platform = new FakePlatform(isMac);
        return (new MainActions(_host.Sync, _host.Config, platform, _dialogs), platform);
    }

    Subscription Sub => _host.Sync.Subscriptions.Single();

    // Apple Kalender laver webcal:// om til https (som AulaSync ikke taler) og prøver ikke http bagefter, så Kalender får
    // http-adressen direkte.
    [Fact]
    public async Task Apple_calendar_gets_the_http_address_and_marks_added()
    {
        var (actions, platform) = await CreateAsync(CalendarApp.AppleCalendar);
        DateTimeOffset? addedWhenOpened = null;
        platform.OnOpen = _ => addedWhenOpened = Sub.AddedAt; // markeret før Kalender kan nå at hente
        await actions.RunAsync(SevenA);
        Assert.Equal(["http://localhost:9876/klasse-88231.ics"], platform.OpenedInCalendar);
        Assert.Empty(platform.Opened);
        Assert.Equal(TestHost.Start, Sub.AddedAt);
        Assert.Equal(TestHost.Start, addedWhenOpened);
        Assert.Empty(_dialogs.Fallbacks);
    }

    [Fact]
    public async Task Mac_without_choice_uses_apple_calendar()
    {
        var (actions, platform) = await CreateAsync(null);
        await actions.RunAsync(SevenA);
        Assert.Equal(["http://localhost:9876/klasse-88231.ics"], platform.OpenedInCalendar);
    }

    [Fact]
    public async Task Classic_outlook_also_shows_fallback()
    {
        var (actions, platform) = await CreateAsync(CalendarApp.OutlookClassic, isMac: false);
        await actions.RunAsync(SevenA);
        Assert.Equal(["webcal://localhost:9876/klasse-88231.ics"], platform.Opened);
        Assert.Empty(platform.OpenedInCalendar);
        Assert.Equal([SevenA], _dialogs.Fallbacks);
    }

    // Porten er brugerens egen (9877, når 9876 var optaget ved første start).
    [Fact]
    public async Task Addresses_use_the_users_port()
    {
        var (apple, platform) = await CreateAsync(CalendarApp.AppleCalendar, port: 9877);
        await apple.RunAsync(SevenA);
        await apple.CopyAddressAsync(SevenA);
        Assert.Equal(["http://localhost:9877/klasse-88231.ics"], platform.OpenedInCalendar);
        Assert.Equal(["http://localhost:9877/klasse-88231.ics"], platform.Copied);

        _host.Config.Save(new AppConfig(CalendarApp.OutlookClassic, Port: 9877));
        await apple.RunAsync(SevenA);
        Assert.Equal(["webcal://localhost:9877/klasse-88231.ics"], platform.Opened);
    }

    [Fact]
    public async Task Import_marks_imported_only_when_guide_is_completed()
    {
        var (actions, _) = await CreateAsync(CalendarApp.OutlookImport);
        _dialogs.ImportGuideResult = false;
        await actions.RunAsync(SevenA);
        Assert.Null(Sub.ImportedAt);

        _dialogs.ImportGuideResult = true;
        await actions.RunAsync(SevenA);
        Assert.Equal(TestHost.Start, Sub.ImportedAt);
        Assert.Equal(IcsInspect.ReadFile(_host.Sync.FilePath(SevenA), Sub.ImportedAt, Sub.ImportUntil)!.ContentHash, Sub.ImportHash);
    }

    [Fact]
    public async Task Other_program_copies_http_address()
    {
        var (actions, platform) = await CreateAsync(CalendarApp.Other);
        await actions.RunAsync(SevenA);
        Assert.Equal(["http://localhost:9876/klasse-88231.ics"], platform.Copied);
        Assert.NotNull(Sub.AddedAt);
    }

    // Holder et andet program udklipsholderen, er adressen ikke kopieret, og skemaet markeres ikke som tilføjet.
    [Fact]
    public async Task Other_program_does_not_mark_added_when_the_copy_fails()
    {
        var (actions, platform) = await CreateAsync(CalendarApp.Other);
        platform.CopyWorks = false;
        await actions.RunAsync(SevenA);
        Assert.Null(Sub.AddedAt);
    }

    [Fact]
    public async Task Menu_copy_and_reveal()
    {
        var (actions, platform) = await CreateAsync(CalendarApp.AppleCalendar);
        await actions.CopyAddressAsync(SevenA);
        actions.RevealFile(SevenA);
        Assert.Equal(["http://localhost:9876/klasse-88231.ics"], platform.Copied);
        Assert.Equal([_host.Sync.FilePath(SevenA)], platform.Revealed);
        Assert.Null(Sub.AddedAt); // kun hovedknappen markerer
    }
}

public class ImportGuideTests
{
    static readonly ScheduleRef SevenA = new(ScheduleKind.Group, "4711", "7A");

    [Fact]
    public async Task Buttons_copy_name_open_outlook_and_reveal_file()
    {
        var platform = new FakePlatform(isMac: false);
        var time = new Microsoft.Extensions.Time.Testing.FakeTimeProvider();
        // Filen kan få nyt navn, mens guiden er åben (første opdatering efter 3.0.0); stien slås op, når den bruges.
        var path = "/data/kalendere/klasse-4711.ics";
        var vm = new ImportGuideViewModel(SevenA, () => path, platform, time);
        Assert.Equal("Importér 7A", vm.Title);
        Assert.Equal("klasse-4711.ics", vm.FileName);
        path = "/data/kalendere/101001-7A-klasse-4711.ics";
        Assert.Equal("101001-7A-klasse-4711.ics", vm.FileName); // filens navn, som man ser det, når man vælger den
        Assert.Equal("Vis fil i Stifinder", vm.RevealLabel);

        var copy = vm.CopyNameCommand.ExecuteAsync(null);
        Assert.Equal("✓ Kopieret", vm.CopyNameLabel);
        time.Advance(TimeSpan.FromSeconds(2));
        await copy;
        Assert.Equal("Kopiér navn", vm.CopyNameLabel);

        vm.OpenOutlookCommand.Execute(null);
        vm.RevealCommand.Execute(null);
        Assert.Equal(["7A"], platform.Copied);
        Assert.Equal(["https://outlook.office.com/calendar"], platform.Opened);
        Assert.Equal(["/data/kalendere/101001-7A-klasse-4711.ics"], platform.Revealed);
    }

    [Fact]
    public async Task Outlook_fallback_copies_http_address_with_the_users_port()
    {
        var platform = new FakePlatform(isMac: false);
        var time = new Microsoft.Extensions.Time.Testing.FakeTimeProvider();
        var vm = new OutlookFallbackViewModel(SevenA, 9877, platform, time);
        Assert.Equal("http://localhost:9877/klasse-4711.ics", vm.Address);
        var copy = vm.CopyAddressCommand.ExecuteAsync(null);
        Assert.Equal("✓ Kopieret", vm.CopyLabel);
        time.Advance(TimeSpan.FromSeconds(2));
        await copy;
        Assert.Equal(["http://localhost:9877/klasse-4711.ics"], platform.Copied);
        Assert.Equal("Kopiér adresse", vm.CopyLabel);
    }

    // Holder et andet program udklipsholderen, siger knappen det i stedet for "✓ Kopieret".
    [Fact]
    public async Task Copy_buttons_say_when_the_copy_failed()
    {
        var platform = new FakePlatform(isMac: false) { CopyWorks = false };
        var time = new Microsoft.Extensions.Time.Testing.FakeTimeProvider();
        var guide = new ImportGuideViewModel(SevenA, () => "/x.ics", platform, time);
        var fallback = new OutlookFallbackViewModel(SevenA, 9876, platform, time);

        var copyName = guide.CopyNameCommand.ExecuteAsync(null);
        var copyAddress = fallback.CopyAddressCommand.ExecuteAsync(null);
        Assert.Equal("Kunne ikke kopiere", guide.CopyNameLabel);
        Assert.Equal("Kunne ikke kopiere", fallback.CopyLabel);
        time.Advance(TimeSpan.FromSeconds(2));
        await Task.WhenAll(copyName, copyAddress);
        Assert.Equal("Kopiér navn", guide.CopyNameLabel);
        Assert.Equal("Kopiér adresse", fallback.CopyLabel);
    }

    [AvaloniaFact]
    public void Guide_window_shows_three_steps_and_hint()
    {
        var window = new ImportGuideWindow(new ImportGuideViewModel(SevenA, () => "/x.ics", new FakePlatform(), TimeProvider.System));
        window.Show();
        var buttons = window.GetVisualDescendants().OfType<Button>().Select(b => b.Content as string).ToList();
        Assert.Contains("Kopiér navn", buttons);
        Assert.Contains("Åbn Outlook på nettet", buttons);
        Assert.Contains("Vis fil i Finder", buttons);
        Assert.Contains(window.GetVisualDescendants().OfType<TextBlock>(), t => t.Text?.StartsWith("Importen er et øjebliksbillede.") == true);
    }
}
