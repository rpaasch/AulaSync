using Avalonia.Controls;
using Avalonia.Headless.XUnit;
using Avalonia.VisualTree;
using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.App.Tests;

public class OnboardingTests : IDisposable
{
    static readonly ScheduleRef Anna = new(ScheduleKind.Employee, "1001", "Anna Eksempel", "AE");

    readonly TestHost _host = new();
    readonly FakeAutostart _autostart = new();
    FakePlatform _platform = new(isMac: true);

    public OnboardingTests() => _host.Sync.SetClient(_host.Client);

    public void Dispose() => _host.Dispose();

    OnboardingViewModel Create(bool isMac = true)
    {
        _platform = new FakePlatform(isMac);
        var catalog = new ScheduleCatalog(new FakeDirectory());
        var add = new AddScheduleViewModel(() => catalog, () => Anna, _host.Sync, _host.Log, a => a());
        var main = new MainViewModel(_host.Sync, _host.Config, _platform, new FakeDialogs(), new FakeActions(),
            _host.Time, Copenhagen, () => true, () => true, a => a());
        return new OnboardingViewModel(_host.Config, _platform, _autostart, add, main, () => Anna);
    }

    [Fact]
    public async Task Five_steps_from_welcome_to_done()
    {
        var vm = Create();
        int finished = 0;
        vm.Finished += () => finished++;
        Assert.True(vm.IsWelcome);
        Assert.False(vm.ShowFooter);
        Assert.Equal([true, false, false, false, false], vm.Progress);

        vm.LogInCommand.Execute(null);
        Assert.True(vm.IsLogin);

        vm.SignedIn();
        Assert.True(vm.IsCalendarApp);
        Assert.True(vm.Cards.Single(c => c.App == CalendarApp.AppleCalendar).IsSelected); // standard på Mac
        Assert.False(vm.CanGoBack);

        await vm.ContinueCommand.ExecuteAsync(null);
        Assert.True(vm.IsSchedules);
        Assert.Equal(CalendarApp.AppleCalendar, _host.Config.Load().CalendarApp);
        Assert.Equal(1, vm.Add.SelectedCount); // eget skema forudvalgt
        Assert.True(vm.CanContinue);
        await WaitUntil(() => _host.Store.Load().Any(s => s.Key == Anna.Key));

        await vm.ContinueCommand.ExecuteAsync(null);
        Assert.True(vm.IsAddToCalendar);
        Assert.Equal("Færdig", vm.ContinueLabel);
        Assert.Equal(["AE Anna Eksempel"], vm.Main.Rows.Select(r => r.Name));
        Assert.Equal("Tilføj til Kalender", vm.Main.Rows.Single().PrimaryLabel);

        await vm.ContinueCommand.ExecuteAsync(null);
        Assert.Equal(1, finished);
        Assert.True(_host.Config.Load().FirstRunDone);
        Assert.True(_autostart.IsEnabled);
    }

    [Fact]
    public async Task Windows_has_no_default_and_shows_outlook_hint()
    {
        var vm = Create(isMac: false);
        vm.SignedIn();
        Assert.True(vm.ShowOutlookHint);
        Assert.DoesNotContain(vm.Cards, c => c.IsSelected);
        Assert.False(vm.CanContinue);

        vm.Cards.Single(c => c.App == CalendarApp.OutlookImport).SelectCommand.Execute(null);
        Assert.True(vm.CanContinue);
        await vm.ContinueCommand.ExecuteAsync(null);
        Assert.Equal(CalendarApp.OutlookImport, _host.Config.Load().CalendarApp);
    }

    [Fact]
    public void Card_texts_follow_the_design()
    {
        var card = Create().Cards.Single(c => c.App == CalendarApp.OutlookImport);
        Assert.Equal("Ny Outlook, Outlook til Mac eller web", card.Title);
        Assert.Equal("○ Øjebliksbillede", card.Mode);
        Assert.Equal("Henter kalendere gennem Microsofts servere, som ikke kan nå denne computer. Du importerer en fil, og AulaSync siger til, når skemaet er ændret, så du kan importere igen.", card.Description);
    }

    [Fact]
    public async Task Back_from_step_five_and_four()
    {
        var vm = Create();
        vm.SignedIn();
        await vm.ContinueCommand.ExecuteAsync(null);
        await vm.ContinueCommand.ExecuteAsync(null);
        vm.BackCommand.Execute(null);
        Assert.True(vm.IsSchedules);
        vm.BackCommand.Execute(null);
        Assert.True(vm.IsCalendarApp);
        vm.BackCommand.Execute(null);
        Assert.True(vm.IsCalendarApp);
    }

    [Fact]
    public void Windows_finish_message_mentions_taskbar() =>
        Assert.Equal("Når du klikker Færdig, kører AulaSync videre i systembakken. Træk ikonet ned på proceslinjen for at have det ved hånden.", Create(isMac: false).FinishMessage);

    // AulaSync serverer kalenderne fra denne computer: kalenderprogrammet skal køre samme sted (spec §2).
    [AvaloniaFact]
    public void Step_three_says_that_the_calendar_program_must_run_on_this_computer()
    {
        var vm = Create();
        var window = new OnboardingWindow(vm, new TextBlock { Text = "(login)" });
        window.Show();
        vm.Step = OnboardingStep.CalendarApp;
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();
        Assert.Contains(window.GetVisualDescendants().OfType<TextBlock>(), t => t.Text == OnboardingViewModel.SameComputerText && t.IsEffectivelyVisible);
        Assert.StartsWith("AulaSync og kalenderprogrammet skal køre på samme computer.", OnboardingViewModel.SameComputerText);
    }

    // Apple Kalender: abonnementet skal ligge "På min Mac". Med placeringen iCloud henter Apples servere filen, og de kan
    // ikke nå denne computer.
    [AvaloniaFact]
    public void Step_five_tells_apple_calendar_users_to_choose_on_my_mac()
    {
        var vm = Create();
        vm.SignedIn();
        Assert.True(vm.ShowAppleCalendarHint); // Apple Kalender er forvalgt på Mac
        vm.Cards.Single(c => c.App == CalendarApp.OutlookClassic).SelectCommand.Execute(null);
        Assert.False(vm.ShowAppleCalendarHint);

        vm.Cards.Single(c => c.App == CalendarApp.AppleCalendar).SelectCommand.Execute(null);
        var window = new OnboardingWindow(vm, new TextBlock { Text = "(login)" });
        window.Show();
        vm.Step = OnboardingStep.AddToCalendar;
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();
        Assert.Contains(window.GetVisualDescendants().OfType<TextBlock>(),
            t => t.Text == "Når Kalender spørger, så vælg Placering: På min Mac (ikke iCloud)." && t.IsEffectivelyVisible);
    }

    // Ingen notifikation ved Færdig: trin 5 siger selv, at AulaSync kører videre.
    [AvaloniaFact]
    public void Step_five_says_that_AulaSync_keeps_running()
    {
        var vm = Create();
        var window = new OnboardingWindow(vm, new TextBlock { Text = "(login)" });
        window.Show();
        vm.Step = OnboardingStep.AddToCalendar;
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();
        Assert.Contains(window.GetVisualDescendants().OfType<TextBlock>(),
            t => t.Text == "Når du klikker Færdig, kører AulaSync videre i menulinjen." && t.IsEffectivelyVisible);
    }

    // Windows: WebView2 starter, så snart login-visningen er i vinduet, også når den er skjult. Lukkede man velkomsten, mens
    // den startede, lukkede appen (E_ABORT). Login-visningen sættes derfor først ind ved trin 2.
    [AvaloniaFact]
    public void Login_view_is_added_at_the_login_step()
    {
        var vm = Create();
        var login = new TextBlock { Text = "(login)" };
        var window = new OnboardingWindow(vm, login);
        window.Show();
        Assert.Null(login.Parent);

        vm.Step = OnboardingStep.Login;
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();
        Assert.True(login.IsEffectivelyVisible);
    }

    [AvaloniaFact]
    public void Window_shows_welcome_then_cards()
    {
        var vm = Create();
        var window = new OnboardingWindow(vm, new TextBlock { Text = "(login)" });
        window.Show();
        Assert.Contains(window.GetVisualDescendants().OfType<Button>(), b => (string?)b.Content == "Log ind med Aula" && b.IsEffectivelyVisible);

        vm.SignedIn();
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();
        var radios = window.GetVisualDescendants().OfType<RadioButton>().ToList();
        Assert.Equal(4, radios.Count);
        Assert.Contains(window.GetVisualDescendants().OfType<TextBlock>(), t => t.Text == "● Opdateres automatisk" && t.IsEffectivelyVisible);
    }
}
