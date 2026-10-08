using System.Net;
using Avalonia.Controls;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.VisualTree;
using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.App.Tests;

public class SettingsViewModelTests : IDisposable
{
    readonly TestHost _host = new();
    readonly FakeDialogs _dialogs = new();
    readonly FakeAutostart _autostart = new(enabled: true);
    readonly FakePlatform _platform = new();
    readonly SessionController _session;
    int _signOuts, _appChanges, _intervalChanges;

    public SettingsViewModelTests()
    {
        _session = new SessionController(_host.Sync, _host.Log, _ => new HttpClient(new FakeHandler((r, _) =>
            r.RequestUri!.Query.Contains("getProfileContext") ? FakeHandler.Json(Fixture.Read("profileContext.json")) : FakeHandler.Json(Fixture.Read("profilesByLogin.json")))));
    }

    public void Dispose() => _host.Dispose();

    SettingsViewModel Create() => new(_host.Config, _autostart, _session, _dialogs, _platform, _host.Paths,
        () => { _signOuts++; return Task.CompletedTask; }, () => _appChanges++, () => _intervalChanges++);

    Task SignInAsync() => _session.SignInAsync([new Cookie("PHPSESSID", "a", "/", ".aula.dk"), new Cookie("Csrfp-Token", "b", "/", "www.aula.dk")], default);

    // Kalenderprogrammet vælges med de samme fire kort som ved første start (titel, opdatering, beskrivelse).
    [Fact]
    public void Calendar_app_choice_is_saved_and_changes_buttons()
    {
        var vm = Create();
        Assert.DoesNotContain(vm.Cards, c => c.IsSelected);
        Assert.Equal(["Apple Kalender", "Outlook (klassisk)", "Ny Outlook, Outlook til Mac eller web", "Andet program"], vm.Cards.Select(c => c.Title));
        vm.Cards[2].SelectCommand.Execute(null);
        Assert.Equal(CalendarApp.OutlookImport, _host.Config.Load().CalendarApp);
        Assert.Equal([false, false, true, false], vm.Cards.Select(c => c.IsSelected));
        Assert.Equal(1, _appChanges);
        vm.Cards[2].SelectCommand.Execute(null); // samme valg igen ændrer intet
        Assert.Equal(1, _appChanges);
    }

    [Fact]
    public void Saved_choice_is_the_selected_card()
    {
        _host.Config.Save(new AppConfig(CalendarApp.OutlookClassic, FirstRunDone: true));
        Assert.Equal([false, true, false, false], Create().Cards.Select(c => c.IsSelected));
    }

    [Fact]
    public void Start_at_login_toggles_autostart()
    {
        var vm = Create();
        Assert.True(vm.StartAtLogin);
        vm.StartAtLogin = false;
        Assert.False(_autostart.IsEnabled);
    }

    [Fact]
    public async Task Account_shows_name_initials_and_school()
    {
        var vm = Create();
        Assert.Equal("Ikke logget ind", vm.AccountText);
        Assert.False(vm.IsSignedIn);
        await SignInAsync();
        Assert.Equal("Logget ind som Test Bruger (TB)", vm.AccountText);
        Assert.Equal("Testskolen", vm.InstitutionText);
        Assert.True(vm.IsSignedIn);
    }

    [Fact]
    public async Task Log_out_only_after_confirmation()
    {
        var vm = Create();
        _dialogs.LogoutResult = false;
        await vm.LogOutCommand.ExecuteAsync(null);
        Assert.Equal(0, _signOuts);
        _dialogs.LogoutResult = true;
        await vm.LogOutCommand.ExecuteAsync(null);
        Assert.Equal(1, _signOuts);
    }

    [Fact]
    public void Troubleshooting_and_version()
    {
        var vm = Create();
        vm.OpenCalendarFolderCommand.Execute(null);
        vm.OpenLogCommand.Execute(null);
        Assert.Equal([_host.Paths.Calendars, _host.Paths.Log], _platform.Opened);
        Assert.Equal("Version 3.1.0", vm.VersionText);
    }

    // Hvor tit skemaerne hentes fra Aula: standard hver 4. time. Et nyt valg gemmes, og planen regnes om med det samme.
    [Fact]
    public void Update_interval_is_saved_and_reschedules()
    {
        var vm = Create();
        Assert.Equal(["Hver halve time", "Hver time", "Hver 2. time", "Hver 4. time (standard)", "Hver 8. time"], vm.Intervals.Select(i => i.Label));
        Assert.Equal("Hver 4. time (standard)", vm.SelectedInterval.ToString()); // det, en skærmlæser læser op
        Assert.Equal(240, vm.SelectedInterval.Minutes);

        vm.SelectedInterval = vm.Intervals[0];
        Assert.Equal(30, _host.Config.Load().UpdateMinutes);
        Assert.Equal(TimeSpan.FromMinutes(30), _host.Config.Load().UpdateInterval);
        Assert.Equal(1, _intervalChanges);
        Assert.Equal(30, Create().SelectedInterval.Minutes); // gemt
    }

    [Fact]
    public void Status_event_toggle_is_saved()
    {
        var vm = Create();
        Assert.False(vm.StatusEvent);
        vm.StatusEvent = true;
        Assert.True(_host.Config.Load().StatusEvent);
        Assert.True(Create().StatusEvent);
        Assert.Equal(0, _intervalChanges);
    }

    // Hjælp: vejledningen på GitHub, og første start kan gennemgås igen.
    [Fact]
    public void Help_opens_guide_and_first_start()
    {
        var vm = Create();
        vm.OpenGuideCommand.Execute(null);
        Assert.Equal(["https://github.com/rpaasch/AulaSync/blob/master/docs/vejledning.md"], _platform.Opened);
        vm.ShowOnboardingCommand.Execute(null);
        Assert.Equal(1, _dialogs.Onboardings);
    }

    [AvaloniaFact]
    public void Window_shows_rows()
    {
        var window = new SettingsWindow(Create());
        window.Show();
        var texts = window.GetVisualDescendants().OfType<TextBlock>().Select(t => t.Text).ToList();
        Assert.Contains("Kalenderprogram", texts);
        Assert.Contains("Fx Thunderbird. Kopiér en .ics-adresse, og indsæt den i et kalenderprogram på denne computer. Opdateres, så længe AulaSync kører.", texts); // kortene
        Assert.Single(window.GetVisualDescendants().OfType<ScrollViewer>(), v => v.Parent == window); // kan rulles på en lav skærm
        var combo = Assert.Single(window.GetVisualDescendants().OfType<ComboBox>()); // kalenderprogrammet er kort, ikke en rullemenu
        Assert.Equal(5, combo.ItemCount);
        Assert.Contains("Hent skemaer fra Aula", texts);
        Assert.Contains("Vis opdateringstid i kalenderen", texts);
        Assert.Contains("Start AulaSync, når jeg logger ind", texts);
        Assert.Contains("Hjælp", texts);
        Assert.Contains("Fejlfinding", texts);
        Assert.Contains("Version 3.1.0", texts);
        var buttons = window.GetVisualDescendants().OfType<Button>().Select(b => b.Content as string).ToList();
        Assert.Contains("Vejledning", buttons);
        Assert.Contains("Kom i gang igen…", buttons);
    }
}

public class TrayMenuTests
{
    static readonly DateTimeOffset At = new(2026, 10, 3, 12, 2, 0, TimeSpan.Zero);
    static readonly SyncStatus Ok = new(ConnectionState.Online, 5, At, At.AddHours(6), null, null);

    static List<string> Labels(SyncStatus s, bool server = true) => TrayMenu.Build(s, server, Copenhagen).Select(i => i.Label).ToList();

    [Fact]
    public void Normal_menu() =>
        Assert.Equal(["5 skemaer · opdateret 14:02", "-", "Åbn AulaSync…", "Opdatér nu", "Indstillinger…", "-", "Afslut AulaSync"], Labels(Ok));

    [Fact]
    public void Logged_out_puts_active_login_first()
    {
        var items = TrayMenu.Build(Ok with { Connection = ConnectionState.LoggedOut }, true, Copenhagen);
        Assert.Equal(new TrayItem("Log ind igen…", TrayCommand.LogIn), items[0]);
        Assert.Equal(new TrayItem("Logget ud af Aula", null), items[1]);
    }

    [Fact]
    public void Port_busy_offers_retry() =>
        Assert.Equal(["Start kalender-server igen", "Kalender-server kunne ikke starte"], Labels(Ok, server: false).Take(2));

    // Er et importeret skema ændret, står det øverst som en aktiv linje, der åbner hovedvinduet ("Importér igen…").
    [Fact]
    public void Changed_import_puts_an_open_action_first()
    {
        ScheduleRef[] changed = [new(ScheduleKind.Group, "3", "7A")];
        var items = TrayMenu.Build(Ok, true, Copenhagen, changed);
        Assert.Equal(new TrayItem("7A er ændret siden import…", TrayCommand.Open), items[0]);
        Assert.Equal(new TrayItem("5 skemaer · opdateret 14:02", null), items[1]);
        // Et problem, der kræver login, står stadig først.
        Assert.Equal("Log ind igen…", TrayMenu.Build(Ok with { Connection = ConnectionState.LoggedOut }, true, Copenhagen, changed)[0].Label);
    }

    [Fact]
    public void Mac_shortcuts()
    {
        var items = TrayMenu.Build(Ok, true, Copenhagen);
        Assert.Equal('r', items.Single(i => i.Command == TrayCommand.Refresh).Shortcut);
        Assert.Equal(',', items.Single(i => i.Command == TrayCommand.Settings).Shortcut);
        Assert.Equal('q', items.Single(i => i.Command == TrayCommand.Quit).Shortcut);
    }

    [Fact]
    public void Icon_states()
    {
        Assert.Equal(TrayState.Normal, TrayMenu.IconState(Ok, true));
        Assert.Equal(TrayState.Updating, TrayMenu.IconState(Ok with { Progress = new(1, 5) }, true));
        Assert.Equal(TrayState.Attention, TrayMenu.IconState(Ok with { Connection = ConnectionState.LoggedOut }, true));
        Assert.Equal(TrayState.Attention, TrayMenu.IconState(Ok, false));
        Assert.Equal(TrayState.Normal, TrayMenu.IconState(Ok with { Connection = ConnectionState.Offline }, true));
        Assert.Equal(TrayState.Attention, TrayMenu.IconState(Ok, true, importChanged: true)); // prikken
    }
}

// Avalonia på Mac låser sig til den første NativeMenu, et ikon får; en ny menu senere får appen til at gå ned
// ("The menu being updated does not match"). Ikonets menu fyldes derfor igen i samme NativeMenu.
public class NativeMenusTests
{
    static readonly SyncStatus Ok = SyncStatus.Initial(2) with { Connection = ConnectionState.Online };

    static void Click(NativeMenuItemBase item) => ((INativeMenuItemExporterEventsImplBridge)item).RaiseClicked();

    static List<string?> Labels(NativeMenu menu) =>
        menu.Items.Select(i => i is NativeMenuItem m ? m.Header : "-").ToList();

    [AvaloniaFact]
    public void Tray_menu_is_refilled_in_the_same_menu()
    {
        var menu = new NativeMenu();
        var run = new List<TrayCommand>();
        NativeMenus.Fill(menu, TrayMenu.Build(Ok, true, Copenhagen), run.Add, mac: true);
        var first = menu.Items.ToList();
        NativeMenus.Fill(menu, TrayMenu.Build(Ok with { Connection = ConnectionState.LoggedOut }, true, Copenhagen), run.Add, mac: true);

        Assert.Equal("Log ind igen…", Labels(menu)[0]);
        Assert.DoesNotContain(menu.Items, first.Contains);
        Assert.All(first, i => Assert.Null(i.Parent));
        Click(menu.Items[0]);
        Assert.Equal([TrayCommand.LogIn], run);
    }

    [AvaloniaTheory]
    [InlineData(true)]
    [InlineData(false)]
    public void Shortcuts_only_on_mac(bool mac)
    {
        var menu = new NativeMenu();
        NativeMenus.Fill(menu, TrayMenu.Build(Ok, true, Copenhagen), _ => { }, mac);
        var refresh = menu.Items.OfType<NativeMenuItem>().Single(i => i.Header == "Opdatér nu");
        var settings = menu.Items.OfType<NativeMenuItem>().Single(i => i.Header == "Indstillinger…");
        Assert.Equal(mac ? new KeyGesture(Key.R, KeyModifiers.Meta) : null, refresh.Gesture);
        Assert.Equal(mac ? new KeyGesture(Key.OemComma, KeyModifiers.Meta) : null, settings.Gesture);
        Assert.False(menu.Items.OfType<NativeMenuItem>().First().IsEnabled); // statuslinjen
    }

    // Mac-appens menu på dansk (Avalonias egne punkter er engelske og slået fra).
    [AvaloniaFact]
    public void App_menu_is_danish_with_settings_hide_and_quit()
    {
        int opened = 0, quit = 0;
        var menu = NativeMenus.AppMenu(() => opened++, () => quit++);
        Assert.Equal(["Indstillinger…", "-", "Skjul AulaSync", "Skjul andre", "Vis alle", "-", "Afslut AulaSync"], Labels(menu));
        NativeMenuItem Item(string header) => menu.Items.OfType<NativeMenuItem>().Single(i => i.Header == header);
        Assert.Equal(new KeyGesture(Key.OemComma, KeyModifiers.Meta), Item("Indstillinger…").Gesture);
        Assert.Equal(new KeyGesture(Key.H, KeyModifiers.Meta), Item("Skjul AulaSync").Gesture);
        Assert.Equal(new KeyGesture(Key.H, KeyModifiers.Meta | KeyModifiers.Alt), Item("Skjul andre").Gesture);
        Assert.Equal(new KeyGesture(Key.Q, KeyModifiers.Meta), Item("Afslut AulaSync").Gesture);
        Click(Item("Indstillinger…"));
        Click(Item("Afslut AulaSync"));
        Assert.Equal((1, 1), (opened, quit));
    }

    // At lukke hovedvinduet skjuler det — men kun når brugeren lukker det. Afslut (⌘Q, Dock) og log ud/genstart af
    // styresystemet skal kunne lukke det, ellers afvises de.
    [Theory]
    [InlineData(WindowCloseReason.WindowClosing, false, true)]
    [InlineData(WindowCloseReason.WindowClosing, true, false)]
    [InlineData(WindowCloseReason.ApplicationShutdown, false, false)]
    [InlineData(WindowCloseReason.OSShutdown, false, false)]
    [InlineData(WindowCloseReason.OwnerWindowClosing, false, false)]
    public void Main_window_hides_only_when_the_user_closes_it(WindowCloseReason reason, bool quitting, bool hides) =>
        Assert.Equal(hides, AppWindows.HidesInsteadOfClosing(reason, quitting));
}
