using System.Reflection;
using AulaSync.Core;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace AulaSync.App;

// Et valg under "Hent skemaer fra Aula". ToString er det, en skærmlæser læser op for det valgte.
public sealed record IntervalOption(int Minutes, string Label)
{
    public override string ToString() => Label;
}

// Indstillinger (spec §3.2): kalenderprogram, opdatering, start ved login, konto med Log ud…, hjælp, fejlfinding, version og
// Afinstallér AulaSync….
public sealed partial class SettingsViewModel : ObservableObject, ICalendarCards
{
    public const string GuideUrl = "https://github.com/rpaasch/AulaSync/blob/master/docs/vejledning.md";

    readonly ConfigStore _config;
    readonly IAutostart _autostart;
    readonly SessionController _session;
    readonly IDialogs _dialogs;
    readonly IPlatform _platform;
    readonly AppPaths _paths;
    readonly Func<Task> _signOut;
    readonly Action _calendarAppChanged;
    readonly Action _intervalChanged;
    readonly IUninstall _uninstall;

    // intervalChanged: baggrundsplanen regner næste opdatering om (BackgroundScheduler.Reschedule). uninstall:
    // Afinstallér AulaSync… (UninstallLauncher).
    public SettingsViewModel(ConfigStore config, IAutostart autostart, SessionController session, IDialogs dialogs, IPlatform platform,
        AppPaths paths, Func<Task> signOut, Action calendarAppChanged, Action intervalChanged, IUninstall uninstall)
    {
        _config = config;
        _autostart = autostart;
        _session = session;
        _dialogs = dialogs;
        _platform = platform;
        _paths = paths;
        _signOut = signOut;
        _calendarAppChanged = calendarAppChanged;
        _intervalChanged = intervalChanged;
        _uninstall = uninstall;
        var saved = config.Load();
        var chosen = saved.CalendarApp;
        Cards = CalendarApps.All.Select(a => new CalendarCardViewModel(a, SelectCard) { IsSelected = a == chosen }).ToList();
        Intervals = UpdateIntervals.Minutes.Select(m => new IntervalOption(m, UpdateIntervals.Label(m))).ToList();
        SelectedInterval = Intervals.Single(i => TimeSpan.FromMinutes(i.Minutes) == saved.UpdateInterval);
        StatusEvent = saved.StatusEvent;
        StartAtLogin = autostart.IsEnabled;
        session.Changed += () => OnPropertyChanged(string.Empty);
    }

    public IReadOnlyList<CalendarCardViewModel> Cards { get; }

    public IReadOnlyList<IntervalOption> Intervals { get; }

    [ObservableProperty] public partial IntervalOption SelectedInterval { get; set; }
    [ObservableProperty] public partial bool StatusEvent { get; set; }
    [ObservableProperty] public partial bool StartAtLogin { get; set; }

    public bool IsSignedIn => _session.Profile is not null;
    public string AccountText => _session.Profile is { } p
        ? $"Logget ind som {p.Name}{(p.Initials == "" ? "" : $" ({p.Initials})")}"
        : "Ikke logget ind";
    public string InstitutionText => _session.Profile?.InstitutionName ?? "";
    public string VersionText => $"Version {Version}";

    public static string Version =>
        (typeof(SettingsViewModel).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "?").Split('+')[0];

    void SelectCard(CalendarCardViewModel card)
    {
        foreach (var c in Cards) c.IsSelected = c == card;
        if (_config.Load().CalendarApp == card.App) return;
        _config.Save(_config.Load() with { CalendarApp = card.App });
        _calendarAppChanged(); // at skifte ændrer hovedknapperne, intet andet
    }

    partial void OnSelectedIntervalChanged(IntervalOption value)
    {
        if (value is null || _config.Load().UpdateInterval == TimeSpan.FromMinutes(value.Minutes)) return; // også ved start
        _config.Save(_config.Load() with { UpdateMinutes = value.Minutes });
        _intervalChanged();
    }

    // Kommer med i kalenderfilerne ved næste opdatering.
    partial void OnStatusEventChanged(bool value)
    {
        if (_config.Load().StatusEvent != value) _config.Save(_config.Load() with { StatusEvent = value });
    }

    partial void OnStartAtLoginChanged(bool value)
    {
        if (_autostart.IsEnabled != value) _autostart.SetEnabled(value);
    }

    [RelayCommand]
    async Task LogOut()
    {
        if (await _dialogs.ConfirmLogoutAsync()) await _signOut();
    }

    [RelayCommand] void LogIn() => _dialogs.ShowLogin();

    [RelayCommand] void OpenGuide() => _platform.Open(GuideUrl);

    [RelayCommand] void ShowOnboarding() => _dialogs.ShowOnboarding();

    [RelayCommand]
    void OpenCalendarFolder()
    {
        Directory.CreateDirectory(_paths.Calendars);
        _platform.Open(_paths.Calendars);
    }

    [RelayCommand] void OpenLog() => _platform.Open(_paths.Log);

    // Installationsprogrammets afinstallation spørger selv; ellers spørger AulaSync først.
    [RelayCommand]
    async Task Uninstall()
    {
        if (_uninstall.AsksItself || await _dialogs.ConfirmUninstallAsync()) _uninstall.Start();
    }
}
