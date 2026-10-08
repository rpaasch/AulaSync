using System.Reflection;
using AulaSync.Core;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace AulaSync.App;

// Indstillinger (spec §3.2): kalenderprogram, start ved login, konto med Log ud…, fejlfinding og version.
public sealed partial class SettingsViewModel : ObservableObject, ICalendarCards
{
    readonly ConfigStore _config;
    readonly IAutostart _autostart;
    readonly SessionController _session;
    readonly IDialogs _dialogs;
    readonly IPlatform _platform;
    readonly AppPaths _paths;
    readonly Func<Task> _signOut;
    readonly Action _calendarAppChanged;

    public SettingsViewModel(ConfigStore config, IAutostart autostart, SessionController session, IDialogs dialogs, IPlatform platform,
        AppPaths paths, Func<Task> signOut, Action calendarAppChanged)
    {
        _config = config;
        _autostart = autostart;
        _session = session;
        _dialogs = dialogs;
        _platform = platform;
        _paths = paths;
        _signOut = signOut;
        _calendarAppChanged = calendarAppChanged;
        var chosen = config.Load().CalendarApp;
        Cards = CalendarApps.All.Select(a => new CalendarCardViewModel(a, SelectCard) { IsSelected = a == chosen }).ToList();
        StartAtLogin = autostart.IsEnabled;
        session.Changed += () => OnPropertyChanged(string.Empty);
    }

    public IReadOnlyList<CalendarCardViewModel> Cards { get; }

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

    [RelayCommand]
    void OpenCalendarFolder()
    {
        Directory.CreateDirectory(_paths.Calendars);
        _platform.Open(_paths.Calendars);
    }

    [RelayCommand] void OpenLog() => _platform.Open(_paths.Log);
}
