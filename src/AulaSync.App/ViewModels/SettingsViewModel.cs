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

// Et valg under "Vis opdateringstid i kalenderen": fra, eller hvordan tidspunktet vises (StatusStyle).
public sealed record StatusOption(bool On, StatusStyle Style, string Label)
{
    public override string ToString() => Label;
}

// Et valg under "Besked, når du er logget ud af Aula" (LoggedOutNotice).
public sealed record NoticeOption(LoggedOutNotice Notice, string Label)
{
    public override string ToString() => Label;
}

// Indstillinger (spec §3.2): kalenderprogram, opdatering, start ved login, konto med Log ud…, hjælp, fejlfinding, version og
// Afinstallér AulaSync….
public sealed partial class SettingsViewModel : ObservableObject, ICalendarCards
{
    public const string GuideUrl = "https://github.com/rpaasch/AulaSync/blob/master/docs/vejledning.md";
    public const string WebsiteUrl = "https://rpaasch.github.io/AulaSync/";

    readonly ConfigStore _config;
    readonly IAutostart _autostart;
    readonly SessionController _session;
    readonly IDialogs _dialogs;
    readonly IPlatform _platform;
    readonly AppPaths _paths;
    readonly Func<Task> _signOut;
    readonly Action _calendarAppChanged;
    readonly Action _intervalChanged;
    readonly Action _statusEventChanged;
    readonly IUninstall _uninstall;

    // intervalChanged: baggrundsplanen regner næste opdatering om (BackgroundScheduler.Reschedule). statusEventChanged:
    // skemaerne hentes med det samme, så kalenderfilerne får eller mister statusaftalen (BackgroundScheduler.SyncSoon).
    // uninstall: Afinstallér AulaSync… (UninstallLauncher).
    public SettingsViewModel(ConfigStore config, IAutostart autostart, SessionController session, IDialogs dialogs, IPlatform platform,
        AppPaths paths, Func<Task> signOut, Action calendarAppChanged, Action intervalChanged, Action statusEventChanged, IUninstall uninstall)
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
        _statusEventChanged = statusEventChanged;
        _uninstall = uninstall;
        var saved = config.Load();
        var chosen = saved.CalendarApp;
        Cards = CalendarApps.All.Select(a => new CalendarCardViewModel(a, SelectCard) { IsSelected = a == chosen }).ToList();
        Intervals = UpdateIntervals.Minutes.Select(m => new IntervalOption(m, UpdateIntervals.Label(m))).ToList();
        SelectedInterval = Intervals.Single(i => TimeSpan.FromMinutes(i.Minutes) == saved.UpdateInterval);
        StatusOptions =
        [
            new(false, StatusStyle.Timed, "Fra"),
            new(true, StatusStyle.AllDay, "Heldagsbegivenhed hver mandag"),
            new(true, StatusStyle.Timed, "Aftale hver mandag"),
            new(true, StatusStyle.InDescription, "I hver lektions beskrivelse"),
        ];
        SelectedStatus = StatusOptions.FirstOrDefault(o => o.On && saved.StatusEvent && o.Style == saved.StatusStyle) ?? StatusOptions[0];
        StatusTime = saved.StatusTime.ToTimeSpan();
        NoticeOptions =
        [
            new(LoggedOutNotice.Daily, "Hver dag"),
            new(LoggedOutNotice.Once, "Én gang"),
            new(LoggedOutNotice.Never, "Ingen"),
        ];
        SelectedNotice = NoticeOptions.Single(o => o.Notice == saved.LoggedOutNotice);
        StartAtLogin = autostart.IsEnabled;
        session.Changed += () => OnPropertyChanged(string.Empty);
    }

    public IReadOnlyList<CalendarCardViewModel> Cards { get; }

    public IReadOnlyList<IntervalOption> Intervals { get; }

    [ObservableProperty] public partial IntervalOption SelectedInterval { get; set; }
    public IReadOnlyList<StatusOption> StatusOptions { get; }

    [ObservableProperty] [NotifyPropertyChangedFor(nameof(IsTimedStatus))] public partial StatusOption SelectedStatus { get; set; }
    [ObservableProperty] public partial TimeSpan? StatusTime { get; set; }

    // Klokkeslættet vises kun ved "Aftale hver mandag".
    public bool IsTimedStatus => SelectedStatus is { On: true, Style: StatusStyle.Timed };
    [ObservableProperty] public partial bool StartAtLogin { get; set; }

    public IReadOnlyList<NoticeOption> NoticeOptions { get; }
    [ObservableProperty] public partial NoticeOption SelectedNotice { get; set; }

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

    // Skemaerne hentes med det samme, så kalenderfilerne viser tidspunktet på den nye måde (eller ikke mere). Fra beholder
    // den måde, der var valgt.
    partial void OnSelectedStatusChanged(StatusOption value)
    {
        if (value is null) return;
        var saved = _config.Load();
        var next = value.On ? saved with { StatusEvent = true, StatusStyle = value.Style } : saved with { StatusEvent = false };
        if (next == saved) return; // også ved start
        _config.Save(next);
        _statusEventChanged();
    }

    // Aftalens klokkeslæt hver mandag; vises aftalen, hentes skemaerne med det samme, så den flytter.
    partial void OnStatusTimeChanged(TimeSpan? value)
    {
        if (value is not { } time || time < TimeSpan.Zero || time >= TimeSpan.FromDays(1)) return;
        var saved = _config.Load();
        var minutes = (int)time.TotalMinutes;
        if (saved.StatusTime == TimeOnly.FromTimeSpan(TimeSpan.FromMinutes(minutes))) return; // også ved start
        _config.Save(saved with { StatusMinutes = minutes });
        if (saved is { StatusEvent: true, StatusStyle: StatusStyle.Timed }) _statusEventChanged();
    }

    // Gælder fra næste besked (Relogin læser den hver gang).
    partial void OnSelectedNoticeChanged(NoticeOption value)
    {
        if (value is null || _config.Load().LoggedOutNotice == value.Notice) return; // også ved start
        _config.Save(_config.Load() with { LoggedOutNotice = value.Notice });
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

    [RelayCommand] void OpenWebsite() => _platform.Open(WebsiteUrl);

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
