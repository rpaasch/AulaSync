using System.Collections.ObjectModel;
using AulaSync.Core;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace AulaSync.App;

// Hovedvinduet "Mine skemaer" (spec §3.2): rækker, ét banner ad gangen, bundlinje med status og "Opdatér nu".
public sealed partial class MainViewModel : ObservableObject
{
    readonly SyncService _sync;
    readonly ConfigStore _config;
    readonly IPlatform _platform;
    readonly IDialogs _dialogs;
    readonly IMainActions _actions;
    readonly TimeProvider _time;
    readonly TimeZoneInfo _tz;
    readonly Func<bool> _serverRunning;
    readonly Func<bool> _retryServer;
    readonly Action<Action> _ui;
    Banner? _banner;
    ITimer? _rowTimer;
    DateTimeOffset? _rowTimerAt;

    public MainViewModel(SyncService sync, ConfigStore config, IPlatform platform, IDialogs dialogs, IMainActions actions,
        TimeProvider time, TimeZoneInfo tz, Func<bool> serverRunning, Func<bool> retryServer, Action<Action> ui)
    {
        _sync = sync;
        _config = config;
        _platform = platform;
        _dialogs = dialogs;
        _actions = actions;
        _time = time;
        _tz = tz;
        _serverRunning = serverRunning;
        _retryServer = retryServer;
        _ui = ui;
        sync.StatusChanged += _ => _ui(Reload);
        Reload();
    }

    public ObservableCollection<ScheduleRowViewModel> Rows { get; } = [];

    [ObservableProperty] public partial bool IsEmpty { get; set; }
    [ObservableProperty] public partial string? BannerText { get; set; }
    [ObservableProperty] public partial string? BannerAction { get; set; }
    [ObservableProperty] public partial string BottomText { get; set; } = "";
    [ObservableProperty] public partial string RefreshLabel { get; set; } = "Opdatér nu";
    [ObservableProperty] public partial bool IsRefreshing { get; set; }
    [ObservableProperty] public partial bool IsWarning { get; set; }
    [ObservableProperty] public partial bool IsError { get; set; }

    public bool HasBanner => BannerText is not null;
    public bool HasBannerAction => BannerAction is not null;

    partial void OnBannerTextChanged(string? value) => OnPropertyChanged(nameof(HasBanner));
    partial void OnBannerActionChanged(string? value) => OnPropertyChanged(nameof(HasBannerAction));

    // Læser alt igen fra SyncService og config. Kaldes på UI-tråden.
    public void Reload()
    {
        var status = _sync.Status;
        var serverRunning = _serverRunning();
        var now = _time.GetUtcNow();
        var app = CalendarChoice.Current(_config, _platform.IsMac);
        var subscriptions = _sync.Subscriptions;
        // Kører kalender-serveren ikke, kan intet hentes; det siger banneret, så rækkerne siger intet om hentning.
        FetchCheck? Fetch(Subscription s) => serverRunning ? _sync.CheckFetch(s) : null;

        var keys = subscriptions.Select(s => s.Key).ToHashSet();
        for (int i = Rows.Count - 1; i >= 0; i--)
            if (!keys.Contains(Rows[i].Key)) Rows.RemoveAt(i);
        for (int i = 0; i < subscriptions.Count; i++)
        {
            var s = subscriptions[i];
            var row = Rows.FirstOrDefault(r => r.Key == s.Key);
            if (row is null)
            {
                row = new ScheduleRowViewModel(s.Schedule, _actions, RemoveAsync, _platform.FileManager);
                Rows.Insert(Math.Min(i, Rows.Count), row);
            }
            row.Update(StatusText.Row(s.Schedule, _sync.StateOf(s.Schedule), status.NextSync, _tz),
                RowPresenter.Buttons(app, s, _sync.ChangedSinceImport(s), now, _tz, Fetch(s)));
        }
        IsEmpty = Rows.Count == 0;
        ScheduleRowChange(subscriptions.Select(s => RowPresenter.ChangesAt(app, s, now, Fetch(s))).Min(), now);

        _banner = StatusText.Banner(status, serverRunning, _config.Load().ServerPort);
        BannerText = _banner?.Text;
        BannerAction = _banner?.ActionLabel;

        BottomText = StatusText.BottomBar(status with { Progress = null }, now, _tz);
        IsRefreshing = status.Progress is not null;
        RefreshLabel = status.Progress is { } p ? $"Henter {p.Done} af {p.Total}…" : "Opdatér nu";
        var level = StatusText.Level(status, serverRunning);
        IsWarning = level == StatusLevel.Warning;
        IsError = level == StatusLevel.Error;
    }

    // En række, der venter på kalenderprogrammet, skifter af sig selv efter RowPresenter.FetchWait, og en række, der ikke er
    // hentet længe, efter FetchWatch' frister; så læses der igen.
    // Højst en dag ad gangen (en timer kan højst vente ca. 49 dage, og uret kan være stillet forkert), og efter hver gang
    // sættes den igen, også hvis den kom lidt før tid.
    void ScheduleRowChange(DateTimeOffset? at, DateTimeOffset now)
    {
        if (at == _rowTimerAt) return;
        _rowTimer?.Dispose();
        _rowTimerAt = at;
        var due = at is { } when ? TimeSpan.FromTicks(Math.Clamp((when - now).Ticks, 0, TimeSpan.TicksPerDay)) : (TimeSpan?)null;
        _rowTimer = due is { } wait
            ? _time.CreateTimer(_ => _ui(() => { _rowTimerAt = null; Reload(); }), null, wait, Timeout.InfiniteTimeSpan)
            : null;
    }

    [RelayCommand] Task AddSchedule() => _dialogs.ShowAddScheduleAsync();

    [RelayCommand] Task Refresh() => _sync.SyncAllAsync(CancellationToken.None);

    [RelayCommand]
    void RunBannerAction()
    {
        switch (_banner?.Kind)
        {
            case BannerKind.LoggedOut: _dialogs.ShowLogin(); break;
            case BannerKind.PortBusy: _retryServer(); Reload(); break;
        }
    }

    async Task RemoveAsync(ScheduleRef schedule)
    {
        await _sync.RemoveAsync(schedule, CancellationToken.None);
        _ui(Reload);
    }
}
