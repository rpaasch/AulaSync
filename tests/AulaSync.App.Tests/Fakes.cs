using AulaSync.Core;

namespace AulaSync.App.Tests;

sealed class FakePlatform(bool isMac = true) : IPlatform
{
    public bool IsMac => isMac;
    public List<string> Opened { get; } = [];
    public List<string> Revealed { get; } = [];
    public List<string> Copied { get; } = [];
    public List<string> OpenedInCalendar { get; } = [];

    public Action<string>? OnOpen { get; set; }
    public bool CopyWorks { get; set; } = true;

    public void Open(string target) { OnOpen?.Invoke(target); Opened.Add(target); }
    public void OpenInCalendar(string url) { OnOpen?.Invoke(url); OpenedInCalendar.Add(url); }
    public void RevealFile(string path) => Revealed.Add(path);
    public Task<bool> CopyTextAsync(string text)
    {
        if (CopyWorks) Copied.Add(text);
        return Task.FromResult(CopyWorks);
    }
}

sealed class FakeAutostart(bool enabled = false) : IAutostart
{
    public bool IsEnabled { get; private set; } = enabled;
    public void SetEnabled(bool enabled) => IsEnabled = enabled;
}

// Et komplet sæt Core-tjenester i en midlertidig mappe.
sealed class TestHost : IDisposable
{
    public static readonly DateTimeOffset Start = new(2026, 10, 6, 12, 0, 0, TimeSpan.Zero);

    // En lektion dagen efter Start. En import tæller kun lektioner fra importen og frem (SyncService.ChangedSinceImport).
    public static AulaEvent Ahead(string id) => TestData.Ev(id: id, start: "2026-10-07T08:00:00+02:00", end: "2026-10-07T08:45:00+02:00");

    public TempDir Dir { get; } = new();
    public Microsoft.Extensions.Time.Testing.FakeTimeProvider Time { get; } = new(Start);
    public AppPaths Paths { get; }
    public FileLog Log { get; }
    public SubscriptionStore Store { get; }
    public ConfigStore Config { get; }
    public SyncService Sync { get; }
    public FakeAulaClient Client { get; } = new();

    public TestHost()
    {
        Paths = new AppPaths(Dir.Path);
        Log = new FileLog(Paths.Log);
        Store = new SubscriptionStore(Paths.Subscriptions);
        Config = new ConfigStore(Paths.Config);
        Sync = new SyncService(Store, Paths.Calendars, Log, Time, (_, _) => Task.CompletedTask);
    }

    public void Dispose() => Dir.Dispose();
}

sealed class FakeDialogs : IDialogs
{
    public int Logins, Settings, AddSchedules;
    public List<ScheduleRef> ImportGuides { get; } = [];
    public List<ScheduleRef> Fallbacks { get; } = [];
    public bool ImportGuideResult { get; set; } = true;
    public bool LogoutResult { get; set; } = true;

    public void ShowLogin() => Logins++;
    public void ShowSettings() => Settings++;
    public Task ShowAddScheduleAsync() { AddSchedules++; return Task.CompletedTask; }
    public Task<bool> ShowImportGuideAsync(ScheduleRef schedule) { ImportGuides.Add(schedule); return Task.FromResult(ImportGuideResult); }
    public Task ShowOutlookFallbackAsync(ScheduleRef schedule) { Fallbacks.Add(schedule); return Task.CompletedTask; }
    public Task<bool> ConfirmLogoutAsync() => Task.FromResult(LogoutResult);
}

sealed class FakeActions : IMainActions
{
    public List<string> Calls { get; } = [];
    public Task RunAsync(ScheduleRef schedule) { Calls.Add($"run {schedule.Key}"); return Task.CompletedTask; }
    public Task CopyAddressAsync(ScheduleRef schedule) { Calls.Add($"copy {schedule.Key}"); return Task.CompletedTask; }
    public void RevealFile(ScheduleRef schedule) => Calls.Add($"reveal {schedule.Key}");
}

sealed class FakeDirectory : IAulaDirectory
{
    public int Calls;
    public bool Fail;

    public Task<IReadOnlyList<Employee>> GetEmployeesAsync(CancellationToken ct)
    {
        Calls++;
        if (Fail) throw new AulaException("nede");
        return Task.FromResult<IReadOnlyList<Employee>>([new("1001", "Anna Eksempel", "AE", "teacher"), new("1002", "Bo Testesen", "BT", "teacher")]);
    }

    public Task<IReadOnlyList<NamedItem>> GetGroupsAsync(CancellationToken ct) =>
        Task.FromResult<IReadOnlyList<NamedItem>>([new("10", "10A"), new("7", "7A"), new("1", "1A")]);

    public Task<IReadOnlyList<NamedItem>> GetResourcesAsync(CancellationToken ct) =>
        Task.FromResult<IReadOnlyList<NamedItem>>([new("412", "Lokale 53"), new("413", "Gymnastiksal")]);
}

sealed class FakeNotifier : INotifier
{
    public List<string> Shown { get; } = [];
    public event Action? Clicked;
    public event Action? Dismissed;
    Action? _onClick;
    public void Show(string text, Action? onClick = null) { Shown.Add(text); _onClick = onClick; }
    public void Click() { if (_onClick is { } onClick) onClick(); else Clicked?.Invoke(); }
    public void Dismiss() => Dismissed?.Invoke();
}
