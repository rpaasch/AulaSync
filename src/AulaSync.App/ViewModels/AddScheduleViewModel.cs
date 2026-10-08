using System.Collections.ObjectModel;
using AulaSync.Core;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace AulaSync.App;

public sealed record HeaderRow(string Title);

public sealed partial class CatalogEntryViewModel(ScheduleRef schedule, bool suggested, bool added, Action<CatalogEntryViewModel> add) : ObservableObject
{
    public ScheduleRef Schedule { get; } = schedule;
    public string Name => Schedule.CalendarName;
    // Første/sidste række under en overskrift: kortet får runde hjørner øverst/nederst.
    public bool IsFirst { get; init; }
    public bool IsLast { get; init; }
    public string Sub => suggested ? "Dit eget skema" : StatusText.KindName(Schedule);
    public string BadgeText => Badge.TextFor(Schedule);
    public bool IsEmployee => Schedule.Kind == ScheduleKind.Employee;
    public bool IsGroup => Schedule.Kind == ScheduleKind.Group;
    public bool IsResource => Schedule.Kind == ScheduleKind.Resource;
    public string AddLabel => $"Tilføj {Name}";

    [ObservableProperty] public partial bool IsAdded { get; set; } = added;

    [RelayCommand] void Add() => add(this);
}

// "Tilføj skema" (spec §3.2): ét søgefelt, filteret Alle · Medarbejdere · Klasser · Lokaler og en grupperet liste.
// Bruges også i første start, trin 4.
public sealed partial class AddScheduleViewModel : ObservableObject
{
    readonly Func<ScheduleCatalog?> _catalog;
    readonly Func<ScheduleRef?> _ownSchedule;
    readonly SyncService _sync;
    readonly FileLog _log;
    readonly Action<Action> _ui;
    readonly HashSet<string> _subscribed = [];
    IReadOnlyList<ScheduleRef> _all = [];

    public AddScheduleViewModel(Func<ScheduleCatalog?> catalog, Func<ScheduleRef?> ownSchedule, SyncService sync, FileLog log, Action<Action> ui)
    {
        _catalog = catalog;
        _ownSchedule = ownSchedule;
        _sync = sync;
        _log = log;
        _ui = ui;
    }

    public ObservableCollection<object> Rows { get; } = [];

    [ObservableProperty] public partial string Query { get; set; } = "";
    [ObservableProperty] public partial CatalogFilter Filter { get; set; } = CatalogFilter.All;
    [ObservableProperty] public partial bool IsLoading { get; set; }
    [ObservableProperty] public partial string? Error { get; set; }
    [ObservableProperty] public partial string? EmptyText { get; set; }
    [ObservableProperty] public partial string CountText { get; set; } = "";

    public bool HasError => Error is not null;
    public bool HasEmptyText => EmptyText is not null;

    public bool IsAll { get => Filter == CatalogFilter.All; set { if (value) Filter = CatalogFilter.All; } }
    public bool IsEmployees { get => Filter == CatalogFilter.Employees; set { if (value) Filter = CatalogFilter.Employees; } }
    public bool IsGroups { get => Filter == CatalogFilter.Groups; set { if (value) Filter = CatalogFilter.Groups; } }
    public bool IsResources { get => Filter == CatalogFilter.Resources; set { if (value) Filter = CatalogFilter.Resources; } }

    partial void OnQueryChanged(string value) => Rebuild();
    partial void OnErrorChanged(string? value) => OnPropertyChanged(nameof(HasError));
    partial void OnEmptyTextChanged(string? value) => OnPropertyChanged(nameof(HasEmptyText));

    partial void OnFilterChanged(CatalogFilter value)
    {
        OnPropertyChanged(nameof(IsAll));
        OnPropertyChanged(nameof(IsEmployees));
        OnPropertyChanged(nameof(IsGroups));
        OnPropertyChanged(nameof(IsResources));
        Rebuild();
    }

    // Henter listerne første gang efter login; derefter kommer de fra cachen i ScheduleCatalog.
    [RelayCommand]
    public async Task LoadAsync()
    {
        _subscribed.Clear();
        _subscribed.UnionWith(_sync.Subscriptions.Select(s => s.Key));
        if (_catalog() is not { } catalog)
        {
            Error = "Du er ikke logget ind i Aula.";
            Rebuild();
            return;
        }
        IsLoading = true;
        Error = null;
        try { _all = await catalog.GetAllAsync(); }
        catch (Exception ex) when (ex is AulaException or HttpRequestException or TaskCanceledException)
        {
            _log.Error("Listen over skemaer kunne ikke hentes", ex);
            Error = $"Listen kunne ikke hentes: {ex.Message}";
        }
        finally { IsLoading = false; }
        Rebuild();
    }

    void Rebuild()
    {
        Rows.Clear();
        var rows = ScheduleSearch.BuildRows(_all, Query, Filter, _ownSchedule(), _subscribed);
        for (int i = 0; i < rows.Count; i++)
            Rows.Add(rows[i] switch
            {
                CatalogHeader h => new HeaderRow(h.Title),
                CatalogEntry e => new CatalogEntryViewModel(e.Schedule, e.Suggested, e.Subscribed, Add)
                {
                    IsFirst = i == 0 || rows[i - 1] is CatalogHeader,
                    IsLast = i == rows.Count - 1 || rows[i + 1] is CatalogHeader,
                },
                _ => throw new InvalidOperationException(),
            });
        var q = Query.Trim();
        EmptyText = rows.Count == 0 && _all.Count > 0 ? $"Ingen resultater for \"{q}\"." : null;
        UpdateCount();
    }

    public int SelectedCount => _subscribed.Count;

    // Tilføj med det samme; hentningen fortsætter i baggrunden, og rækken i hovedvinduet viser "henter…".
    public void Select(ScheduleRef schedule)
    {
        if (!_subscribed.Add(schedule.Key)) return;
        foreach (var e in Rows.OfType<CatalogEntryViewModel>().Where(e => e.Schedule.Key == schedule.Key)) e.IsAdded = true;
        UpdateCount();
        _ = AddInBackgroundAsync(schedule);
    }

    void Add(CatalogEntryViewModel entry) => Select(entry.Schedule);

    void UpdateCount()
    {
        CountText = _subscribed.Count == 1 ? "1 skema valgt" : $"{_subscribed.Count} skemaer valgt";
        OnPropertyChanged(nameof(SelectedCount));
    }

    async Task AddInBackgroundAsync(ScheduleRef schedule)
    {
        try { await _sync.AddAsync(schedule, CancellationToken.None); }
        catch (Exception ex) { _log.Error($"Kunne ikke tilføje {schedule.FileName}", ex); }
    }

}
