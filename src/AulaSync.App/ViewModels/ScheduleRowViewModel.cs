using AulaSync.Core;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace AulaSync.App;

// Én række i "Mine skemaer". Opdateres på stedet, så en åben menu eller bekræftelse ikke forsvinder ved hver statusændring.
public sealed partial class ScheduleRowViewModel : ObservableObject
{
    readonly IMainActions _actions;
    readonly Func<ScheduleRef, Task> _remove;

    public ScheduleRowViewModel(ScheduleRef schedule, IMainActions actions, Func<ScheduleRef, Task> remove, string fileManager)
    {
        Schedule = schedule;
        _actions = actions;
        _remove = remove;
        RevealLabel = $"Vis fil i {fileManager}";
    }

    public ScheduleRef Schedule { get; }
    public string Key => Schedule.Key;
    public string Name => Schedule.CalendarName;
    public string BadgeText => Badge.TextFor(Schedule);
    public bool IsEmployee => Schedule.Kind == ScheduleKind.Employee;
    public bool IsGroup => Schedule.Kind == ScheduleKind.Group;
    public bool IsResource => Schedule.Kind == ScheduleKind.Resource;
    public string MoreLabel => $"Flere handlinger for {Name}";
    public string RevealLabel { get; }
    public string ConfirmText => $"Fjern {Name}? Kalenderen holder op med at blive opdateret. Slet den også i dit kalenderprogram.";

    [ObservableProperty] public partial string SubText { get; set; } = "";
    [ObservableProperty] public partial bool SubIsError { get; set; }
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(HasPrimary))] public partial string? PrimaryLabel { get; set; }
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(HasDone))] public partial string? DoneText { get; set; }
    [ObservableProperty] public partial bool IsWaiting { get; set; }
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(HasNote))] public partial string? Note { get; set; }
    [ObservableProperty] public partial string AgainLabel { get; set; } = "Tilføj igen";
    [ObservableProperty] public partial bool IsConfirmingRemove { get; set; }

    public bool HasPrimary => PrimaryLabel is not null;
    public bool HasDone => DoneText is not null;
    public bool HasNote => Note is not null;

    public void Update(RowText text, RowButtons buttons)
    {
        SubText = text.Text;
        SubIsError = text.IsError;
        // En række, der ikke kunne hentes, har ingen hovedknap (mockup: "Lokale 53"); "…"-menuen er der stadig.
        PrimaryLabel = text.IsError ? null : buttons.Primary;
        DoneText = text.IsError ? null : buttons.Done;
        IsWaiting = buttons.Waiting;
        Note = text.IsError ? null : buttons.Note;
        AgainLabel = buttons.Again;
    }

    [RelayCommand] Task Primary() => _actions.RunAsync(Schedule);
    [RelayCommand] Task Again() => _actions.RunAsync(Schedule);
    [RelayCommand] Task CopyAddress() => _actions.CopyAddressAsync(Schedule);
    [RelayCommand] void Reveal() => _actions.RevealFile(Schedule);
    [RelayCommand] void AskRemove() => IsConfirmingRemove = true;
    [RelayCommand] void CancelRemove() => IsConfirmingRemove = false;
    [RelayCommand] Task ConfirmRemove() => _remove(Schedule);
}
