using AulaSync.Core;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace AulaSync.App;

public enum OnboardingStep { Welcome = 1, Login, CalendarApp, Schedules, AddToCalendar }

// Et valgkort for kalenderprogram: navn, "● Opdateres automatisk" / "○ Øjebliksbillede" og en kort forklaring.
public sealed partial class CalendarCardViewModel(CalendarApp app, Action<CalendarCardViewModel> select) : ObservableObject
{
    public CalendarApp App { get; } = app;
    public string Title => CalendarApps.Title(App);
    public bool IsAutomatic => CalendarApps.UpdatesAutomatically(App);
    public string Mode => (IsAutomatic ? "● " : "○ ") + CalendarApps.ModeLabel(App);
    public string Description => CalendarApps.Description(App);

    [ObservableProperty] public partial bool IsSelected { get; set; }

    [RelayCommand] void Select() => select(this);
}

// De fire kort med kalenderprogrammer (CalendarCardsView); bruges ved første start og i Indstillinger.
public interface ICalendarCards
{
    IReadOnlyList<CalendarCardViewModel> Cards { get; }
}

// Første start: ét vindue, fem trin (spec §3.2).
public sealed partial class OnboardingViewModel : ObservableObject, ICalendarCards
{
    readonly ConfigStore _config;
    readonly IPlatform _platform;
    readonly IAutostart _autostart;
    readonly Func<ScheduleRef?> _ownSchedule;
    bool _ownPreselected;

    public OnboardingViewModel(ConfigStore config, IPlatform platform, IAutostart autostart,
        AddScheduleViewModel add, MainViewModel main, Func<ScheduleRef?> ownSchedule)
    {
        _config = config;
        _platform = platform;
        _autostart = autostart;
        _ownSchedule = ownSchedule;
        Add = add;
        Main = main;
        var chosen = config.Load().CalendarApp ?? CalendarApps.DefaultFor(platform.IsMac);
        Cards = CalendarApps.All.Select(a => new CalendarCardViewModel(a, SelectCard) { IsSelected = a == chosen }).ToList();
        add.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(AddScheduleViewModel.SelectedCount)) OnPropertyChanged(nameof(CanContinue)); };
    }

    public AddScheduleViewModel Add { get; }
    public MainViewModel Main { get; }
    public IReadOnlyList<CalendarCardViewModel> Cards { get; }

    public string WelcomeText => "AulaSync henter skemaer fra Aula og holder dem opdateret i din kalender. Den taler kun med Aula og gemmer alt på din computer.";
    public bool ShowOutlookHint => !_platform.IsMac;

    // AulaSync serverer kalenderne fra denne computer (spec §2): kalenderprogrammet skal køre samme sted.
    public const string SameComputerText =
        "AulaSync og kalenderprogrammet skal køre på samme computer. En kalender på nettet eller på telefonen kan ikke selv hente skemaerne.";

    // Apple Kalender: med placeringen iCloud henter Apples servere filen, og de kan ikke nå denne computer.
    public bool ShowAppleCalendarHint => Cards.Any(c => c.IsSelected && c.App == CalendarApp.AppleCalendar);
    // Står i trin 5 (ingen notifikation ved Færdig).
    public string FinishMessage => _platform.IsMac
        ? "Når du klikker Færdig, kører AulaSync videre i menulinjen."
        : "Når du klikker Færdig, kører AulaSync videre i systembakken. Træk ikonet ned på proceslinjen for at have det ved hånden.";

    [ObservableProperty] public partial OnboardingStep Step { get; set; } = OnboardingStep.Welcome;

    public bool IsWelcome => Step == OnboardingStep.Welcome;
    public bool IsLogin => Step == OnboardingStep.Login;
    public bool IsCalendarApp => Step == OnboardingStep.CalendarApp;
    public bool IsSchedules => Step == OnboardingStep.Schedules;
    public bool IsAddToCalendar => Step == OnboardingStep.AddToCalendar;
    public bool ShowFooter => Step >= OnboardingStep.CalendarApp;
    public bool CanGoBack => Step > OnboardingStep.CalendarApp;
    public string ContinueLabel => Step == OnboardingStep.AddToCalendar ? "Færdig" : "Fortsæt";
    public bool CanContinue => Step switch
    {
        OnboardingStep.CalendarApp => Cards.Any(c => c.IsSelected),
        OnboardingStep.Schedules => Add.SelectedCount > 0,
        _ => true,
    };

    // Fem streger øverst; de første <trin> er fremhævet.
    public IReadOnlyList<bool> Progress => Enumerable.Range(1, 5).Select(i => i <= (int)Step).ToList();

    public event Action? Finished;

    partial void OnStepChanged(OnboardingStep value)
    {
        foreach (var name in new[] { nameof(IsWelcome), nameof(IsLogin), nameof(IsCalendarApp), nameof(IsSchedules), nameof(IsAddToCalendar),
                     nameof(ShowFooter), nameof(CanGoBack), nameof(ContinueLabel), nameof(CanContinue), nameof(Progress) })
            OnPropertyChanged(name);
    }

    [RelayCommand] void LogIn() => Step = OnboardingStep.Login;

    // Kaldes, når login i trin 2 er gennemført; appen går selv videre.
    public void SignedIn() => Step = OnboardingStep.CalendarApp;

    [RelayCommand]
    async Task Continue()
    {
        switch (Step)
        {
            case OnboardingStep.CalendarApp:
                var app = Cards.Single(c => c.IsSelected).App;
                _config.Save(_config.Load() with { CalendarApp = app });
                Step = OnboardingStep.Schedules;
                await Add.LoadAsync();
                PreselectOwnSchedule();
                break;
            case OnboardingStep.Schedules:
                Step = OnboardingStep.AddToCalendar;
                Main.Reload();
                break;
            case OnboardingStep.AddToCalendar:
                Finish();
                break;
        }
    }

    [RelayCommand]
    void Back()
    {
        if (CanGoBack) Step--;
    }

    void SelectCard(CalendarCardViewModel card)
    {
        foreach (var c in Cards) c.IsSelected = c == card;
        OnPropertyChanged(nameof(CanContinue));
        OnPropertyChanged(nameof(ShowAppleCalendarHint));
    }

    // Trin 4: dit eget skema er forudvalgt (én gang; fjerner brugeren det senere, kommer det ikke igen).
    void PreselectOwnSchedule()
    {
        if (_ownPreselected || _ownSchedule() is not { } own) return;
        _ownPreselected = true;
        Add.Select(own);
    }

    void Finish()
    {
        _config.Save(_config.Load() with { FirstRunDone = true });
        _autostart.SetEnabled(true); // "Start AulaSync, når jeg logger ind" er til som standard
        Finished?.Invoke();
    }
}
