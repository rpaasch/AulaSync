using Avalonia.Controls;
using Avalonia.Threading;
using AulaSync.Core;

namespace AulaSync.App;

// Alle vinduer og dialoger. Hovedvinduet, Indstillinger, login og første start findes højst én gang ad gangen.
public sealed class AppWindows(AppHost host) : IDialogs
{
    MainWindow? _main;
    MainViewModel? _mainVm;
    SettingsWindow? _settings;
    Window? _login;
    OnboardingWindow? _onboarding;

    static void Post(Action action) => Dispatcher.UIThread.Post(action);

    public MainViewModel MainViewModel => _mainVm ??= NewMainViewModel();

    MainViewModel NewMainViewModel() =>
        new(host.Sync, host.Config, host.Platform, this, host.Actions, host.Time, host.TimeZone,
            () => host.ServerRunning, host.RetryServer, Post);

    AddScheduleViewModel NewAddScheduleViewModel() =>
        new(() => host.Session.Catalog, () => host.Session.OwnSchedule, host.Sync, host.Log, Post);

    public void ShowMainOrOnboarding()
    {
        if (host.Config.Load().FirstRunDone) ShowMain();
        else ShowOnboarding();
    }

    public void ShowMain()
    {
        if (_main is null)
        {
            _main = new MainWindow(MainViewModel);
            WindowChrome.Apply(_main);
            // At lukke hovedvinduet skjuler det; appen kører videre i menulinjen/systembakken.
            _main.Closing += (_, e) =>
            {
                if (!HidesInsteadOfClosing(e.CloseReason, host.IsQuitting)) return;
                e.Cancel = true;
                _main.Hide();
                UpdateDock();
            };
        }
        MainViewModel.Reload();
        Present(_main);
    }

    // Vises første start igen fra Indstillinger, lukkes Indstillinger, så dialoger i trin 5 hører til første start, og
    // kortene ikke viser et gammelt valg bagefter.
    public void ShowOnboarding()
    {
        _settings?.Close();
        if (_onboarding is null)
        {
            var login = new LoginViewModel(host.Session);
            bool SignedIn() => host.Session.Profile is not null && host.Sync.Status.LoggedIn;
            var vm = new OnboardingViewModel(host.Config, host.Platform, host.Autostart, NewAddScheduleViewModel(), MainViewModel,
                () => host.Session.OwnSchedule, SignedIn);
            login.SignedIn += _ => vm.SignedIn();
            // Første start, der blev afbrudt efter login, fortsætter ved kalenderprogrammet. Vist igen fra Indstillinger
            // begynder den ved velkomsten.
            if (SignedIn() && !host.Config.Load().FirstRunDone) vm.SignedIn();
            _onboarding = new OnboardingWindow(vm, new LoginView(host.Browser, login));
            WindowChrome.Apply(_onboarding);
            WindowFit.ToScreen(_onboarding);
            _onboarding.Closed += (_, _) =>
            {
                _onboarding = null;
                if (host.Config.Load().FirstRunDone) MainViewModel.Reload();
                host.UpdateTray(); // prikken følger kalenderprogrammet, der kan være valgt om
                Post(UpdateDock);
            };
        }
        Present(_onboarding);
    }

    // "Log ind igen": login-webview i et lille vindue. Før første start er gennemført, er det første start, der vises.
    public void ShowLogin()
    {
        if (!host.Config.Load().FirstRunDone) { ShowOnboarding(); return; }
        if (_login is null)
        {
            var vm = new LoginViewModel(host.Session);
            _login = new Window { Title = "Log ind på Aula", Width = 900, Height = 760, WindowStartupLocation = WindowStartupLocation.CenterScreen };
            _login.Content = new LoginView(host.Browser, vm) { Margin = new Avalonia.Thickness(12) };
            WindowFit.ToScreen(_login);
            vm.SignedIn += _ => _login?.Close();
            _login.Closed += (_, _) => { _login = null; Post(UpdateDock); };
        }
        Present(_login);
    }

    public void ShowSettings()
    {
        if (_settings is null)
        {
            var vm = new SettingsViewModel(host.Config, host.Autostart, host.Session, this, host.Platform, host.Paths,
                host.SignOutAsync, () => { MainViewModel.Reload(); host.UpdateTray(); }, // prikken følger kalenderprogrammet
                host.Scheduler.Reschedule);
            _settings = new SettingsWindow(vm);
            WindowChrome.Apply(_settings);
            WindowFit.CapToScreen(_settings);
            _settings.Closed += (_, _) => { _settings = null; Post(UpdateDock); };
        }
        Present(_settings);
    }

    public Task ShowAddScheduleAsync() => ShowModalAsync<object?>(new AddScheduleWindow(NewAddScheduleViewModel()));

    public Task<bool> ShowImportGuideAsync(ScheduleRef schedule) =>
        ShowModalAsync<bool>(new ImportGuideWindow(new ImportGuideViewModel(schedule, () => host.Sync.FilePath(schedule), host.Platform, host.Time)));

    public Task ShowOutlookFallbackAsync(ScheduleRef schedule) =>
        ShowModalAsync<object?>(new OutlookFallbackWindow(new OutlookFallbackViewModel(schedule, host.Config.Load().ServerPort, host.Platform, host.Time)));

    public Task<bool> ConfirmLogoutAsync() => ShowModalAsync<bool>(new ConfirmLogoutWindow());

    // Log ud: luk alt; første start vises igen, fra begyndelsen (også hvis den var vist igen fra Indstillinger).
    public void CloseAll()
    {
        _settings?.Close();
        _login?.Close();
        _onboarding?.Close();
        _onboarding = null;
        _main?.Hide();
        _mainVm?.Reload();
        UpdateDock();
    }

    async Task<T> ShowModalAsync<T>(Window dialog)
    {
        var owner = new Window?[] { _settings, _onboarding, _main }.FirstOrDefault(w => w is { IsVisible: true });
        if (owner is not null) return await dialog.ShowDialog<T>(owner);
        var closed = new TaskCompletionSource<T>();
        dialog.Closed += (_, _) => closed.TrySetResult(default!);
        Present(dialog);
        return await closed.Task;
    }

    void Present(Window window)
    {
        ShowRestored(window);
        UpdateDock();
    }

    // Show og Activate gør intet ved et minimeret vindue på Windows; det skal gendannes først.
    internal static void ShowRestored(Window window)
    {
        window.Show();
        if (window.WindowState == WindowState.Minimized) window.WindowState = WindowState.Normal;
        window.Activate();
    }

    // Mac: Dock-ikon, mens et vindue er åbent; ellers kun menulinjen.
    // Kun når brugeren selv lukker vinduet. Afslut (⌘Q, Dock) og log ud/genstart af styresystemet skal kunne lukke det,
    // ellers afvises de.
    public static bool HidesInsteadOfClosing(WindowCloseReason reason, bool quitting) =>
        !quitting && reason == WindowCloseReason.WindowClosing;

    // Er et af AulaSyncs vinduer synligt (notifikationsboksen tæller ikke med)?
    public bool AnyVisible => new Window?[] { _main, _settings, _login, _onboarding }.Any(w => w is { IsVisible: true });

    public void UpdateDock() => MacDock.SetVisible(AnyVisible);
}
