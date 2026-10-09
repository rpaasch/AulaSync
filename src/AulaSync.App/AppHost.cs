using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Threading;
using AulaSync.Core;

namespace AulaSync.App;

// Programmets rod: opretter tjenesterne, starter server, baggrundsplan og ikon, og bestemmer hvad der vises ved start.
public sealed class AppHost
{
    readonly Application _app;
    readonly IClassicDesktopStyleApplicationLifetime _desktop;
    readonly bool _silent;
    readonly CancellationTokenSource _cts = new();
    readonly InstanceChannel _channel = new(InstanceChannel.DefaultName);
    // Ikonet har én menu hele tiden; UpdateTray udskifter punkterne (se NativeMenus).
    readonly NativeMenu _trayMenu = new();
    readonly TrayIcon _tray;
    IcsServer? _server;

    // Sæt til false, hvis spiken viste, at stille genlogin ikke virker (punkt 5c); så er genlogin altid et klik.
    const bool SilentReloginSupported = true;

    public AppHost(Application app, IClassicDesktopStyleApplicationLifetime desktop, AppPaths paths, bool silent)
    {
        _app = app;
        _desktop = desktop;
        _silent = silent;
        _tray = new TrayIcon { ToolTipText = "AulaSync", Menu = _trayMenu, IsVisible = true };
        Paths = paths;
        Log = new FileLog(paths.Log);
        Config = new ConfigStore(paths.Config);
        Sync = new SyncService(new SubscriptionStore(paths.Subscriptions), paths.Calendars, Log, Time, config: Config.Load);
        Scheduler = new BackgroundScheduler(Sync, Time, Log, () => Config.Load().UpdateInterval);
        Session = new SessionController(Sync, Log);
        Browser = new BrowserProfile(paths, OperatingSystem.IsMacOS());
        Platform = new DesktopPlatform(Log);
        Autostart = AutostartFactory.ForCurrentPlatform(ExecutablePath);
        Windows = new AppWindows(this);
        Actions = new MainActions(Sync, Config, Platform, Windows);
        Notifier = new NotificationBox(Platform, Log, Time);
        // Uden WebView2 Runtime kan login ikke lykkes; så vises notifikationen med det samme i stedet for efter 30 s.
        Relogin = new Relogin(ct => SilentLogin.RunAsync(Browser, Session, ct), Notifier, Log, Time,
            SilentReloginSupported && !BrowserProfile.WebView2Missing);
        ImportReminder = new ImportReminder(Sync, Config, Platform, Notifier, () => Windows.ShowMain(), a => Dispatcher.UIThread.Post(a));
        Uninstaller = new UninstallLauncher(ExecutablePath, OperatingSystem.IsWindows(), OperatingSystem.IsMacOS(), Autostart, Quit,
            text => Notifier.Show(text, () => { }), UninstallLauncher.StartProcess, Uninstall.RegisteredSetupDir);
    }

    // Den exe, der kører; startet gennem winget's henvisning (Links\AulaSync.exe) selve filen, så start ved login og genvejen
    // i Start-menuen peger samme sted hen.
    static string ExecutablePath { get; } = Environment.ProcessPath is { } path ? StartMenuShortcut.RealPath(path) : "AulaSync";

    public AppPaths Paths { get; }
    public FileLog Log { get; }
    public ConfigStore Config { get; }
    public SyncService Sync { get; }
    public BackgroundScheduler Scheduler { get; }
    public SessionController Session { get; }
    public BrowserProfile Browser { get; }
    public IPlatform Platform { get; }
    public IAutostart Autostart { get; }
    public AppWindows Windows { get; }
    public IMainActions Actions { get; }
    public INotifier Notifier { get; }
    public UninstallLauncher Uninstaller { get; }
    public Relogin Relogin { get; }
    public ImportReminder ImportReminder { get; }
    public TimeProvider Time { get; } = TimeProvider.System;
    public TimeZoneInfo TimeZone { get; } = TimeZoneInfo.Local;
    public bool ServerRunning { get; private set; }
    public bool IsQuitting { get; private set; }

    public void Start()
    {
        Log.Info($"AulaSync {SettingsViewModel.Version} starter{(_silent ? " (--silent)" : "")}");
        if (OperatingSystem.IsWindows() && OldVersion.IsRunning()) Notifier.Show(OldVersion.Text, () => { });
        _ = Task.Run(UpdateStartMenuShortcut);
        // Start ved login følger med, når AulaSync er flyttet eller installeret et nyt sted (fx den løse AulaSync.exe fra før
        // installationsprogrammet).
        try { if (Autostart.Retarget()) Log.Info("Start ved login peger nu på denne AulaSync"); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            Log.Error("Kunne ikke rette start ved login", ex);
        }
        ServerRunning = StartServer();
        _ = Scheduler.RunAsync(_cts.Token).ContinueWith(t => Log.Error("Baggrundsplanen stoppede", t.Exception!),
            TaskContinuationOptions.OnlyOnFaulted);
        _channel.Listen(() => Dispatcher.UIThread.Post(Windows.ShowMainOrOnboarding), () => Dispatcher.UIThread.Post(Quit));

        // Sync ved login; ikon og menu følger status.
        Session.Changed += () =>
        {
            if (!Sync.Status.LoggedIn) return;
            Relogin.SignedIn();
            _ = Sync.SyncAllAsync(_cts.Token);
        };
        // Udløbet session: prøv stille genlogin; ellers notifikationsboksen én gang, og ikonet viser "Log ind igen…".
        // Klik på boksen åbner login (eller første start, hvis den ikke er gennemført).
        Sync.SessionExpired += () => Dispatcher.UIThread.Post(() => _ = Relogin.TryAsync(Relogin.ExpiredText));
        Notifier.Clicked += Windows.ShowLogin;
        ImportReminder.Start();
        // Mac: ✕ på boksen giver fokus tilbage til det program, man arbejdede i (medmindre et AulaSync-vindue er åbent).
        Notifier.Dismissed += () => { if (!Windows.AnyVisible) MacDock.HideApp(); };
        Sync.StatusChanged += _ => Dispatcher.UIThread.Post(UpdateTray);
        BuildTray();
        _desktop.ShutdownRequested += (_, _) => Stop();
        // Mac: klik på Dock-ikonet eller ny start fra Finder viser hovedvinduet.
        if (_app.TryGetFeature<IActivatableLifetime>() is { } activatable)
            activatable.Activated += (_, e) => { if (e.Kind == ActivationKind.Reopen) Windows.ShowMainOrOnboarding(); };

        _ = ShowAtStartAsync();
    }

    // Uden gennemført første start: første start. Ellers stille login med den gemte profil; lykkes det ikke, vises login.
    // --silent (autostart): intet vindue; mangler login, sættes ikonet til "kræver handling", og notifikationsboksen vises én gang.
    async Task ShowAtStartAsync()
    {
        if (!Config.Load().FirstRunDone)
        {
            if (_silent) Notifier.Show(Relogin.StartText);
            else Windows.ShowOnboarding();
            return;
        }
        if (!_silent) Windows.ShowMain();
        if (!await Relogin.TryAsync(_silent ? Relogin.StartText : null) && !_silent) Windows.ShowLogin();
    }

    // Windows: genvejen AulaSync i Start-menuen peger på den exe, der kører (StartMenuShortcut). Lykkes det ikke, kører
    // AulaSync videre; genvejen er kun en hjælp til at finde den.
    void UpdateStartMenuShortcut()
    {
        if (!OperatingSystem.IsWindows() || Environment.ProcessPath is null) return;
        try
        {
            if (StartMenuShortcut.Ensure(StartMenuShortcut.DefaultPath, ExecutablePath))
                Log.Info("Genvejen AulaSync i Start-menuen er lavet eller rettet");
        }
        catch (Exception ex) { Log.Error("Kunne ikke lave genvejen AulaSync i Start-menuen", ex); }
    }

    // Brugerens egen port (UserPort): første gang 9876 eller den første ledige.
    bool StartServer()
    {
        _server?.Dispose();
        Directory.CreateDirectory(Paths.Calendars);
        _server = UserPort.Start(Config, port =>
        {
            var server = new IcsServer(Paths.Calendars, Log, port);
            server.Fetched += (name, agent) => _ = RecordFetchAsync(name, agent);
            if (server.TryStart()) return server;
            server.Dispose();
            return null;
        });
        return _server is not null;
    }

    // Kalenderprogrammet hentede et skema: rækken skifter fra "Venter på Kalender…" til "✓ Tilføjet", og AulaSync kan se,
    // om et skema ikke længere bliver hentet (FetchWatch).
    async Task RecordFetchAsync(string fileName, string userAgent)
    {
        try { await Sync.MarkFetchedAsync(fileName, userAgent); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Log.Error($"Kunne ikke gemme, at {fileName} blev hentet", ex);
        }
    }

    public bool RetryServer()
    {
        ServerRunning = StartServer();
        UpdateTray();
        return ServerRunning;
    }

    public async Task SignOutAsync()
    {
        await Session.SignOutAsync(Browser);
        Config.Save(Config.Load() with { FirstRunDone = false });
        Windows.CloseAll();
        Windows.ShowOnboarding();
    }

    void BuildTray()
    {
        if (OperatingSystem.IsMacOS()) MacOSProperties.SetIsTemplateIcon(_tray, true);
        // Windows: venstreklik åbner hovedvinduet, højreklik åbner menuen. Mac: klik åbner menuen.
        if (!OperatingSystem.IsMacOS()) _tray.Clicked += (_, _) => Windows.ShowMainOrOnboarding();
        TrayIcon.SetIcons(_app, [_tray]);
        UpdateTray();
    }

    public void UpdateTray()
    {
        var status = Sync.Status;
        var changed = ImportReminder.Pending();
        _tray.Icon = TrayIconImage.Create(TrayMenu.IconState(status, ServerRunning, changed.Count > 0), template: OperatingSystem.IsMacOS());
        NativeMenus.Fill(_trayMenu, TrayMenu.Build(status, ServerRunning, TimeZone, changed), Run, OperatingSystem.IsMacOS());
    }

    void Run(TrayCommand command)
    {
        switch (command)
        {
            case TrayCommand.LogIn: Windows.ShowLogin(); break;
            case TrayCommand.RetryServer: RetryServer(); break;
            case TrayCommand.Open: Windows.ShowMainOrOnboarding(); break;
            case TrayCommand.Refresh: _ = Sync.SyncAllAsync(_cts.Token); break;
            case TrayCommand.Settings: Windows.ShowSettings(); break;
            case TrayCommand.Quit: Quit(); break;
        }
    }

    public void Quit()
    {
        Stop();
        _desktop.Shutdown();
    }

    // Afslut fra menuen, ⌘Q, Dock eller fordi styresystemet logger ud: ryd op én gang.
    void Stop()
    {
        if (IsQuitting) return;
        IsQuitting = true;
        Log.Info("AulaSync afsluttes");
        _cts.Cancel();
        _server?.Dispose();
        _channel.Dispose();
        _tray.IsVisible = false;
    }
}
