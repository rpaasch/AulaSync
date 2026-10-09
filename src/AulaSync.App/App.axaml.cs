using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Platform;
using Avalonia.Styling;
using Avalonia.Threading;
using AulaSync.Core;

namespace AulaSync.App;

public partial class App : Application
{
    // --silent (autostart): start uden vindue, kun ikonet.
    public static bool Silent { get; set; }

    // AulaSyncs ikon (packaging/icon). Windows: på alle vinduer og i proceslinjen; Mac bruger ikonet i AulaSync.app.
    public static readonly Uri IconUri = new("avares://AulaSync/Assets/AulaSync.ico");

    AppHost? _host;

    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
        if (OperatingSystem.IsWindows())
            Styles.Add(new Style(x => x.Is<Window>()) { Setters = { new Setter(Window.IconProperty, new WindowIcon(AssetLoader.Open(IconUri))) } });
        // Mac: app-menuen ("Indstillinger… ⌘,") skal sættes her. Avalonia læser den, før OnFrameworkInitializationCompleted,
        // og ignorerer en app-menu, der sættes senere.
        if (OperatingSystem.IsMacOS()) NativeMenu.SetMenu(this, NativeMenus.AppMenu(() => _host?.Windows.ShowSettings(), () => _host?.Quit()));
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.ShutdownMode = ShutdownMode.OnExplicitShutdown; // lukkede vinduer afslutter ikke appen
            var host = _host = new AppHost(this, desktop, AppPaths.Default(), Silent);
            // Fejl fra login-webviewet (fx et login-vindue, der lukkes, mens WebView2 starter) må ikke lukke appen.
            Dispatcher.UIThread.UnhandledException += (_, e) => { if (UiErrors.TryHandle(e.Exception, host.Log)) e.Handled = true; };
            host.Start();
        }
        base.OnFrameworkInitializationCompleted();
    }
}
