using Avalonia;
using AulaSync.Core;

namespace AulaSync.App;

static class Program
{
    [STAThread]
    public static int Main(string[] args)
    {
        var paths = AppPaths.Default();
        // En fejl, som intet fanger, skrives i loggen, før appen lukker (Indstillinger → Fejlfinding).
        AppDomain.CurrentDomain.UnhandledException += (_, e) => new FileLog(paths.Log).Error("Uventet fejl", e.ExceptionObject as Exception);
        using var instance = SingleInstance.TryAcquire(paths.LockFile);
        if (instance is null)
        {
            // AulaSync kører allerede: bed den kørende instans om at vise hovedvinduet, og afslut.
            InstanceChannel.SignalShowAsync(InstanceChannel.DefaultName, TimeSpan.FromSeconds(2)).GetAwaiter().GetResult();
            return 0;
        }
        App.Silent = args.Contains("--silent");
        return BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder.Configure<App>().UsePlatformDetect().LogToTrace().With(MacOptions);

    // Mac: app-menuen er AulaSyncs egen på dansk (NativeMenus.AppMenu); Avalonias engelske standardpunkter slås fra.
    // Appen starter uden Dock-ikon og får det, når et vindue vises (AppWindows.UpdateDock), så en start ved login
    // (--silent) kun ses i menulinjen.
    internal static MacOSPlatformOptions MacOptions => new() { DisableDefaultApplicationMenuItems = true, ShowInDock = false };
}
