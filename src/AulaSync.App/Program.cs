using Avalonia;
using AulaSync.Core;

namespace AulaSync.App;

static class Program
{
    [STAThread]
    public static int Main(string[] args)
    {
        var paths = AppPaths.Default();
        // Installationsprogrammet: afslut en kørende AulaSync, før filerne skiftes ud eller fjernes.
        if (args.Contains("--quit")) return QuitRunning(paths.LockFile, InstanceChannel.DefaultName, TimeSpan.FromSeconds(10));
        // "Afinstallér AulaSync…" uden installationsprogram (Mac, løs exe): slet alt, når den kørende AulaSync er afsluttet.
        if (args.Contains(Uninstall.Argument))
        {
            // Den kørende AulaSync afslutter selv; ellers bedes den om det, ligesom med --quit.
            if (QuitRunning(paths.LockFile, InstanceChannel.DefaultName, TimeSpan.FromSeconds(60)) == 1) return 1;
            var exe = Environment.ProcessPath is { } path ? StartMenuShortcut.RealPath(path) : "";
            return Uninstall.Run(paths, exe, TimeSpan.FromSeconds(5), () =>
            {
                if (OperatingSystem.IsWindows()) StartMenuShortcut.Remove(StartMenuShortcut.DefaultPath, exe);
            });
        }
        // En fejl, som intet fanger, skrives i loggen, før appen lukker (Indstillinger → Fejlfinding).
        AppDomain.CurrentDomain.UnhandledException += (_, e) => new FileLog(paths.Log).Error("Uventet fejl", e.ExceptionObject as Exception);
        // Låsen slippes først, når processen er helt væk (styresystemet lukker den), så --quit først melder "afsluttet",
        // når AulaSync.exe ikke længere er i brug og kan skiftes ud eller slettes.
        var instance = SingleInstance.TryAcquire(paths.LockFile);
        if (instance is null)
        {
            // AulaSync kører allerede: bed den kørende instans om at vise hovedvinduet, og afslut.
            InstanceChannel.SignalShowAsync(InstanceChannel.DefaultName, TimeSpan.FromSeconds(2)).GetAwaiter().GetResult();
            return 0;
        }
        App.Silent = args.Contains("--silent");
        var code = BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
        GC.KeepAlive(instance);
        return code;
    }

    // Beder den kørende AulaSync om at afslutte og venter, til låsen er fri. 0: ingen AulaSync kørte; 2: den kørte og er
    // afsluttet (installationsprogrammet starter den igen bagefter); 1: den kører stadig (fx en ældre udgave, der ikke kender
    // beskeden; så lukker installationsprogrammet den selv).
    internal static int QuitRunning(string lockFile, string channel, TimeSpan wait)
    {
        if (!File.Exists(lockFile)) return 0;
        var deadline = DateTime.UtcNow + wait;
        var asked = false;
        while (true)
        {
            using (var free = SingleInstance.TryAcquire(lockFile)) { if (free is not null) return asked ? 2 : 0; }
            if (DateTime.UtcNow >= deadline) return 1;
            if (!asked) asked = InstanceChannel.SignalQuitAsync(channel, TimeSpan.FromSeconds(2)).GetAwaiter().GetResult();
            Thread.Sleep(200);
        }
    }

    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder.Configure<App>().UsePlatformDetect().LogToTrace().With(MacOptions);

    // Mac: app-menuen er AulaSyncs egen på dansk (NativeMenus.AppMenu); Avalonias engelske standardpunkter slås fra.
    // Appen starter uden Dock-ikon og får det, når et vindue vises (AppWindows.UpdateDock), så en start ved login
    // (--silent) kun ses i menulinjen.
    internal static MacOSPlatformOptions MacOptions => new() { DisableDefaultApplicationMenuItems = true, ShowInDock = false };
}
