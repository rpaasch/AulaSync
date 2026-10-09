using System.Diagnostics;

namespace AulaSync.App;

// "Afinstallér AulaSync…" i Indstillinger (SettingsViewModel).
public interface IUninstall
{
    // Installationsprogrammets afinstallation spørger selv; så spørger AulaSync ikke først.
    bool AsksItself { get; }
    void Start();
}

// Installeret med installationsprogrammet (Windows): dets afinstallation startes; den spørger, afslutter AulaSync med --quit
// og viser sine beskeder. Ellers startes oprydningen (Uninstall, --uninstall), og først når den er startet, slås start ved
// login fra, og AulaSync afslutter. Kan intet startes, siger AulaSync til (notify) og kører videre som før.
// setupDir: hvor installationsprogrammet har installeret AulaSync (Uninstall.RegisteredSetupDir).
public sealed class UninstallLauncher(string exePath, bool isWindows, bool isMac, IAutostart autostart, Action quit,
    Action<string> notify, Func<ProcessStartInfo, bool> start, Func<string?> setupDir) : IUninstall
{
    public const string FailedWindows = "AulaSync kunne ikke afinstalleres. Prøv igen, eller afinstallér den fra Indstillinger › Apps.";
    public const string FailedMac = "AulaSync kunne ikke afinstalleres. Afslut AulaSync, og træk den til papirkurven.";

    string? Installer => isWindows ? Uninstall.InstallerFor(exePath, setupDir()) : null;

    public bool AsksItself => Installer is not null;

    public void Start()
    {
        if (Installer is { } uninstaller)
        {
            if (!start(new ProcessStartInfo(uninstaller) { UseShellExecute = false })) notify(FailedWindows);
            return;
        }
        if (!start(Uninstall.CleanupStart(exePath, isMac)))
        {
            notify(isMac ? FailedMac : FailedWindows);
            return;
        }
        try { autostart.SetEnabled(false); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException) { }
        quit();
    }

    // Starter en proces; false, hvis det ikke kunne lade sig gøre. Mac: open venter, til LaunchServices har startet
    // oprydningen (et øjeblik), så den ikke lukkes sammen med AulaSync, og skal lykkes.
    public static bool StartProcess(ProcessStartInfo info)
    {
        try
        {
            using var process = Process.Start(info);
            if (process is null) return false;
            return info.FileName != "/usr/bin/open" || (process.WaitForExit(10_000) && process.ExitCode == 0);
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or IOException or InvalidOperationException) { return false; }
    }
}
