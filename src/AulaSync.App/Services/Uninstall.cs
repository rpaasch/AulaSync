using System.Diagnostics;
using System.Runtime.Versioning;
using AulaSync.Core;
using Microsoft.Win32;

namespace AulaSync.App;

// "Afinstallér AulaSync…" i Indstillinger fjerner alt, AulaSync har lagt på computeren (UninstallLauncher). Er AulaSync
// installeret med installationsprogrammet (unins000.exe ved siden af exe'en), køres dets afinstallation (AulaSync.iss), som
// gør det samme. Ellers (Mac, løs AulaSync.exe eller den portable fra winget) starter AulaSync sig selv igen med
// --uninstall og afslutter; den nye proces venter, til den første er væk, og sletter så alt (Run). Kalenderne i
// kalenderprogrammet kan AulaSync ikke slette.
public static class Uninstall
{
    public const string Argument = "--uninstall";
    public const string StartupApprovedKey = @"Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run";
    const string UninstallKeys = @"Software\Microsoft\Windows\CurrentVersion\Uninstall";
    // Installationsprogrammets post under Installerede apps (AulaSync.iss, AppGuid).
    const string SetupKey = UninstallKeys + @"\{EB8DDE1A-F008-42A1-AA73-7110F88E4044}_is1";
    // Login-browseren (WebView2) kan holde filer i datamappen åbne et øjeblik, efter at AulaSync er afsluttet.
    static readonly TimeSpan DeleteFor = TimeSpan.FromSeconds(30);

    // Windows: installationsprogrammets afinstallation, hvis AulaSync er installeret med AulaSync-Setup.exe: unins000.exe
    // ligger ved siden af exe'en, og AulaSyncs post under Installerede apps (setupDir) peger på den mappe. En anden
    // unins000.exe i samme mappe hører til et andet program og køres ikke.
    public static string? InstallerFor(string exePath, string? setupDir)
    {
        if (Path.GetDirectoryName(exePath) is not { } dir || setupDir is null) return null;
        if (!string.Equals(Path.TrimEndingDirectorySeparator(setupDir), dir, StringComparison.OrdinalIgnoreCase)) return null;
        var uninstaller = Path.Combine(dir, "unins000.exe");
        return File.Exists(uninstaller) ? uninstaller : null;
    }

    // Mappen, AulaSync er installeret i med installationsprogrammet (InstallLocation i dets post); null uden.
    public static string? RegisteredSetupDir()
    {
        if (!OperatingSystem.IsWindows()) return null;
        using var key = Registry.CurrentUser.OpenSubKey(SetupKey);
        return key?.GetValue("InstallLocation") as string;
    }

    // Mac: AulaSync.app ud fra stien til programfilen (…/AulaSync.app/Contents/MacOS/AulaSync); null uden for en .app.
    public static string? BundleOf(string exePath)
    {
        var bundle = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(exePath) ?? "", "..", ".."));
        return bundle.EndsWith(".app", StringComparison.OrdinalIgnoreCase) ? bundle : null;
    }

    // Mac: macOS' egne mapper og filer for AulaSync (login-browserens WebKit-data med cookies, cache, gemt vinduestilstand).
    public static IReadOnlyList<string> MacLibrary(string home, string bundleId)
    {
        var library = Path.Combine(home, "Library");
        return
        [
            Path.Combine(library, "WebKit", bundleId),
            Path.Combine(library, "Caches", bundleId),
            Path.Combine(library, "HTTPStorages", bundleId),
            Path.Combine(library, "HTTPStorages", bundleId + ".binarycookies"),
            Path.Combine(library, "Saved Application State", bundleId + ".savedState"),
        ];
    }

    // Windows: den portable AulaSync fra winget ligger i %LOCALAPPDATA%\Microsoft\WinGet\Packages\rpaasch.AulaSync_<kilde>;
    // winget husker den under Installerede apps med mappens navn og har en henvisning i Links. null for andre placeringer.
    public static (string PackageDir, string UninstallKey, string Link)? WingetPortable(string exePath, string localAppData)
    {
        var dir = Path.GetDirectoryName(exePath);
        var packages = Path.Combine(localAppData, "Microsoft", "WinGet", "Packages");
        if (dir is null || !string.Equals(Path.GetDirectoryName(dir), packages, StringComparison.OrdinalIgnoreCase)) return null;
        var name = Path.GetFileName(dir);
        if (!name.StartsWith("rpaasch.AulaSync_", StringComparison.OrdinalIgnoreCase)) return null;
        return (dir, $@"{UninstallKeys}\{name}", Path.Combine(localAppData, "Microsoft", "WinGet", "Links", "AulaSync.exe"));
    }

    // Oprydningen (--uninstall) i en ny proces. Mac: gennem LaunchServices (open -n), så den ikke lukkes sammen med
    // AulaSync, når den er startet ved login (launchd lukker hele procesgruppen).
    public static ProcessStartInfo CleanupStart(string exePath, bool isMac)
    {
        var start = isMac && BundleOf(exePath) is { } bundle
            ? new ProcessStartInfo("/usr/bin/open") { ArgumentList = { "-n", "-a", bundle, "--args", Argument } }
            : new ProcessStartInfo(exePath) { ArgumentList = { Argument } };
        start.UseShellExecute = false;
        return start;
    }

    // --uninstall: venter, til AulaSync er afsluttet, og sletter alt. 0: slettet; 1: AulaSync kører stadig.
    // removeShortcut: Windows: fjerner genvejen i Start-menuen (StartMenuShortcut.Remove).
    public static int Run(AppPaths paths, string exePath, TimeSpan wait, Action? removeShortcut = null)
    {
        if (!WaitForExit(paths.LockFile, wait)) return 1;
        if (OperatingSystem.IsWindows()) RemoveWindows(exePath, removeShortcut);
        if (OperatingSystem.IsMacOS()) RemoveMac(exePath);
        DeleteDirectory(paths.Root, DeleteFor);
        if (OperatingSystem.IsWindows())
        {
            // AulaSync 2's mappe med det gamle login (som installationsprogrammets afinstallation).
            if (Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) is { Length: > 0 } profile)
                DeleteDirectory(Path.Combine(profile, ".aulasync"), DeleteFor);
            UninstallSetup(exePath);
            DeleteAfterExit(exePath, paths.Root);
        }
        return 0;
    }

    // Venter, til låsen er fri (AulaSync er afsluttet). Låsen slippes igen, så mappen kan slettes.
    public static bool WaitForExit(string lockFile, TimeSpan wait)
    {
        if (!File.Exists(lockFile)) return true;
        var deadline = DateTime.UtcNow + wait;
        while (true)
        {
            using (var free = SingleInstance.TryAcquire(lockFile)) { if (free is not null) return true; }
            if (DateTime.UtcNow >= deadline) return false;
            Thread.Sleep(200);
        }
    }

    // Sletter en mappe helt; prøver igen, mens filer er i brug, højst "within".
    public static bool DeleteDirectory(string dir, TimeSpan within)
    {
        var deadline = DateTime.UtcNow + within;
        while (true)
        {
            try
            {
                if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true);
                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                if (DateTime.UtcNow >= deadline) return false;
                Thread.Sleep(250);
            }
        }
    }

    static void DeleteFile(string path)
    {
        try { File.Delete(path); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    [SupportedOSPlatform("windows")]
    static void RemoveWindows(string exePath, Action? removeShortcut)
    {
        using (var run = Registry.CurrentUser.OpenSubKey(WindowsAutostart.RunKey, writable: true))
            run?.DeleteValue("AulaSync", throwOnMissingValue: false);
        using (var approved = Registry.CurrentUser.OpenSubKey(StartupApprovedKey, writable: true))
            approved?.DeleteValue("AulaSync", throwOnMissingValue: false);
        try { removeShortcut?.Invoke(); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (WingetPortable(exePath, localAppData) is { } winget)
        {
            // Kunne winget ikke lave henvisningen (uden udviklertilstand), har den lagt mappen i brugerens PATH.
            using (var arp = Registry.CurrentUser.OpenSubKey(winget.UninstallKey))
                if (arp?.GetValue("InstallDirectoryAddedToPath") is int added && added != 0) RemoveFromUserPath(winget.PackageDir);
            Registry.CurrentUser.DeleteSubKeyTree(winget.UninstallKey, throwOnMissingSubKey: false);
            if (new FileInfo(winget.Link).LinkTarget is { } to
                && string.Equals(Path.GetFullPath(to, Path.GetDirectoryName(winget.Link)!), exePath, StringComparison.OrdinalIgnoreCase))
                DeleteFile(winget.Link);
        }
    }

    // Fjerner en mappe fra brugerens PATH (HKCU\Environment), med samme værditype (REG_EXPAND_SZ) og uden at folde
    // %…% ud, og giver programmerne besked.
    [SupportedOSPlatform("windows")]
    static void RemoveFromUserPath(string dir)
    {
        using var env = Registry.CurrentUser.OpenSubKey("Environment", writable: true);
        if (env?.GetValue("Path", null, RegistryValueOptions.DoNotExpandEnvironmentNames) is not string path) return;
        var kept = WithoutDir(path, dir);
        if (kept == path) return;
        env.SetValue("Path", kept, env.GetValueKind("Path"));
        SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, "Environment", SMTO_ABORTIFHUNG, 1000, out _);
    }

    // PATH uden mappen (uden forskel på store og små bogstaver og en afsluttende \).
    public static string WithoutDir(string path, string dir)
    {
        static string Norm(string p) => p.Trim().TrimEnd('\\', '/');
        var parts = path.Split(';');
        var kept = parts.Where(p => !string.Equals(Norm(p), Norm(dir), StringComparison.OrdinalIgnoreCase)).ToArray();
        return kept.Length == parts.Length ? path : string.Join(';', kept);
    }

    const int HWND_BROADCAST = 0xffff, WM_SETTINGCHANGE = 0x001A, SMTO_ABORTIFHUNG = 0x0002;

    [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    static extern IntPtr SendMessageTimeout(IntPtr hWnd, int msg, IntPtr wParam, string lParam, int flags, int timeout, out IntPtr result);

    // Er AulaSync også installeret med installationsprogrammet (et andet sted end denne exe), afinstalleres den også, uden
    // vinduer.
    [SupportedOSPlatform("windows")]
    static void UninstallSetup(string exePath)
    {
        using var key = Registry.CurrentUser.OpenSubKey(SetupKey);
        if (key?.GetValue("InstallLocation") is not string location || key.GetValue("UninstallString") is not string command) return;
        if (string.Equals(Path.TrimEndingDirectorySeparator(location), Path.GetDirectoryName(exePath), StringComparison.OrdinalIgnoreCase)) return;
        var uninstaller = command.Trim('"');
        if (!File.Exists(uninstaller)) return;
        try
        {
            Process.Start(new ProcessStartInfo(uninstaller) { ArgumentList = { "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART" }, UseShellExecute = false })?.Dispose();
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or IOException) { }
    }

    // Programfilen kan ikke slette sig selv, mens den kører; det gør cmd, når den er afsluttet. Også datamappen, hvis noget
    // i den var i brug, .NET's udpakkede filer (%TEMP%\.net\AulaSync, samme TEMP som installationsprogrammet) og winget's
    // tomme mappe.
    [SupportedOSPlatform("windows")]
    static void DeleteAfterExit(string exePath, string dataDir)
    {
        var temp = Path.Combine(Path.GetTempPath(), ".net", "AulaSync"); // samme mappe som .NET (TMP før TEMP)
        var package = WingetPortable(exePath, Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))?.PackageDir;
        var start = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "cmd.exe"))
        {
            Arguments = DeleteCommand(package is not null),
            UseShellExecute = false,
            CreateNoWindow = true,
            WorkingDirectory = Path.GetTempPath(), // ikke exe'ens mappe, så den kan slettes
        };
        foreach (var (name, value) in DeletePaths(exePath, dataDir, temp, package)) start.Environment[name] = value;
        try { Process.Start(start)?.Dispose(); }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or IOException) { }
    }

    // Argumenterne til cmd.exe i DeleteAfterExit: prøv at slette exe'en hvert sekund i op til 30 sekunder (ping venter), og
    // slet så mapperne. Stierne kommer i miljøvariabler (DeletePaths) og foldes først ud til sidst (!…!, /v:on), så cmd ikke
    // ændrer fx "%i" i en sti. /s: kun de yderste anførselstegn fjernes.
    public static string DeleteCommand(bool package)
    {
        var command = "for /l %i in (1,1,30) do @if exist \"!AULASYNC_EXE!\" (ping -n 2 127.0.0.1 >nul & del /f /q \"!AULASYNC_EXE!\")"
            + " & rmdir /s /q \"!AULASYNC_DATA!\" & rmdir /s /q \"!AULASYNC_NET!\"";
        if (package) command += " & rmdir \"!AULASYNC_PACKAGE!\"";
        return $"/v:on /d /s /c \"{command}\"";
    }

    public static IEnumerable<(string Name, string Value)> DeletePaths(string exePath, string dataDir, string tempDir, string? packageDir)
    {
        yield return ("AULASYNC_EXE", exePath);
        yield return ("AULASYNC_DATA", dataDir);
        yield return ("AULASYNC_NET", tempDir);
        if (packageDir is not null) yield return ("AULASYNC_PACKAGE", packageDir);
    }

    // Mac: start ved login (LaunchAgent, også i launchd), macOS' mapper for AulaSync, indstillinger i cfprefsd og selve
    // AulaSync.app, som flyttes til papirkurven (ikke fra .dmg-filen eller macOS' midlertidige kopi).
    static void RemoveMac(string exePath)
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        DeleteFile(LaunchAgent.PlistPath(home));
        RunQuietly("/bin/launchctl", "bootout", $"gui/{UserId()}/{LaunchAgent.Label}");
        foreach (var path in MacLibrary(home, LaunchAgent.Label))
            if (Directory.Exists(path)) DeleteDirectory(path, DeleteFor); else DeleteFile(path);
        RunQuietly("/usr/bin/defaults", "delete", LaunchAgent.Label);
        // Kan AulaSync.app ikke flyttes (fx ikke administrator på Macen), vises den i Finder, så den kan trækkes væk.
        if (BundleOf(exePath) is { } bundle && !MacAutostart.IsTransient(bundle) && !MoveToTrash(bundle, home))
            RunQuietly("/usr/bin/open", "-R", bundle);
    }

    // Flytter AulaSync.app til papirkurven (~/.Trash); findes navnet dér, får den et nyt. false, hvis det ikke kunne lade sig gøre.
    public static bool MoveToTrash(string bundle, string home)
    {
        var trash = Path.Combine(home, ".Trash");
        try
        {
            Directory.CreateDirectory(trash);
            var name = Path.GetFileNameWithoutExtension(bundle);
            var target = Path.Combine(trash, name + ".app");
            for (var i = 2; Directory.Exists(target) || File.Exists(target); i++) target = Path.Combine(trash, $"{name} {i}.app");
            Directory.Move(bundle, target);
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return false; }
    }

    [System.Runtime.InteropServices.DllImport("libc")] static extern uint getuid();

    static string UserId() => getuid().ToString(System.Globalization.CultureInfo.InvariantCulture);

    static void RunQuietly(string file, params string[] args)
    {
        try
        {
            var start = new ProcessStartInfo(file) { UseShellExecute = false };
            foreach (var a in args) start.ArgumentList.Add(a);
            using var p = Process.Start(start);
            p?.WaitForExit(10_000);
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or IOException) { }
    }
}
