using AulaSync.Core;

namespace AulaSync.App.Tests;

// "Afinstallér AulaSync…" uden installationsprogram (Mac, løs exe): hvad der fjernes, og hvordan.
public class UninstallTests
{
    // Installeret med AulaSync-Setup.exe: afinstallationen ligger ved siden af exe'en, og AulaSyncs post under Installerede
    // apps peger på mappen. En anden unins000.exe (et andet program i samme mappe) køres ikke.
    [Fact]
    public void Installer_is_next_to_the_exe_and_registered()
    {
        using var dir = new TempDir();
        var exe = dir.File("AulaSync.exe");
        File.WriteAllText(exe, "");
        Assert.Null(Uninstall.InstallerFor(exe, dir.Path));
        File.WriteAllText(dir.File("unins000.exe"), "");
        Assert.Equal(dir.File("unins000.exe"), Uninstall.InstallerFor(exe, dir.Path));
        Assert.Equal(dir.File("unins000.exe"), Uninstall.InstallerFor(exe, dir.Path + Path.DirectorySeparatorChar));
        Assert.Null(Uninstall.InstallerFor(exe, null));
        Assert.Null(Uninstall.InstallerFor(exe, dir.File("andet")));
    }

    [Fact]
    public void Bundle_of_the_mac_app()
    {
        Assert.Equal("/Applications/AulaSync.app", Uninstall.BundleOf("/Applications/AulaSync.app/Contents/MacOS/AulaSync"));
        Assert.Null(Uninstall.BundleOf("/home/anna/bin/AulaSync"));
    }

    // Login-browserens WebKit-data (cookies), cache og gemt vinduestilstand under ~/Library.
    [Fact]
    public void Mac_library_folders()
    {
        Assert.Equal(
        [
            "/Users/anna/Library/WebKit/dk.example.aulasync",
            "/Users/anna/Library/Caches/dk.example.aulasync",
            "/Users/anna/Library/HTTPStorages/dk.example.aulasync",
            "/Users/anna/Library/HTTPStorages/dk.example.aulasync.binarycookies",
            "/Users/anna/Library/Saved Application State/dk.example.aulasync.savedState",
        ], Uninstall.MacLibrary("/Users/anna", "dk.example.aulasync"));
    }

    // Den portable AulaSync fra winget: winget's mappe, posten under Installerede apps og henvisningen i Links.
    [Fact]
    public void Winget_portable_is_recognised()
    {
        var local = Path.Combine(Path.GetTempPath(), "Local");
        var dir = Path.Combine(local, "Microsoft", "WinGet", "Packages", "rpaasch.AulaSync_Microsoft.Winget.Source_8wekyb3d8bbwe");
        var winget = Uninstall.WingetPortable(Path.Combine(dir, "AulaSync.exe"), local);
        Assert.NotNull(winget);
        Assert.Equal(dir, winget.Value.PackageDir);
        Assert.Equal(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\rpaasch.AulaSync_Microsoft.Winget.Source_8wekyb3d8bbwe", winget.Value.UninstallKey);
        Assert.Equal(Path.Combine(local, "Microsoft", "WinGet", "Links", "AulaSync.exe"), winget.Value.Link);
        Assert.Null(Uninstall.WingetPortable(Path.Combine(local, "Programs", "AulaSync", "AulaSync.exe"), local));
        Assert.Null(Uninstall.WingetPortable(Path.Combine(local, "Microsoft", "WinGet", "Packages", "Andet.Program_x", "AulaSync.exe"), local));
    }

    // Oprydningen venter, til den kørende AulaSync er afsluttet (låsen er fri), og sletter så datamappen.
    [Fact]
    public void Waits_for_exit_and_deletes_the_data()
    {
        using var dir = new TempDir();
        var paths = new AppPaths(dir.File("AulaSync"));
        Directory.CreateDirectory(paths.Calendars);
        File.WriteAllText(Path.Combine(paths.Calendars, "123456-AE-medarbejder-1001.ics"), "");
        var running = SingleInstance.TryAcquire(paths.LockFile)!;
        Assert.False(Uninstall.WaitForExit(paths.LockFile, TimeSpan.FromMilliseconds(300)));
        running.Dispose();
        Assert.True(Uninstall.WaitForExit(paths.LockFile, TimeSpan.FromSeconds(5)));
        Assert.True(Uninstall.DeleteDirectory(paths.Root, TimeSpan.FromSeconds(5)));
        Assert.False(Directory.Exists(paths.Root));
        Assert.True(Uninstall.DeleteDirectory(paths.Root, TimeSpan.Zero)); // findes ikke: intet at gøre
    }

    // Mac: AulaSync.app flyttes til papirkurven; findes navnet dér, får den et nyt.
    [Fact]
    public void App_goes_to_the_trash()
    {
        using var dir = new TempDir();
        var home = dir.File("home");
        var app = dir.File("Applications/AulaSync.app");
        Directory.CreateDirectory(Path.Combine(app, "Contents"));
        Directory.CreateDirectory(Path.Combine(home, ".Trash", "AulaSync.app"));
        Assert.True(Uninstall.MoveToTrash(app, home));
        Assert.False(Directory.Exists(app));
        Assert.True(Directory.Exists(Path.Combine(home, ".Trash", "AulaSync 2.app", "Contents")));
    }

    // Windows: cmd prøver at slette exe'en hvert sekund i op til 30 sekunder, når AulaSync er afsluttet, og sletter så
    // datamappen (hvis noget var i brug), .NET's filer og winget's mappe. Stierne kommer i miljøvariabler, så cmd ikke
    // ændrer fx "%i" i en sti.
    [Fact]
    public void Delete_command_takes_paths_from_the_environment()
    {
        Assert.Equal("/v:on /d /s /c \"for /l %i in (1,1,30) do @if exist \"!AULASYNC_EXE!\" "
            + "(ping -n 2 127.0.0.1 >nul & del /f /q \"!AULASYNC_EXE!\") "
            + "& rmdir /s /q \"!AULASYNC_DATA!\" & rmdir /s /q \"!AULASYNC_NET!\" & rmdir \"!AULASYNC_PACKAGE!\"\"",
            Uninstall.DeleteCommand(package: true));
        Assert.DoesNotContain("PACKAGE", Uninstall.DeleteCommand(package: false));
        Assert.Equal([("AULASYNC_EXE", @"C:\Users\Anna\Downloads\20%ig\AulaSync.exe"), ("AULASYNC_DATA", @"C:\Data"), ("AULASYNC_NET", @"C:\Temp\.net\AulaSync")],
            Uninstall.DeletePaths(@"C:\Users\Anna\Downloads\20%ig\AulaSync.exe", @"C:\Data", @"C:\Temp\.net\AulaSync", null));
    }

    // winget's mappe fjernes fra PATH (uden forskel på store og små bogstaver og en afsluttende \); resten bliver.
    [Fact]
    public void Package_folder_leaves_the_path()
    {
        const string dir = @"C:\Users\Anna\AppData\Local\Microsoft\WinGet\Packages\rpaasch.AulaSync_x";
        Assert.Equal(@"%USERPROFILE%\bin;C:\Andet", Uninstall.WithoutDir(@"%USERPROFILE%\bin;c:\users\anna\appdata\local\microsoft\winget\packages\rpaasch.aulasync_x\;C:\Andet", dir));
        Assert.Equal(@"C:\Andet", Uninstall.WithoutDir(@"C:\Andet", dir));
    }

    // Oprydningen startes som en ny proces; på Mac gennem LaunchServices, så launchd ikke lukker den sammen med AulaSync.
    [Fact]
    public void Cleanup_starts_a_new_process()
    {
        var mac = Uninstall.CleanupStart("/Applications/AulaSync.app/Contents/MacOS/AulaSync", isMac: true);
        Assert.Equal("/usr/bin/open", mac.FileName);
        Assert.Equal(["-n", "-a", "/Applications/AulaSync.app", "--args", "--uninstall"], mac.ArgumentList);
        var windows = Uninstall.CleanupStart(@"C:\Users\Anna\AulaSync.exe", isMac: false);
        Assert.Equal(@"C:\Users\Anna\AulaSync.exe", windows.FileName);
        Assert.Equal(["--uninstall"], windows.ArgumentList);
    }

    // Installeret med installationsprogrammet: dets afinstallation startes, og den spørger selv. Ellers startes
    // oprydningen, og først derefter slås start ved login fra, og AulaSync afslutter. Kan intet startes, siger AulaSync til.
    [Fact]
    public void Launcher_starts_the_right_thing()
    {
        using var dir = new TempDir();
        var exe = dir.File("AulaSync.exe");
        File.WriteAllText(exe, "");
        var started = new List<string>();
        var notes = new List<string>();
        var quits = 0;
        var autostart = new FakeAutostart(enabled: true);
        var ok = true;
        UninstallLauncher Launcher(bool isWindows) => new(exe, isWindows, isMac: false, autostart, () => quits++, notes.Add,
            info => { started.Add(Path.GetFileName(info.FileName) + " " + string.Join(" ", info.ArgumentList)); return ok; },
            () => dir.Path);

        var loose = Launcher(isWindows: true);
        Assert.False(loose.AsksItself);
        loose.Start();
        Assert.Equal(["AulaSync.exe --uninstall"], started);
        Assert.Equal((false, 1), (autostart.IsEnabled, quits));

        File.WriteAllText(dir.File("unins000.exe"), "");
        var installed = Launcher(isWindows: true);
        Assert.True(installed.AsksItself);
        installed.Start();
        Assert.Equal("unins000.exe ", started[^1]);
        Assert.Equal(1, quits); // afinstallationen afslutter AulaSync selv (--quit)
        Assert.False(Launcher(isWindows: false).AsksItself);

        ok = false;
        autostart.SetEnabled(true);
        Launcher(isWindows: false).Start();
        Assert.Equal((true, 1), (autostart.IsEnabled, quits));
        Assert.Equal([UninstallLauncher.FailedWindows], notes);
    }
}
