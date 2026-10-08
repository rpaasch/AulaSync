using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless.XUnit;
using Avalonia.Platform;
using AulaSync.Core;

namespace AulaSync.App.Tests;

public class AutostartTests
{
    [Fact]
    public void Mac_launch_agent_is_written_and_removed()
    {
        using var home = new TempDir();
        var autostart = new MacAutostart(home.Path, "/Applications/AulaSync.app/Contents/MacOS/AulaSync");
        Assert.False(autostart.IsEnabled);

        autostart.SetEnabled(true);
        Assert.True(autostart.IsEnabled);
        Assert.Contains("--silent", File.ReadAllText(LaunchAgent.PlistPath(home.Path)));

        autostart.SetEnabled(false);
        Assert.False(autostart.IsEnabled);
    }
}

public class TrayIconImageTests
{
    [AvaloniaTheory]
    [InlineData(TrayState.Normal)]
    [InlineData(TrayState.Updating)]
    [InlineData(TrayState.Attention)]
    public void Renders_every_state(TrayState state)
    {
        using var bitmap = TrayIconImage.Render(state, template: true);
        Assert.Equal(new PixelSize(64, 64), bitmap.PixelSize);
    }
}

// Windows: login-webviewet (WebView2) er et indlejret vindue, og Avalonia kan kun lave det, når programmets manifest siger,
// at det er lavet til Windows 10/11. Uden det lukkede appen ved første start ("Unable to create child window for native
// control host. Application manifest with supported OS list might be required.").
public class WindowsManifestTests
{
    [Fact]
    public void App_declares_windows_10_and_11_in_its_manifest()
    {
        var dll = File.ReadAllBytes(typeof(SessionController).Assembly.Location);
        var manifest = System.Text.Encoding.UTF8.GetString(dll);
        Assert.Contains("<supportedOS Id=\"{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}\"", manifest);
    }
}

// Windows: holder et andet program udklipsholderen (fjernskrivebord, Office, en udklipsholder-app), kaster Avalonia en
// COMException efter ca. et sekund. Den skrives i loggen, og kopieringen melder fejl i stedet for at lukke appen.
public class ClipboardTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task A_busy_clipboard_is_logged_not_thrown(bool com)
    {
        using var dir = new TempDir();
        Exception busy = com ? new COMException("OpenClipboard fejlede", unchecked((int)0x800401D0)) : new TimeoutException();
        var platform = new DesktopPlatform(new FileLog(dir.File("log.txt")), _ => Task.FromException(busy));
        Assert.False(await platform.CopyTextAsync("http://localhost:9876/klasse-7.ics"));
        Assert.Contains("Kunne ikke kopiere til udklipsholderen", File.ReadAllText(dir.File("log.txt")));
    }

    [Fact]
    public async Task A_free_clipboard_gets_the_text()
    {
        using var dir = new TempDir();
        var copied = new List<string>();
        var platform = new DesktopPlatform(new FileLog(dir.File("log.txt")), text => { copied.Add(text); return Task.CompletedTask; });
        Assert.True(await platform.CopyTextAsync("7A"));
        Assert.Equal(["7A"], copied);
    }
}

// Windows uden WebView2 Runtime: Avalonia ville falde tilbage til EdgeHTML, som ikke kan give login-cookies.
public class WebView2RuntimeTests
{
    static DetailedWebViewAdapterInfo Info(bool supported, bool installed) =>
        new(WebViewAdapterType.WebView2, WebViewEngine.Blink, null, supported, installed, null, WebViewEmbeddingScenario.NativeControlHost);

    [Fact]
    public void Missing_only_when_windows_supports_it_but_it_is_not_installed()
    {
        Assert.True(BrowserProfile.RuntimeMissing(Info(supported: true, installed: false)));
        Assert.False(BrowserProfile.RuntimeMissing(Info(supported: true, installed: true)));
        Assert.False(BrowserProfile.RuntimeMissing(Info(supported: false, installed: false))); // Mac
    }
}

// Fejl, som Avalonias login-webview kaster på brugerfladens tråd, fx når et login-vindue lukkes, mens WebView2 starter
// (E_ABORT). De skrives i loggen, og appen kører videre. Andre fejl lukker appen som før.
public class UiErrorsTests
{
    sealed class FakeStackException(string stack) : COMException("Handlingen blev afbrudt", unchecked((int)0x80004004))
    {
        public override string? StackTrace => stack;
    }

    [Fact]
    public void Login_webview_errors_are_logged_and_handled()
    {
        using var dir = new TempDir();
        var log = new FileLog(dir.File("log.txt"));
        var aborted = new FakeStackException("   at Avalonia.Controls.NativeWebViewControlHost.<CreateNativeControlCore>g__CompleteAdapter|5_0()");
        Assert.True(UiErrors.TryHandle(aborted, log));
        Assert.Contains("Login-webviewet fejlede", File.ReadAllText(dir.File("log.txt")));

        var adapter = new FakeStackException("   at Avalonia.Controls.Win.WebView2.WebView2BaseAdapter.InitializeAsync()");
        Assert.True(UiErrors.TryHandle(adapter, log));
    }

    [Fact]
    public void Other_errors_are_not_handled()
    {
        using var dir = new TempDir();
        var log = new FileLog(dir.File("log.txt"));
        Assert.False(UiErrors.TryHandle(new FakeStackException("   at AulaSync.App.MainViewModel.Reload()"), log));
        Assert.False(UiErrors.TryHandle(new InvalidOperationException("noget andet"), log));
        Assert.False(File.Exists(dir.File("log.txt")));
    }
}

// Windows: login-vinduet (760 punkter) og første start (640) var højere end skærmen på almindelige bærbare (1920x1080 ved
// 150 %, 1366x768). CenterScreen satte så titellinjen over skærmen, og vinduet kunne hverken flyttes eller lukkes.
public class WindowFitTests
{
    [Fact]
    public void Height_fits_the_work_area_minus_the_title_bar()
    {
        Assert.Equal(640, WindowFit.Height(760, new PixelRect(0, 0, 1920, 1032), 1.5));
        Assert.Equal(760, WindowFit.Height(760, new PixelRect(0, 0, 2560, 1400), 1.0));
    }

    [AvaloniaFact]
    public void Fit_lowers_the_minimum_height_too()
    {
        var window = new Window { Height = 640, MinHeight = 480 };
        WindowFit.Apply(window, new PixelRect(0, 0, 1366, 728), 1.5);
        Assert.Equal(728 / 1.5 - 48, window.Height, 3);
        Assert.Equal(window.Height, window.MinHeight);
    }

    // Indstillinger tilpasser højden til indholdet; på en lav skærm bliver den højst skærmens højde og kan rulles.
    [AvaloniaFact]
    public void Cap_limits_a_window_that_sizes_to_its_content()
    {
        var window = new Window { SizeToContent = SizeToContent.Height };
        WindowFit.Cap(window, new PixelRect(0, 0, 1920, 1032), 1.5);
        Assert.Equal(1032 / 1.5 - 48, window.MaxHeight, 3);
        Assert.Equal(SizeToContent.Height, window.SizeToContent);
    }

    // Klik på ikonet, "Åbn AulaSync…", en notifikation eller en ny start skal vise et minimeret vindue igen. Show og
    // Activate alene gør intet ved et minimeret vindue på Windows.
    [AvaloniaFact]
    public void Showing_a_minimized_window_restores_it()
    {
        var window = new Window();
        window.Show();
        window.WindowState = WindowState.Minimized;
        AppWindows.ShowRestored(window);
        Assert.Equal(WindowState.Normal, window.WindowState);
    }
}

// AulaSync 2 (Windows) kørte med mutexen "AulaSync_SingleInstance". Kører den stadig efter opgraderingen, siger AulaSync 3
// til, så man ikke har to ikoner og to programmer, der svarer på kalenderadresser.
public class OldVersionTests
{
    [Fact]
    public void Sees_a_running_old_version_by_its_mutex()
    {
        var name = $"AulaSyncTest_{Guid.NewGuid():N}";
        Assert.False(OldVersion.IsRunning(name));
        using (new Mutex(false, name)) Assert.True(OldVersion.IsRunning(name));
        Assert.False(OldVersion.IsRunning(name));
        Assert.Equal("AulaSync_SingleInstance", OldVersion.MutexName);
    }
}

// Windows: genvejen AulaSync i Start-menuen skrives, når den mangler, er ødelagt eller peger på en anden exe (fx den gamle
// AulaSync 2), og røres ikke, når den peger rigtigt. Selve genvejen kan kun testes på Windows (prøvebygningen kører testen dér).
public class StartMenuShortcutTests
{
    [Fact]
    public void Same_file_regardless_of_case()
    {
        Assert.True(StartMenuShortcut.PointsTo(@"C:\Users\Anna\Documents\AulaSync.exe", @"c:\users\anna\documents\AULASYNC.EXE"));
        Assert.False(StartMenuShortcut.PointsTo(@"C:\Users\Anna\AulaSync\AulaSync.exe", @"C:\Users\Anna\Documents\AulaSync.exe"));
        Assert.False(StartMenuShortcut.PointsTo(null, @"C:\Users\Anna\Documents\AulaSync.exe"));
        Assert.False(StartMenuShortcut.PointsTo("", @"C:\Users\Anna\Documents\AulaSync.exe"));
    }

    [Fact]
    public void A_missing_file_is_its_own_real_path()
    {
        using var dir = new TempDir();
        var exe = Path.Combine(dir.Path, "AulaSync.exe");
        Assert.Equal(exe, StartMenuShortcut.RealPath(exe));
        File.WriteAllText(exe, "");
        Assert.Equal(exe, StartMenuShortcut.RealPath(exe));
    }

    // winget starter AulaSync gennem en henvisning (Links\AulaSync.exe); genvejen skal pege på selve filen.
    [Fact]
    public void A_link_points_to_the_real_file()
    {
        using var dir = new TempDir();
        var real = Path.Combine(dir.Path, "AulaSync.exe");
        File.WriteAllText(real, "");
        var alias = Path.Combine(dir.Path, "Henvisning.exe");
        try { File.CreateSymbolicLink(alias, real); }
        catch (Exception ex) when (OperatingSystem.IsWindows() && ex is IOException or UnauthorizedAccessException)
        {
            Assert.Skip("Henvisninger kræver udviklertilstand eller administrator på Windows");
            return;
        }
        // Windows kan give TEMP-stien i lang form, så kun filnavnet sammenlignes.
        var resolved = StartMenuShortcut.RealPath(alias);
        Assert.Equal("AulaSync.exe", Path.GetFileName(resolved));
        Assert.True(File.Exists(resolved));

        File.Delete(real);
        Assert.Equal(alias, StartMenuShortcut.RealPath(alias));
    }

    [Fact]
    public void Windows_writes_the_shortcut_when_missing_or_wrong_and_leaves_it_when_right()
    {
        if (!OperatingSystem.IsWindows())
        {
            Assert.Skip("Genveje findes kun på Windows");
            return;
        }
        Assert.EndsWith(@"\Start Menu\Programs\AulaSync.lnk", StartMenuShortcut.DefaultPath);
        using var dir = new TempDir();
        var link = Path.Combine(dir.Path, "Start Menu", "Programs", StartMenuShortcut.FileName);
        var old = Path.Combine(dir.Path, "Gammel", "AulaSync.exe");
        var exe = Path.Combine(dir.Path, "Ny", "AulaSync.exe");
        foreach (var file in new[] { old, exe })
        {
            Directory.CreateDirectory(Path.GetDirectoryName(file)!);
            File.WriteAllText(file, "");
        }

        Assert.True(StartMenuShortcut.Ensure(link, old));
        Assert.Equal(old, StartMenuShortcut.ReadTarget(link), ignoreCase: true);
        var written = File.GetLastWriteTimeUtc(link);

        Assert.False(StartMenuShortcut.Ensure(link, old.ToUpperInvariant()));
        Assert.Equal(written, File.GetLastWriteTimeUtc(link));

        Assert.True(StartMenuShortcut.Ensure(link, exe));
        Assert.Equal(exe, StartMenuShortcut.ReadTarget(link), ignoreCase: true);

        File.WriteAllText(link, "ikke en genvej");
        Assert.True(StartMenuShortcut.Ensure(link, exe));
        Assert.Equal(exe, StartMenuShortcut.ReadTarget(link), ignoreCase: true);
    }
}

// Mac: AulaSync starter uden Dock-ikon (fx ved login med --silent) og får det først, når et vindue vises (AppWindows.Present
// kalder UpdateDock). Uden det stod appen i Dock hele dagen, selv om den kun skal findes i menulinjen (spec §3.2).
public class MacOptionsTests
{
    [Fact]
    public void Starts_without_dock_icon_and_with_our_own_app_menu()
    {
        Assert.False(Program.MacOptions.ShowInDock);
        Assert.True(Program.MacOptions.DisableDefaultApplicationMenuItems);
    }
}
