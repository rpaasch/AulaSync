using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.App.Tests;

// Skærmbillederne af AulaSync på hjemmesiden og i README'en (docs/billeder): appens egne vinduer med opdigtede skemaer, i
// dobbelt opløsning, lyst og mørkt tema (-moerk), som på den Mac eller Windows-pc, de tegnes på (-mac/-windows). Appen
// tegner tekst med styresystemets skrift, så de skal tegnes på Mac og Windows (.github/workflows/skaermbilleder.yml);
// andre steder fejler testen. Uden miljøvariablen AULASYNC_SKAERMBILLEDER gør den intet:
//
//   AULASYNC_SKAERMBILLEDER=$PWD/docs/billeder dotnet test tests/AulaSync.App.Tests --filter Skaermbilleder
public sealed class Skaermbilleder : IDisposable
{
    // Rollen "other" står som Medarbejder i rækkerne.
    static readonly ScheduleRef Anna = new(ScheduleKind.Employee, "1001", "Anna Eksempel", "AE", "other");
    static readonly ScheduleRef Bo = new(ScheduleKind.Employee, "1002", "Bo Testesen", "BT", "other");
    static readonly ScheduleRef SevenA = new(ScheduleKind.Group, "7", "7A");
    static readonly ScheduleRef EightB = new(ScheduleKind.Group, "8", "8B");
    static readonly ScheduleRef Room = new(ScheduleKind.Resource, "412", "Lokale 53");
    static readonly ScheduleRef[] All = [Anna, SevenA, EightB, Room, Bo];

    // Første start, trin 5: to skemaer er tilføjet, resten mangler endnu. Mine skemaer: alle er tilføjet.
    readonly TestHost _partly = new();
    readonly TestHost _done = new();
    readonly string? _dir = Environment.GetEnvironmentVariable("AULASYNC_SKAERMBILLEDER");
    readonly bool _mac = OperatingSystem.IsMacOS();

    public void Dispose() { _partly.Dispose(); _done.Dispose(); }

    sealed class Directory : IAulaDirectory
    {
        public Task<IReadOnlyList<Employee>> GetEmployeesAsync(CancellationToken ct) =>
            Task.FromResult<IReadOnlyList<Employee>>([new("1001", "Anna Eksempel", "AE", "other"), new("1002", "Bo Testesen", "BT", "other")]);
        public Task<IReadOnlyList<NamedItem>> GetGroupsAsync(CancellationToken ct) =>
            Task.FromResult<IReadOnlyList<NamedItem>>([new("7", "7A"), new("8", "8B"), new("9", "9C")]);
        public Task<IReadOnlyList<NamedItem>> GetResourcesAsync(CancellationToken ct) =>
            Task.FromResult<IReadOnlyList<NamedItem>>([new("412", "Lokale 53"), new("413", "Gymnastiksal")]);
    }

    // En uges lektioner pr. skema, så rækkerne viser et realistisk antal.
    static IReadOnlyList<AulaEvent> Week(ScheduleRef s, DateOnly from, DateOnly to) =>
        Enumerable.Range(0, s.Kind == ScheduleKind.Resource ? 19 : s.Kind == ScheduleKind.Group ? 30 : 24)
            .Select(i => Ev(id: $"{s.Key}-{i}")).ToList();

    async Task PrepareAsync(TestHost host, IEnumerable<ScheduleRef> added)
    {
        host.Config.Save(new AppConfig(_mac ? CalendarApp.AppleCalendar : CalendarApp.OutlookClassic));
        host.Client.Events = Week;
        host.Sync.SetClient(host.Client);
        foreach (var s in All) await host.Sync.AddAsync(s, default);
        await host.Sync.SyncAllAsync(default);
        foreach (var s in added)
        {
            await host.Sync.MarkAddedAsync(s);
            await host.Sync.MarkFetchedAsync(s.FileName);
        }
        host.Sync.ReportNextSync(host.Time.GetUtcNow().AddHours(4));
    }

    MainViewModel Main(TestHost host) =>
        new(host.Sync, host.Config, new FakePlatform(_mac), new FakeDialogs(), new FakeActions(), host.Time, Copenhagen,
            () => true, () => true, a => a());

    async Task<OnboardingViewModel> OnboardingAsync(TestHost host, bool signedIn)
    {
        var catalog = new ScheduleCatalog(new Directory());
        var add = new AddScheduleViewModel(() => catalog, () => Anna, host.Sync, host.Log, a => a());
        await add.LoadAsync();
        return new OnboardingViewModel(host.Config, new FakePlatform(_mac), new FakeAutostart(), add, Main(host), () => Anna, () => signedIn);
    }

    // Dobbelt opløsning: vinduets indhold skaleres to gange (en RenderTargetBitmap med 192 dpi tegner ikke indholdet af kort
    // med skygge, fx det valgte kalenderprogram). Gråtoneudglatning, så teksten ikke får farvede kanter, når billedet vises
    // mindre på hjemmesiden. Med fit får vinduet indholdets højde (plus pad), så lister hverken skal rulles eller har tom
    // plads under sig (skriften fylder forskelligt på Mac og Windows). Virtualiserede lister måler kun det synlige, så
    // vinduet vokser derefter, til intet skal rulles.
    void Save(Window window, string name, string suffix, bool fit = true, double pad = 0)
    {
        var content = (Control)window.Content!;
        window.Content = null;
        window.Content = new LayoutTransformControl { LayoutTransform = new ScaleTransform(2, 2), Child = content };
        TextOptions.SetTextRenderingMode(window, TextRenderingMode.Antialias);
        window.Width *= 2;
        window.Height *= 2;
        window.Show();
        Dispatcher.UIThread.RunJobs();
        if (fit)
        {
            content.Measure(new Size(content.Bounds.Width, double.PositiveInfinity));
            window.Height = Math.Ceiling(content.DesiredSize.Height + pad) * 2;
            Dispatcher.UIThread.RunJobs();
            for (var i = 0; i < 5; i++)
            {
                var hidden = window.GetVisualDescendants().OfType<ScrollViewer>().Where(v => v.IsEffectivelyVisible)
                    .Select(v => v.Extent.Height - v.Viewport.Height).DefaultIfEmpty(0).Max();
                if (hidden < 1) break;
                window.Height += Math.Ceiling(hidden) * 2;
                Dispatcher.UIThread.RunJobs();
            }
        }
        AvaloniaHeadlessPlatform.ForceRenderTimerTick(10);
        window.CaptureRenderedFrame()!.Save(Path.Combine(_dir!, $"{name}-{(_mac ? "mac" : "windows")}{suffix}.png"), new PngBitmapEncoderOptions());
        window.Close();
    }

    [AvaloniaFact]
    public async Task Tegn()
    {
        if (string.IsNullOrEmpty(_dir)) return;
        if (!OperatingSystem.IsMacOS() && !OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("Skærmbillederne skal tegnes på Mac eller Windows, så skriften er den samme som i appen (se .github/workflows/skaermbilleder.yml).");
        System.IO.Directory.CreateDirectory(_dir);
        await PrepareAsync(_partly, [Anna, SevenA]);
        await PrepareAsync(_done, All);

        try
        {
            foreach (var (theme, suffix) in new[] { (ThemeVariant.Light, ""), (ThemeVariant.Dark, "-moerk") })
            {
                Application.Current!.RequestedThemeVariant = theme;
                // Velkommen står midt i vinduet, så det har vinduets mindste højde (MinHeight).
                Save(new OnboardingWindow(await OnboardingAsync(_partly, signedIn: false), new Border()) { Width = 600, Height = 480 }, "velkommen", suffix, fit: false);

                var onboarding = await OnboardingAsync(_partly, signedIn: true);
                onboarding.SignedIn();
                Save(new OnboardingWindow(onboarding, new Border()) { Width = 600, Height = 480 }, "kalenderprogram", suffix, pad: 16);
                await onboarding.ContinueCommand.ExecuteAsync(null);
                Save(new OnboardingWindow(onboarding, new Border()) { Width = 600, Height = 480 }, "vaelg-skemaer", suffix);
                await onboarding.ContinueCommand.ExecuteAsync(null);
                Save(new OnboardingWindow(onboarding, new Border()) { Width = 600, Height = 480 }, "tilfoej-til-kalender", suffix);

                Save(new MainWindow(Main(_done)) { Width = 560, Height = 400 }, "mine-skemaer", suffix);
            }
        }
        finally { Application.Current!.RequestedThemeVariant = ThemeVariant.Default; }
    }
}
