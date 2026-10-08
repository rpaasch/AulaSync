using Avalonia;
using Avalonia.Headless;
using Avalonia.Markup.Xaml.Styling;
using Avalonia.Themes.Fluent;

[assembly: AvaloniaTestApplication(typeof(AulaSync.App.Tests.TestApp))]

namespace AulaSync.App.Tests;

// Headless-app med samme tema som den rigtige, men uden ikon, server og baggrundsarbejde.
public sealed class TestApp : Application
{
    public override void Initialize()
    {
        // Indlæs app-assemblyen først, så dens forkompilerede XAML kan findes, også når én test køres alene.
        _ = typeof(SessionController).Assembly;
        Styles.Add(new FluentTheme());
        Styles.Add(new StyleInclude(new Uri("avares://AulaSync/")) { Source = new Uri("avares://AulaSync/Theme/AulaSyncTheme.axaml") });
    }

    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder.Configure<TestApp>().UseSkia().UseHeadless(new AvaloniaHeadlessPlatformOptions { UseHeadlessDrawing = false });
}
