using System.Net;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Threading;
using AulaSync.Core;

namespace AulaSync.App;

// Aulas login-side i en NativeWebView (WebView2 / WKWebView) med den persistente profil.
// Login er gennemført, når PHPSESSID og Csrfp-Token findes for aula.dk (spec §3.2, trin 2).
public sealed class LoginView : UserControl
{
    public static readonly Uri LoginUri = new("https://www.aula.dk/auth/login.php?type=unilogin");
    public const string MissingRuntimeText =
        "Login kræver Microsoft Edge WebView2 Runtime, som mangler på denne computer. Installér den fra " +
        "https://go.microsoft.com/fwlink/p/?LinkId=2124703 (eller bed IT om det), og start så AulaSync igen.";

    readonly NativeWebView _web = new();
    readonly TextBlock _status = new() { Classes = { "sub", "error" }, TextWrapping = Avalonia.Media.TextWrapping.Wrap, Margin = new(0, 0, 0, 8), IsVisible = false };
    readonly DispatcherTimer _poll = new() { Interval = TimeSpan.FromMilliseconds(500) };
    readonly BrowserProfile _profile;
    readonly LoginViewModel _vm;
    bool _found;
    bool _checking;
    string? _failedSession; // en session, der allerede er afvist, prøves ikke igen

    public LoginView(BrowserProfile profile, LoginViewModel vm) : this(profile, vm, BrowserProfile.WebView2Missing) { }

    // runtimeMissing: Windows uden WebView2 Runtime. Så vises kun en besked om, hvad der mangler (intet webview).
    internal LoginView(BrowserProfile profile, LoginViewModel vm, bool runtimeMissing)
    {
        _profile = profile;
        _vm = vm;
        if (runtimeMissing)
        {
            Content = new SelectableTextBlock { Text = MissingRuntimeText, TextWrapping = Avalonia.Media.TextWrapping.Wrap, Margin = new(0, 8) };
            return;
        }
        _web.EnvironmentRequested += (_, e) => profile.Configure(e);
        _web.AdapterCreated += async (_, _) =>
        {
            await ClearCookiesAfterLogoutAsync();
            _web.Navigate(LoginUri);
        };
        // MitID og visse UniLogin-veje åbner nye vinduer: vis dem i samme webview.
        _web.NewWindowRequested += (_, e) =>
        {
            e.Handled = true;
            if (e.Request is { } uri) _web.Navigate(uri);
        };
        _poll.Tick += async (_, _) => await PollAsync();
        AttachedToVisualTree += (_, _) => _poll.Start();
        DetachedFromVisualTree += (_, _) => _poll.Stop();
        vm.PropertyChanged += (_, _) =>
        {
            _status.Text = vm.Error;
            _status.IsVisible = vm.Error is not null;
        };

        var root = new DockPanel();
        DockPanel.SetDock(_status, Dock.Top);
        root.Children.Add(_status);
        root.Children.Add(_web);
        _web.VerticalAlignment = VerticalAlignment.Stretch;
        Content = root;
    }

    async Task ClearCookiesAfterLogoutAsync()
    {
        if (!_profile.ClearCookiesOnNextLogin || _web.TryGetCookieManager() is not { } cookies) return;
        foreach (var c in await cookies.GetCookiesAsync()) cookies.DeleteCookie(c);
        _profile.CookiesCleared();
    }

    Task PollAsync() =>
        _web.TryGetCookieManager() is { } manager ? PollAsync(async () => await manager.GetCookiesAsync()) : Task.CompletedTask;

    // Én kontrol ad gangen: at hente cookies kan tage længere end et tick (500 ms), fx første gang efter start, og to
    // samtidige kontroller ville logge ind to gange (og synkronisere alt to gange).
    internal async Task PollAsync(Func<Task<IEnumerable<Cookie>>> getCookies)
    {
        if (_found || _checking) return;
        _checking = true;
        try
        {
            var aula = (await getCookies())
                .Where(c => c.Domain.TrimStart('.').EndsWith("aula.dk", StringComparison.OrdinalIgnoreCase))
                .ToList();
            var session = aula.FirstOrDefault(c => c.Name == AulaSession.SessionCookie)?.Value;
            if (!AulaSession.HasRequiredCookies(aula) || session == _failedSession) return;
            _found = true;
            if (!await _vm.CookiesFoundAsync(aula))
            {
                _failedSession = session;
                _found = false;
                _web.Navigate(LoginUri);
            }
        }
        finally { _checking = false; }
    }
}
