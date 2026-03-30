using System.Net;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

namespace AulaSync;

/// <summary>
/// Browser-baseret login til Aula via WebView2.
/// Understøtter alle kommuners IdP (UniLogin, Azure AD, MitID m.fl.).
/// </summary>
class AulaLogin
{
    public CookieContainer Cookies { get; } = new();
    public bool Success { get; private set; }

    private const string LoginUrl = "https://www.aula.dk/auth/login.php?type=unilogin";
    private Form? _loginForm;

    public async Task<bool> LoginAsync()
    {
        var tcs = new TaskCompletionSource<bool>();

        _loginForm = new Form
        {
            Text = "AulaSync - Log ind",
            Width = 520,
            Height = 700,
            StartPosition = FormStartPosition.CenterScreen,
            Icon = TrayApp.AppIcon,
        };

        var webView = new WebView2 { Dock = DockStyle.Fill };
        _loginForm.Controls.Add(webView);

        _loginForm.FormClosed += (_, _) =>
        {
            if (!Success) tcs.TrySetResult(false);
        };

        _loginForm.Load += async (_, _) =>
        {
            // Separat brugerdata-mappe så vi ikke deler cookies med Edge
            var userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".aulasync", "webview2");
            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
            await webView.EnsureCoreWebView2Async(env);

            // Poll cookies periodisk
            var timer = new System.Windows.Forms.Timer { Interval = 500 };
            timer.Tick += async (_, _) =>
            {
                try
                {
                    var cookies = await webView.CoreWebView2.CookieManager
                        .GetCookiesAsync("https://www.aula.dk");

                    bool hasSession = false, hasCsrf = false;
                    foreach (var c in cookies)
                    {
                        Cookies.Add(new Cookie(c.Name, c.Value, c.Path, c.Domain));
                        if (c.Name == "PHPSESSID") hasSession = true;
                        if (c.Name == "Csrfp-Token") hasCsrf = true;
                    }

                    if (hasSession && hasCsrf)
                    {
                        timer.Stop();
                        Success = true;
                        _loginForm.Close();
                        tcs.TrySetResult(true);
                    }
                }
                catch { }
            };
            timer.Start();

            webView.CoreWebView2.Navigate(LoginUrl);
        };

        _loginForm.Show();
        return await tcs.Task;
    }

    public HttpClient CreateHttpClient()
    {
        var handler = new HttpClientHandler { CookieContainer = Cookies };
        var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(30) };
        client.DefaultRequestHeaders.Add("Accept", "application/json, text/plain, */*");

        // Sæt CSRF header
        foreach (Cookie c in Cookies.GetCookies(new Uri("https://www.aula.dk")))
        {
            if (c.Name == "Csrfp-Token")
            {
                client.DefaultRequestHeaders.Add("csrfp-token", c.Value);
                break;
            }
        }
        return client;
    }
}
