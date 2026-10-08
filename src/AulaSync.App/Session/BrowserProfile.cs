using Avalonia.Controls;
using Avalonia.Platform;
using AulaSync.Core;

namespace AulaSync.App;

// Den persistente browserprofil til login (spec §3.3). Windows: mappen webview/. Mac: et WKWebView-datalager,
// hvis id står i webview-id.txt. Log ud: Windows sletter mappen (er den låst, slettes cookies ved næste login);
// Mac skifter til et nyt id.
public sealed class BrowserProfile(AppPaths paths, bool isMac)
{
    public Guid StoreId
    {
        get
        {
            if (File.Exists(paths.WebViewStoreId) && Guid.TryParse(File.ReadAllText(paths.WebViewStoreId).Trim(), out var id)) return id;
            return NewStoreId();
        }
    }

    public bool ClearCookiesOnNextLogin => File.Exists(paths.LogoutMarker);

    // Windows uden WebView2 Runtime (fx visse skole-pc'er): Avalonia ville falde tilbage til EdgeHTML (WebView1), som ikke
    // kan give login-cookies, så login blev aldrig færdigt. På Mac er WebView2 ikke understøttet og mangler derfor ikke.
    public static bool WebView2Missing => RuntimeMissing(WebViewAdapterInfo.GetAdapterInfo(WebViewAdapterType.WebView2));

    public static bool RuntimeMissing(DetailedWebViewAdapterInfo info) => info is { IsSupported: true, IsInstalled: false };

    public void Configure(WebViewEnvironmentRequestedEventArgs e)
    {
        if (e is WindowsWebView2EnvironmentRequestedEventArgs windows) windows.UserDataFolder = paths.WebView;
        if (e is AppleWKWebViewEnvironmentRequestedEventArgs mac) mac.DataStoreIdentifier = StoreId;
    }

    public void Reset()
    {
        if (isMac) { NewStoreId(); return; }
        try
        {
            if (Directory.Exists(paths.WebView)) Directory.Delete(paths.WebView, recursive: true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            AtomicFile.WriteAllText(paths.LogoutMarker, "");
        }
    }

    public void CookiesCleared() => File.Delete(paths.LogoutMarker);

    Guid NewStoreId()
    {
        var id = Guid.NewGuid();
        AtomicFile.WriteAllText(paths.WebViewStoreId, id.ToString());
        return id;
    }
}
