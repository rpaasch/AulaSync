using AulaSync.Core;

namespace AulaSync.App;

// Fejl, som Avalonias login-webview kaster på brugerfladens tråd uden for AulaSyncs kode: fx når et login-vindue lukkes,
// mens WebView2 starter (E_ABORT), eller når WebView2 ikke kan startes. De skrives i loggen, og appen kører videre (login
// kan prøves igen). Alle andre fejl lukker appen som før og skrives i loggen af Program.
public static class UiErrors
{
    // Navnerummene i Avalonia.Controls.WebView: kontrollen og adapterne til WebView2/WebView1 (Windows) og WKWebView (Mac).
    static readonly string[] LoginWebView = ["Avalonia.Controls.NativeWebView", "Avalonia.Controls.Win.", "Avalonia.Controls.Macios."];

    public static bool TryHandle(Exception ex, FileLog log)
    {
        var text = ex.ToString();
        if (!LoginWebView.Any(name => text.Contains(name, StringComparison.Ordinal))) return false;
        log.Error("Login-webviewet fejlede", ex);
        return true;
    }
}
