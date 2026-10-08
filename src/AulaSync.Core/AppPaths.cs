namespace AulaSync.Core;

public sealed record AppPaths(string Root)
{
    public static AppPaths Default()
    {
        var root = OperatingSystem.IsMacOS()
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library", "Application Support", "AulaSync")
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AulaSync");
        return new AppPaths(root);
    }

    public string Calendars => Path.Combine(Root, "kalendere");
    public string Subscriptions => Path.Combine(Root, "abonnementer.json");
    public string Config => Path.Combine(Root, "config.json");
    public string Log => Path.Combine(Root, "aulasync.log");
    public string WebView => Path.Combine(Root, "webview");
    public string WebViewStoreId => Path.Combine(Root, "webview-id.txt");
    public string LogoutMarker => Path.Combine(Root, "logget-ud");
    public string LockFile => Path.Combine(Root, "aulasync.lock");
}
