using Avalonia;
using Avalonia.Controls;

namespace AulaSync.App;

// Åbner login-webview'et usynligt med den persistente profil. Giver IdP'ens SSO nye cookies uden brugerens hjælp,
// er man logget ind igen (afprøvet i spiken, punkt 5c).
public static class SilentLogin
{
    public static async Task<bool> RunAsync(BrowserProfile browser, SessionController session, CancellationToken ct)
    {
        var vm = new LoginViewModel(session);
        var signedIn = WaitAsync(vm, ct);
        var window = new Window
        {
            Width = 400, Height = 400, Opacity = 0, ShowInTaskbar = false, ShowActivated = false,
            WindowDecorations = WindowDecorations.None,
            WindowStartupLocation = WindowStartupLocation.Manual, Position = new PixelPoint(-20000, -20000),
            Content = new LoginView(browser, vm),
        };
        window.Show();
        try { return await signedIn; }
        finally { window.Close(); }
    }

    // true, når der er logget ind. Ved fristen (ct) gives op — medmindre et login er i gang; så ventes på det, så et login,
    // der lykkes lige efter fristen, ikke meldes som "logget ud".
    internal static Task<bool> WaitAsync(LoginViewModel vm, CancellationToken ct)
    {
        var done = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        vm.SignedIn += _ => done.TrySetResult(true);
        void GiveUpUnlessBusy() { if (ct.IsCancellationRequested && !vm.IsBusy) done.TrySetResult(false); }
        vm.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(LoginViewModel.IsBusy)) GiveUpUnlessBusy(); };
        var registration = ct.Register(GiveUpUnlessBusy);
        _ = done.Task.ContinueWith(_ => registration.Dispose(), TaskScheduler.Default);
        return done.Task;
    }
}
