using System.Net;
using Microsoft.Extensions.Time.Testing;

namespace AulaSync.App.Tests;

public class ReloginTests : IDisposable
{
    readonly TestHost _host = new();
    readonly FakeNotifier _notifier = new();
    readonly FakeTimeProvider _time = new();

    public void Dispose() => _host.Dispose();

    Relogin Create(Func<CancellationToken, Task<bool>> trySilent, bool supported = true) =>
        new(trySilent, _notifier, _host.Log, _time, supported);

    [Fact]
    public async Task Silent_success_sends_no_notification()
    {
        var relogin = Create(_ => Task.FromResult(true));
        Assert.True(await relogin.TryAsync(Relogin.ExpiredText));
        Assert.Empty(_notifier.Shown);
    }

    [Fact]
    public async Task Failure_notifies_once_until_signed_in_again()
    {
        var relogin = Create(_ => Task.FromResult(false));
        Assert.False(await relogin.TryAsync(Relogin.ExpiredText));
        Assert.False(await relogin.TryAsync(Relogin.ExpiredText));
        Assert.Equal(["Du er logget ud af Aula. Klik her for at logge ind igen."], _notifier.Shown);

        relogin.SignedIn();
        await relogin.TryAsync(Relogin.ExpiredText);
        Assert.Equal(2, _notifier.Shown.Count);
    }

    [Fact]
    public async Task Gives_up_after_30_seconds()
    {
        var relogin = Create(async ct =>
        {
            try { await Task.Delay(Timeout.Infinite, ct); } catch (OperationCanceledException) { }
            return false;
        });
        var attempt = relogin.TryAsync(Relogin.ExpiredText);
        _time.Advance(TimeSpan.FromSeconds(29));
        Assert.False(attempt.IsCompleted);
        _time.Advance(TimeSpan.FromSeconds(1));
        Assert.False(await attempt);
        Assert.Single(_notifier.Shown);
    }

    [Fact]
    public async Task Without_notification_text_nothing_is_sent()
    {
        var relogin = Create(_ => Task.FromResult(false));
        Assert.False(await relogin.TryAsync(null));
        Assert.Empty(_notifier.Shown);
    }

    [Fact]
    public async Task Disabled_after_spike_skips_silent_attempt()
    {
        int attempts = 0;
        var relogin = Create(_ => { attempts++; return Task.FromResult(true); }, supported: false);
        Assert.False(await relogin.TryAsync(Relogin.StartText));
        Assert.Equal(0, attempts);
        Assert.Equal(["Klik her for at logge ind på Aula, så AulaSync kan holde dine skemaer opdateret."], _notifier.Shown);
    }

    [Fact]
    public async Task Concurrent_attempts_run_once()
    {
        var gate = new TaskCompletionSource<bool>();
        int attempts = 0;
        var relogin = Create(_ => { attempts++; return gate.Task; });
        var first = relogin.TryAsync(Relogin.ExpiredText);
        Assert.False(await relogin.TryAsync(Relogin.ExpiredText));
        gate.SetResult(true);
        Assert.True(await first);
        Assert.Equal(1, attempts);
    }
}

// Når fristen for stille genlogin udløber midt i et login, ventes der på det: ellers meldes "logget ud", mens login
// lykkes et øjeblik efter.
public class SilentLoginTests : IDisposable
{
    readonly TestHost _host = new();

    public void Dispose() => _host.Dispose();

    static readonly Cookie[] Cookies = [new("PHPSESSID", "abc", "/", ".aula.dk"), new("Csrfp-Token", "tok123", "/", "www.aula.dk")];

    static HttpResponseMessage Aula(HttpRequestMessage r, string? _) =>
        System.Web.HttpUtility.ParseQueryString(r.RequestUri!.Query)["method"] switch
        {
            "profiles.getProfilesByLogin" => FakeHandler.Json(Fixture.Read("profilesByLogin.json")),
            "profiles.getProfileContext" => FakeHandler.Json(Fixture.Read("profileContext.json")),
            _ => FakeHandler.Json(Fixture.Read("eventsEmpty.json")),
        };

    [Fact]
    public async Task Waits_for_a_login_in_progress_at_the_deadline()
    {
        var gate = new TaskCompletionSource();
        var vm = new LoginViewModel(new SessionController(_host.Sync, _host.Log, _ => new HttpClient(new GatedHandler(gate.Task, Aula))));
        using var deadline = new CancellationTokenSource();
        var wait = SilentLogin.WaitAsync(vm, deadline.Token);
        var login = vm.CookiesFoundAsync(Cookies);

        deadline.Cancel();
        Assert.False(wait.IsCompleted);
        gate.SetResult();

        Assert.True(await wait);
        Assert.True(await login);
    }

    [Fact]
    public async Task Gives_up_at_the_deadline_when_no_login_is_in_progress()
    {
        var vm = new LoginViewModel(new SessionController(_host.Sync, _host.Log, _ => new HttpClient(new FakeHandler(Aula))));
        using var deadline = new CancellationTokenSource();
        var wait = SilentLogin.WaitAsync(vm, deadline.Token);
        deadline.Cancel();
        Assert.False(await wait);
    }
}
