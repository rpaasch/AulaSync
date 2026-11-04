using System.Net;
using Microsoft.Extensions.Time.Testing;
using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.App.Tests;

public class ReloginTests : IDisposable
{
    readonly TestHost _host = new();
    readonly FakeNotifier _notifier = new();
    readonly FakeTimeProvider _time = new();

    public void Dispose() => _host.Dispose();

    // notice: med indstillingen i config.json (som i AulaSync); uden: én gang pr. kørsel.
    Relogin Create(Func<CancellationToken, Task<bool>> trySilent, bool supported = true, LoggedOutNotice? notice = null)
    {
        if (notice is { } n) _host.Config.Save(_host.Config.Load() with { LoggedOutNotice = n });
        return new(trySilent, _notifier, _host.Log, _time, supported, notice is null ? null : _host.Config);
    }

    // Hver dag: første besked med det samme; en ny en ny dag tidligst kl. 7 (dansk tid) efter et nyt stille genlogin, ikke
    // samme dag. Lykkes genlogin, ingen besked, og efter et login begynder det forfra.
    [Fact]
    public async Task Daily_notice_comes_again_the_next_morning()
    {
        _time.SetLocalTimeZone(Copenhagen);
        _time.SetUtcNow(new DateTimeOffset(2026, 11, 3, 14, 0, 0, TimeSpan.Zero)); // tirsdag 15:00
        int attempts = 0;
        var silent = false;
        var relogin = Create(_ => { attempts++; return Task.FromResult(silent); }, notice: LoggedOutNotice.Daily);
        Assert.False(await relogin.TryAsync(Relogin.ExpiredText));
        Assert.Single(_notifier.Shown);

        _time.Advance(TimeSpan.FromHours(8)); // 23:00
        Assert.False(await relogin.RemindAsync(Relogin.ExpiredText));
        _time.Advance(TimeSpan.FromHours(7)); // onsdag 6:00
        Assert.False(await relogin.RemindAsync(Relogin.ExpiredText));
        Assert.Equal((1, 1), (attempts, _notifier.Shown.Count));
        _time.Advance(TimeSpan.FromHours(1)); // 7:00
        Assert.False(await relogin.RemindAsync(Relogin.ExpiredText));
        Assert.Equal((2, 2), (attempts, _notifier.Shown.Count));
        _time.Advance(TimeSpan.FromHours(1));
        Assert.False(await relogin.RemindAsync(Relogin.ExpiredText));
        Assert.Equal(2, _notifier.Shown.Count);

        _time.Advance(TimeSpan.FromDays(1));
        silent = true;
        Assert.True(await relogin.RemindAsync(Relogin.ExpiredText));
        Assert.Equal(2, _notifier.Shown.Count);
        Assert.Null(_host.Config.Load().LoggedOutNotifiedAt); // logget ind: forfra
    }

    // En besked om natten (før kl. 7) hører til dagen før, så morgenens besked kommer alligevel samme morgen.
    [Fact]
    public async Task Daily_notice_at_night_does_not_skip_the_morning()
    {
        _time.SetLocalTimeZone(Copenhagen);
        _time.SetUtcNow(new DateTimeOffset(2026, 11, 4, 2, 0, 0, TimeSpan.Zero)); // onsdag 3:00
        var relogin = Create(_ => Task.FromResult(false), notice: LoggedOutNotice.Daily);
        await relogin.TryAsync(Relogin.ExpiredText);
        _time.Advance(TimeSpan.FromHours(3)); // 6:00
        await relogin.RemindAsync(Relogin.ExpiredText);
        Assert.Single(_notifier.Shown);
        _time.Advance(TimeSpan.FromHours(1)); // 7:00
        await relogin.RemindAsync(Relogin.ExpiredText);
        Assert.Equal(2, _notifier.Shown.Count);
    }

    // Én gang: kun den første besked; Ingen: aldrig en besked. Stille genlogin prøves stadig igen en ny dag.
    [Theory]
    [InlineData(LoggedOutNotice.Once, 1)]
    [InlineData(LoggedOutNotice.Never, 0)]
    public async Task Once_and_never(LoggedOutNotice notice, int shown)
    {
        _time.SetUtcNow(new DateTimeOffset(2026, 11, 3, 14, 0, 0, TimeSpan.Zero));
        int attempts = 0;
        var relogin = Create(_ => { attempts++; return Task.FromResult(false); }, notice: notice);
        Assert.False(await relogin.TryAsync(Relogin.StartText));
        Assert.False(await relogin.RemindAsync(Relogin.ExpiredText)); // samme dag: intet nyt forsøg
        _time.Advance(TimeSpan.FromDays(2));
        Assert.False(await relogin.RemindAsync(Relogin.ExpiredText));
        Assert.False(await relogin.TryAsync(Relogin.ExpiredText));
        Assert.Equal(shown, _notifier.Shown.Count);
        Assert.Equal(3, attempts);
    }

    // Hvornår beskeden sidst kom, gemmes: efter en genstart (fx autostart næste morgen) kommer Én gang ikke igen, og Hver dag
    // kun en ny dag.
    [Theory]
    [InlineData(LoggedOutNotice.Once, 0, 1)]
    [InlineData(LoggedOutNotice.Daily, 0, 1)]
    [InlineData(LoggedOutNotice.Daily, 1, 2)]
    public async Task Notice_time_survives_a_restart(LoggedOutNotice notice, int days, int shown)
    {
        _time.SetLocalTimeZone(Copenhagen);
        var start = new DateTimeOffset(2026, 11, 3, 14, 0, 0, TimeSpan.Zero); // 15:00
        _time.SetUtcNow(start);
        await Create(_ => Task.FromResult(false), notice: notice).TryAsync(Relogin.ExpiredText);
        _time.SetUtcNow(days == 0 ? start.AddHours(4) : start.AddDays(days).AddHours(-4)); // 19:00 samme dag eller 11:00
        await Create(_ => Task.FromResult(false), notice: notice).TryAsync(Relogin.StartText);
        Assert.Equal(shown, _notifier.Shown.Count);
    }

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
