using System.Net;
using AulaSync.Core;

namespace AulaSync.App.Tests;

public class SessionControllerTests : IDisposable
{
    static readonly Cookie[] Cookies =
    [
        new("PHPSESSID", "abc", "/", ".aula.dk"),
        new("Csrfp-Token", "tok123", "/", "www.aula.dk"),
    ];

    readonly TestHost _host = new();

    public void Dispose() => _host.Dispose();

    static HttpResponseMessage Aula(HttpRequestMessage r, string? _)
    {
        var method = System.Web.HttpUtility.ParseQueryString(r.RequestUri!.Query)["method"];
        return method switch
        {
            "profiles.getProfilesByLogin" => FakeHandler.Json(Fixture.Read("profilesByLogin.json")),
            "profiles.getProfileContext" => FakeHandler.Json(Fixture.Read("profileContext.json")),
            _ => FakeHandler.Json(Fixture.Read("eventsEmpty.json")),
        };
    }

    SessionController Create(Func<HttpRequestMessage, string?, HttpResponseMessage> respond) =>
        new(_host.Sync, _host.Log, _ => new HttpClient(new FakeHandler(respond)));

    [Fact]
    public async Task Sign_in_connects_sync_and_exposes_own_schedule()
    {
        var session = Create(Aula);
        int changed = 0;
        session.Changed += () => changed++;

        var profile = await session.SignInAsync(Cookies, default);

        Assert.Equal("Test Bruger", profile.Name);
        Assert.Equal(ConnectionState.Online, _host.Sync.Status.Connection);
        Assert.Equal(new ScheduleRef(ScheduleKind.Employee, "9000001", "Test Bruger", "TB", "leader"), session.OwnSchedule);
        Assert.NotNull(session.Catalog);
        Assert.Equal(1, changed);
    }

    // Er eget skema valgt uden rolle (fx før rollen blev vist), får det rollen ved login.
    [Fact]
    public async Task Sign_in_fills_in_the_role_of_the_own_schedule()
    {
        _host.Store.Save([new Subscription(new ScheduleRef(ScheduleKind.Employee, "9000001", "Test Bruger", "TB"))]);
        await Create(Aula).SignInAsync(Cookies, default);
        Assert.Equal("leader", _host.Store.Load().Single().Schedule.Role);
    }

    // Kollegernes roller følger Aula, når listen over medarbejdere hentes (fx når "Tilføj skema" åbnes).
    [Fact]
    public async Task Colleagues_get_their_role_when_the_catalog_is_loaded()
    {
        _host.Store.Save([new Subscription(new ScheduleRef(ScheduleKind.Employee, "3001", "Lene Leder", "LL"))]);
        var session = Create((r, b) => System.Web.HttpUtility.ParseQueryString(r.RequestUri!.Query)["method"] switch
        {
            "search.findProfilesAndGroups" => FakeHandler.Json(Fixture.Read("searchEmployees.json")),
            "profiles.getProfileMasterData" => FakeHandler.Json(Fixture.Read("masterData.json")),
            "search.findGroups" => FakeHandler.Json(Fixture.Read("findGroups.json")),
            "resources.listResources" => FakeHandler.Json(Fixture.Read("resourcesAll.json")),
            _ => Aula(r, b),
        });
        await session.SignInAsync(Cookies, default);
        await session.Catalog!.GetAllAsync();
        Assert.Equal("leader", _host.Store.Load().Single().Schedule.Role);
    }

    [Fact]
    public async Task Expired_session_during_sign_in_throws_and_stays_logged_out()
    {
        var session = Create((_, _) => FakeHandler.Json(Fixture.Read("sessionExpired.json")));
        await Assert.ThrowsAsync<SessionExpiredException>(() => session.SignInAsync(Cookies, default));
        Assert.Null(session.Profile);
        Assert.Equal(ConnectionState.LoggedOut, _host.Sync.Status.Connection);
    }

    [Fact]
    public async Task Sign_out_clears_choices_and_browser_profile_but_keeps_files()
    {
        var session = Create(Aula);
        await session.SignInAsync(Cookies, default);
        await _host.Sync.AddAsync(new ScheduleRef(ScheduleKind.Group, "88231", "7A"), default);
        var file = _host.Sync.FilePath(new ScheduleRef(ScheduleKind.Group, "88231", "7A"));
        var browser = new BrowserProfile(_host.Paths, isMac: true);
        var before = browser.StoreId;

        await session.SignOutAsync(browser);

        Assert.Null(session.Profile);
        Assert.Null(session.Catalog);
        Assert.Empty(_host.Store.Load());
        Assert.True(File.Exists(file));
        Assert.NotEqual(before, browser.StoreId);
        Assert.Equal(ConnectionState.LoggedOut, _host.Sync.Status.Connection);
    }
}

public class BrowserProfileTests : IDisposable
{
    readonly TempDir _dir = new();
    AppPaths Paths => new(_dir.Path);

    public void Dispose() => _dir.Dispose();

    [Fact]
    public void Mac_store_id_is_stable_until_reset()
    {
        var profile = new BrowserProfile(Paths, isMac: true);
        var id = profile.StoreId;
        Assert.Equal(id, profile.StoreId);
        profile.Reset();
        Assert.NotEqual(id, profile.StoreId);
    }

    [Fact]
    public void Windows_reset_deletes_webview_folder()
    {
        Directory.CreateDirectory(Path.Combine(Paths.WebView, "Default"));
        new BrowserProfile(Paths, isMac: false).Reset();
        Assert.False(Directory.Exists(Paths.WebView));
        Assert.False(new BrowserProfile(Paths, isMac: false).ClearCookiesOnNextLogin);
    }
}

public class LoginViewModelTests : IDisposable
{
    readonly TestHost _host = new();

    public void Dispose() => _host.Dispose();

    [Fact]
    public async Task Failed_profile_lookup_shows_error_and_asks_to_retry()
    {
        var session = new SessionController(_host.Sync, _host.Log, _ => new HttpClient(new FakeHandler((_, _) => new HttpResponseMessage(HttpStatusCode.Gone))));
        var vm = new LoginViewModel(session);
        Profile? signedIn = null;
        vm.SignedIn += p => signedIn = p;

        var ok = await vm.CookiesFoundAsync([new("PHPSESSID", "a", "/", ".aula.dk"), new("Csrfp-Token", "b", "/", "www.aula.dk")]);

        Assert.False(ok);
        Assert.Null(signedIn);
        Assert.StartsWith("Du er logget ind, men AulaSync kunne ikke hente din profil", vm.Error);
        Assert.False(vm.IsBusy);
    }
}

// Ticket hvert 500. ms: venter to kontroller samtidig på cookies (fx første gang efter start), må der kun logges ind én gang.
public class LoginViewTests : IDisposable
{
    readonly TestHost _host = new();

    public void Dispose() => _host.Dispose();

    static HttpResponseMessage Aula(HttpRequestMessage r, string? _) =>
        System.Web.HttpUtility.ParseQueryString(r.RequestUri!.Query)["method"] switch
        {
            "profiles.getProfilesByLogin" => FakeHandler.Json(Fixture.Read("profilesByLogin.json")),
            "profiles.getProfileContext" => FakeHandler.Json(Fixture.Read("profileContext.json")),
            _ => FakeHandler.Json(Fixture.Read("eventsEmpty.json")),
        };

    [Avalonia.Headless.XUnit.AvaloniaFact]
    public async Task Overlapping_cookie_checks_sign_in_once()
    {
        var session = new SessionController(_host.Sync, _host.Log, _ => new HttpClient(new FakeHandler(Aula)));
        var vm = new LoginViewModel(session);
        var signedIn = 0;
        vm.SignedIn += _ => signedIn++;
        var view = new LoginView(new BrowserProfile(_host.Paths, isMac: false), vm);
        var cookies = new TaskCompletionSource<IEnumerable<Cookie>>();

        var first = view.PollAsync(() => cookies.Task);
        var second = view.PollAsync(() => cookies.Task);
        cookies.SetResult([new("PHPSESSID", "abc", "/", ".aula.dk"), new("Csrfp-Token", "tok123", "/", "www.aula.dk")]);
        await Task.WhenAll(first, second);

        Assert.Equal(1, signedIn);
    }

    // Windows uden WebView2 Runtime (fx visse skole-pc'er): Avalonia ville falde tilbage til den gamle EdgeHTML, som ikke kan
    // give cookies, så login blev aldrig færdigt. I stedet siger login-visningen, hvad der mangler.
    [Avalonia.Headless.XUnit.AvaloniaFact]
    public void Without_webview2_login_says_what_to_install()
    {
        var session = new SessionController(_host.Sync, _host.Log, _ => new HttpClient(new FakeHandler(Aula)));
        var view = new LoginView(new BrowserProfile(_host.Paths, isMac: false), new LoginViewModel(session), runtimeMissing: true);
        new Avalonia.Controls.Window { Content = view }.Show();
        var visuals = Avalonia.VisualTree.VisualExtensions.GetVisualDescendants(view).ToList();
        Assert.Empty(visuals.OfType<Avalonia.Controls.NativeWebView>());
        Assert.Contains(visuals.OfType<Avalonia.Controls.TextBlock>(), t => t.Text == LoginView.MissingRuntimeText && t.IsEffectivelyVisible);
        Assert.Contains("https://go.microsoft.com/fwlink/p/?LinkId=2124703", LoginView.MissingRuntimeText);
    }

    // To loginvisninger på samme Aula-session (det skjulte genlogin og et synligt loginvindue): ét login, ikke to.
    [Avalonia.Headless.XUnit.AvaloniaFact]
    public async Task Two_login_views_with_the_same_session_sign_in_once()
    {
        var gate = new TaskCompletionSource();
        var session = new SessionController(_host.Sync, _host.Log, _ => new HttpClient(new GatedHandler(gate.Task, Aula)));
        var changed = 0;
        session.Changed += () => changed++;
        LoginView View() => new(new BrowserProfile(_host.Paths, isMac: false), new LoginViewModel(session));
        IEnumerable<Cookie> cookies = [new("PHPSESSID", "abc", "/", ".aula.dk"), new("Csrfp-Token", "tok123", "/", "www.aula.dk")];

        var hidden = View().PollAsync(() => Task.FromResult(cookies));
        var visible = View().PollAsync(() => Task.FromResult(cookies));
        gate.SetResult();
        await Task.WhenAll(hidden, visible);
        await View().PollAsync(() => Task.FromResult(cookies)); // også efter et gennemført login

        Assert.Equal(1, changed);
        Assert.Single(File.ReadAllLines(_host.Paths.Log), l => l.Contains("Logget ind"));
    }

    // En ny Aula-session (andet PHPSESSID) logger ind igen.
    [Avalonia.Headless.XUnit.AvaloniaFact]
    public async Task A_new_session_signs_in_again()
    {
        var session = new SessionController(_host.Sync, _host.Log, _ => new HttpClient(new FakeHandler(Aula)));
        var changed = 0;
        session.Changed += () => changed++;
        await session.SignInAsync([new("PHPSESSID", "abc", "/", ".aula.dk"), new("Csrfp-Token", "t", "/", "www.aula.dk")], default);
        await session.SignInAsync([new("PHPSESSID", "def", "/", ".aula.dk"), new("Csrfp-Token", "t", "/", "www.aula.dk")], default);
        Assert.Equal(2, changed);
    }
}

// Svarer først, når gate er åbnet (et login, der er i gang).
sealed class GatedHandler(Task gate, Func<HttpRequestMessage, string?, HttpResponseMessage> respond) : HttpMessageHandler
{
    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
    {
        await gate.WaitAsync(ct);
        return respond(request, request.Content is null ? null : await request.Content.ReadAsStringAsync(ct));
    }
}
