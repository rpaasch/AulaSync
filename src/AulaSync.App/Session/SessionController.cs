using System.Net;
using AulaSync.Core;

namespace AulaSync.App;

// Fra login-cookies til en kørende Aula-forbindelse: AulaSession → AulaClient.ConnectAsync → SyncService.
public sealed class SessionController
{
    readonly SyncService _sync;
    readonly FileLog _log;
    readonly Func<AulaSession, HttpClient> _createHttp;
    readonly string _baseUrl;

    public SessionController(SyncService sync, FileLog log, Func<AulaSession, HttpClient>? createHttp = null, string baseUrl = AulaClient.DefaultBaseUrl)
    {
        _sync = sync;
        _log = log;
        _createHttp = createHttp ?? (session => session.CreateHttpClient());
        _baseUrl = baseUrl;
    }

    // Den indloggede medarbejder. Bevares ved udløbet session, så navnet stadig kan vises i Indstillinger.
    public Profile? Profile { get; private set; }

    // Medarbejdere, klasser og lokaler. Hentes første gang, de skal bruges, og caches til log ud.
    public ScheduleCatalog? Catalog { get; private set; }

    public ScheduleRef? OwnSchedule =>
        Profile is { } p && ScheduleRef.IsValidId(p.Id) ? new ScheduleRef(ScheduleKind.Employee, p.Id, p.Name, p.Initials, p.Role) : null;

    public event Action? Changed;

    // Samme Aula-session (PHPSESSID) to gange — fx det skjulte genlogin og et synligt loginvindue på samme tid — giver
    // ét login: mens det er i gang, eller når det er gennemført og stadig aktivt, får kalderen det samme resultat.
    public Task<Profile> SignInAsync(IEnumerable<Cookie> cookies, CancellationToken ct)
    {
        var list = cookies.ToList();
        var session = list.FirstOrDefault(c => c.Name == AulaSession.SessionCookie)?.Value;
        if (session is not null && session == _signInSession && _signIn is { } current
            && (!current.IsCompleted || current.IsCompletedSuccessfully && _sync.Status.LoggedIn))
            return current;
        _signInSession = session;
        return _signIn = SignInCoreAsync(list, ct);
    }

    Task<Profile>? _signIn;
    string? _signInSession;

    async Task<Profile> SignInCoreAsync(IReadOnlyList<Cookie> cookies, CancellationToken ct)
    {
        var client = new AulaClient(_createHttp(new AulaSession(cookies)), _baseUrl, warn: _log.Info);
        var profile = await client.ConnectAsync(ct);
        Profile = profile;
        // Rollerne på valgte skemaer følger Aula: egen rolle nu, kollegernes når listen over medarbejdere hentes.
        Catalog = new ScheduleCatalog(client, all => _ = _sync.UpdateRolesAsync(all));
        if (OwnSchedule is { } own) await _sync.UpdateRolesAsync([own]);
        _sync.SetClient(client);
        _log.Info($"Logget ind (institution {profile.InstitutionCode})");
        Changed?.Invoke();
        return profile;
    }

    // Log ud: browserprofil, cachede lister og valg slettes; .ics-filerne bliver liggende (spec §3.3).
    public async Task SignOutAsync(BrowserProfile browser)
    {
        _sync.SetClient(null);
        _signIn = null;
        _signInSession = null;
        Profile = null;
        Catalog = null;
        await _sync.ClearAllAsync();
        browser.Reset();
        _log.Info("Logget ud");
        Changed?.Invoke();
    }
}
