using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace AulaSync.Core;

public interface IAulaClient
{
    Task<IReadOnlyList<AulaEvent>> GetEventsAsync(ScheduleRef schedule, DateOnly from, DateOnly to, CancellationToken ct);
    Task PingAsync(CancellationToken ct);
}

public interface IAulaDirectory
{
    Task<IReadOnlyList<Employee>> GetEmployeesAsync(CancellationToken ct);
    Task<IReadOnlyList<NamedItem>> GetGroupsAsync(CancellationToken ct);
    Task<IReadOnlyList<NamedItem>> GetResourcesAsync(CancellationToken ct);
}

public sealed class AulaClient : IAulaClient, IAulaDirectory
{
    public const string DefaultBaseUrl = "https://www.aula.dk/api";
    // Loft for medarbejdersøgningen. Aula overholder værdier over 100 (spiken: 151 resultater med limit=1000).
    public const int SearchLimit = 1000;
    const int FirstApiVersion = 23, LastApiVersion = 40;
    const string EmployeeSearchChars = "abcdefghijklmnopqrstuvwxyzæøå";
    const string GroupSearchChars = "abcdefghijklmnopqrstuvwxyz0123456789æøå";
    const string ResourceSearchChars = "abcdefghijklmnopqrstuvwxyz";
    static readonly string[] EmployeeRoles = ["teacher", "preschool-teacher", "leader", "other"];
    static readonly StringComparer DanishOrder = StringComparer.Create(CultureInfo.GetCultureInfo("da-DK"), ignoreCase: true);
    // Bevar "+" i tidszoner uescaped, som Aulas egen klient gør.
    static readonly JsonSerializerOptions JsonOptions = new() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    readonly HttpClient _http;
    readonly string _baseUrl;
    readonly TimeZoneInfo _tz;
    readonly Func<TimeSpan, CancellationToken, Task> _delay;
    readonly Action<string> _warn;
    string? _apiUrl;

    public AulaClient(HttpClient http, string baseUrl = DefaultBaseUrl, TimeZoneInfo? tz = null,
        Func<TimeSpan, CancellationToken, Task>? delay = null, Action<string>? warn = null)
    {
        _http = http;
        _baseUrl = baseUrl.TrimEnd('/');
        _tz = tz ?? TimeZoneInfo.Local;
        _delay = delay ?? Task.Delay;
        _warn = warn ?? (_ => { });
    }

    public Profile? Profile { get; private set; }

    string Api => _apiUrl ?? throw new InvalidOperationException("ConnectAsync er ikke kaldt");

    string Institution => Uri.EscapeDataString(Profile?.InstitutionCode ?? throw new InvalidOperationException("ConnectAsync er ikke kaldt"));

    public async Task<Profile> ConnectAsync(CancellationToken ct)
    {
        for (int version = FirstApiVersion; version <= LastApiVersion; version++)
        {
            var url = $"{_baseUrl}/v{version}";
            Profile? fallback;
            using (var response = await _http.GetAsync($"{url}?method=profiles.getProfilesByLogin", ct))
            {
                var body = await response.Content.ReadAsStringAsync(ct);
                var known = AulaParsers.KnownError(body);
                if (known is AccessDeniedException) throw known;
                // Ved login er 401/403 på brugerens egen profil et ugyldigt login — ikke manglende adgang til et skema.
                if (known is not null || response.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden || (int)response.StatusCode == 448)
                    throw new SessionExpiredException();
                if (response.StatusCode is HttpStatusCode.NotFound or HttpStatusCode.Gone) continue; // forældet eller ukendt API-version
                if (!response.IsSuccessStatusCode) throw new AulaException($"Aula svarede HTTP {(int)response.StatusCode}");
                try { fallback = AulaParsers.ParseFirstProfile(body); }
                catch (AulaException) { continue; }
            }

            _apiUrl = url;
            try { Profile = AulaParsers.ParseProfileContext(await GetAsync("profiles.getProfileContext&portalrole=employee", ct)); }
            catch (ForbiddenException) { throw new AulaException("Aula gav ikke adgang til din medarbejderprofil"); }
            catch (AulaException ex) when (ex is not (SessionExpiredException or AccessDeniedException))
            {
                _warn($"getProfileContext fejlede ({ex.Message}) — bruger første profil");
                Profile = fallback;
            }
            if (Profile is null || Profile.InstitutionCode == "") throw new AulaException("Kunne ikke finde din institution i Aula");
            return Profile;
        }
        throw new AulaException("Kunne ikke finde en Aula API-version, der svarer");
    }

    // 401/403 på brugerens egen profil er et ugyldigt login, ligesom ved ConnectAsync.
    public async Task PingAsync(CancellationToken ct)
    {
        try { AulaParsers.EnsureOk(await GetAsync("profiles.getProfilesByLogin", ct)); }
        catch (ForbiddenException) { throw new SessionExpiredException(); }
    }

    public async Task<IReadOnlyList<AulaEvent>> GetEventsAsync(ScheduleRef schedule, DateOnly from, DateOnly to, CancellationToken ct)
    {
        string body;
        if (schedule.Kind == ScheduleKind.Group)
        {
            body = await GetAsync(
                $"calendar.geteventsbygroupid&groupId={Uri.EscapeDataString(schedule.Id)}" +
                $"&start={Plus(AulaDates.Format(from, false, 'T', _tz))}" +
                $"&end={Plus(AulaDates.Format(to, true, 'T', _tz))}" +
                "&includeOnlyInvitedEvents=false", ct);
        }
        else
        {
            var id = long.Parse(schedule.Id, CultureInfo.InvariantCulture);
            body = await PostAsync("calendar.getEventsByProfileIdsAndResourceIds", new
            {
                instProfileIds = schedule.Kind == ScheduleKind.Employee ? new[] { id } : Array.Empty<long>(),
                resourceIds = schedule.Kind == ScheduleKind.Resource ? new[] { id } : Array.Empty<long>(),
                start = AulaDates.Format(from, false, ' ', _tz),
                end = AulaDates.Format(to, true, ' ', _tz),
            }, ct);
        }
        var result = AulaParsers.ParseEvents(body);
        if (result.Skipped > 0) _warn($"{schedule.FileName}: {result.Skipped} begivenheder uden gyldig tid blev sprunget over");
        return result.Events;
    }

    public async Task<IReadOnlyList<Employee>> GetEmployeesAsync(CancellationToken ct)
    {
        // portalRoles[] (flertal) filtrerer på serveren; portalRole (ental) ignoreres af Aula, som så fylder loftet med elever
        // og forældre (målt i spiken). Søgning med * giver ikke alle medarbejdere, så der søges bogstav for bogstav.
        var found = new Dictionary<string, Employee>();
        foreach (var c in EmployeeSearchChars)
        {
            await _delay(TimeSpan.FromMilliseconds(50), ct);
            var body = await GetAsync($"search.findProfilesAndGroups&text={Uri.EscapeDataString(c.ToString())}&instCodes[]={Institution}&typeahead=true&limit={SearchLimit}&portalRoles[]=employee", ct);
            foreach (var e in AulaParsers.ParseEmployeeSearch(body)) found.TryAdd(e.Id, e);
            if (AulaParsers.CountResults(body) >= SearchLimit)
                _warn($"Medarbejdersøgningen på '{c}' ramte loftet på {SearchLimit} resultater — nogle medarbejdere kan mangle");
        }

        // Søgningen giver ikke altid institutionsprofil-id; masterdata giver de rigtige id'er og fulde navne.
        var ids = found.Keys.ToList();
        var result = new Dictionary<string, Employee>();
        for (int i = 0; i < ids.Count; i += 20)
        {
            var batch = ids.Skip(i).Take(20).ToList();
            var query = string.Join("&", batch.Select(id => $"instProfileIds[]={Uri.EscapeDataString(id)}"));
            try
            {
                foreach (var e in AulaParsers.ParseProfileMasterData(await GetAsync($"profiles.getProfileMasterData&{query}&fromAdministration=false", ct)))
                    result.TryAdd(e.Id, e);
            }
            catch (AulaException ex) when (ex is not SessionExpiredException)
            {
                _warn($"Opslag af {batch.Count} medarbejdere fejlede ({ex.Message}) — bruger søgeresultatet");
                foreach (var id in batch) result.TryAdd(id, found[id]);
            }
        }

        return result.Values.Where(e => EmployeeRoles.Contains(e.Role)).OrderBy(e => e.Name, DanishOrder).ToList();
    }

    public async Task<IReadOnlyList<NamedItem>> GetGroupsAsync(CancellationToken ct)
    {
        // findGroups med * giver alle institutionens grupper i ét kald (målt i spiken); ellers bogstav for bogstav som før.
        var all = await TryWildcardAsync($"search.findGroups&text=*&institutionCodes[]={Institution}&limit=1000", AulaParsers.ParseGroupSearch, "Gruppelisten", ct);
        if (all.Count > 0) return all.OrderBy(g => g.Name, DanishOrder).ToList();

        var found = new Dictionary<string, NamedItem>();
        foreach (var c in GroupSearchChars)
        {
            await _delay(TimeSpan.FromMilliseconds(50), ct);
            try
            {
                var body = await GetAsync($"search.findProfilesAndGroups&text={Uri.EscapeDataString(c.ToString())}&instCodes[]={Institution}&typeahead=true&limit=100", ct);
                foreach (var g in AulaParsers.ParseGroupSearch(body)) found.TryAdd(g.Id, g);
            }
            catch (AulaException ex) when (ex is not SessionExpiredException) { _warn($"Gruppesøgning '{c}' fejlede: {ex.Message}"); }
        }
        return found.Values.OrderBy(g => g.Name, DanishOrder).ToList();
    }

    public async Task<IReadOnlyList<NamedItem>> GetResourcesAsync(CancellationToken ct)
    {
        // listResources med query=* giver alle ressourcer i ét kald (som Aulas planlægningsassistent); ellers bogstav for bogstav.
        var all = await TryWildcardAsync($"resources.listResources&query=*&institutionCodes[]={Institution}", AulaParsers.ParseResources, "Lokalelisten", ct);
        if (all.Count > 0) return all.OrderBy(r => r.Name, DanishOrder).ToList();

        var found = new Dictionary<string, NamedItem>();
        foreach (var c in ResourceSearchChars)
        {
            await _delay(TimeSpan.FromMilliseconds(50), ct);
            try
            {
                var body = await GetAsync($"resources.listResources&query={c}&institutionCodes[]={Institution}", ct);
                foreach (var r in AulaParsers.ParseResources(body)) found.TryAdd(r.Id, r);
            }
            catch (AulaException ex) when (ex is not SessionExpiredException) { _warn($"Lokalesøgning '{c}' fejlede: {ex.Message}"); }
        }
        return found.Values.OrderBy(r => r.Name, DanishOrder).ToList();
    }

    // Ét kald med *; en fejl eller et tomt svar giver en tom liste, så kalderen falder tilbage til bogstavsøgning.
    // Udløbet session og en konto uden adgang sendes videre.
    async Task<IReadOnlyList<NamedItem>> TryWildcardAsync(string methodAndQuery, Func<string, IReadOnlyList<NamedItem>> parse, string what, CancellationToken ct)
    {
        try { return parse(await GetAsync(methodAndQuery, ct)); }
        catch (AulaException ex) when (ex is not (SessionExpiredException or AccessDeniedException))
        {
            _warn($"{what} med * fejlede ({ex.Message}) — søger bogstav for bogstav");
            return [];
        }
    }

    async Task<string> GetAsync(string methodAndQuery, CancellationToken ct)
    {
        using var response = await _http.GetAsync($"{Api}?method={methodAndQuery}", ct);
        return await ReadAsync(response, ct);
    }

    async Task<string> PostAsync(string method, object body, CancellationToken ct)
    {
        using var content = new StringContent(JsonSerializer.Serialize(body, JsonOptions), Encoding.UTF8, "application/json");
        using var response = await _http.PostAsync($"{Api}?method={method}", content, ct);
        return await ReadAsync(response, ct);
    }

    // Ved HTTP-fejl læses svaret alligevel: Aulas egne koder (udløbet session, ingen adgang) vinder over HTTP-statussen.
    static async Task<string> ReadAsync(HttpResponseMessage response, CancellationToken ct)
    {
        var body = await response.Content.ReadAsStringAsync(ct);
        if (response.IsSuccessStatusCode) return body;
        if (AulaParsers.KnownError(body) is { } error) throw error;
        throw (int)response.StatusCode switch
        {
            401 or 448 => new SessionExpiredException(),
            403 => new ForbiddenException(),
            var code => new AulaException($"Aula svarede HTTP {code}"),
        };
    }

    static string Plus(string value) => value.Replace("+", "%2B");
}
