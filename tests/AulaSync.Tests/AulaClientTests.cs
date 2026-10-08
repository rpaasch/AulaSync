using System.Net;
using AulaSync.Core;
using static AulaSync.Tests.TestData;

namespace AulaSync.Tests;

public class AulaClientTests
{
    static readonly Func<TimeSpan, CancellationToken, Task> NoDelay = (_, _) => Task.CompletedTask;

    static (AulaClient Client, FakeHandler Fake) Create(Func<HttpRequestMessage, string?, HttpResponseMessage> respond, List<string>? warnings = null)
    {
        var fake = new FakeHandler(respond);
        return (new AulaClient(new HttpClient(fake), "https://www.aula.dk/api", Copenhagen, NoDelay, warnings is null ? null : warnings.Add), fake);
    }

    static string Method(HttpRequestMessage r) =>
        System.Web.HttpUtility.ParseQueryString(r.RequestUri!.Query)["method"] ?? "";

    static HttpResponseMessage Connected(HttpRequestMessage r, string? _)
    {
        if (r.RequestUri!.AbsolutePath.EndsWith("/v23")) return new HttpResponseMessage(HttpStatusCode.Gone);
        return Method(r) switch
        {
            "profiles.getProfilesByLogin" => FakeHandler.Json(Fixture.Read("profilesByLogin.json")),
            "profiles.getProfileContext" => FakeHandler.Json(Fixture.Read("profileContext.json")),
            "calendar.getEventsByProfileIdsAndResourceIds" or "calendar.geteventsbygroupid" => FakeHandler.Json(Fixture.Read("events.json")),
            "search.findProfilesAndGroups" when Uri.UnescapeDataString(r.RequestUri.Query).Contains("portalRoles[]=employee") => FakeHandler.Json(Fixture.Read("searchEmployees.json")),
            "search.findProfilesAndGroups" => FakeHandler.Json(Fixture.Read("searchGroups.json")),
            "profiles.getProfileMasterData" => FakeHandler.Json(Fixture.Read("masterData.json")),
            "resources.listResources" => FakeHandler.Json(Fixture.Read("resources.json")),
            _ => new HttpResponseMessage(HttpStatusCode.NotFound),
        };
    }

    [Fact]
    public async Task Connect_skips_gone_versions_and_reads_profile()
    {
        var (client, fake) = Create(Connected);
        var profile = await client.ConnectAsync(default);
        Assert.Equal("101001", profile.InstitutionCode);
        Assert.Equal("Test Bruger", profile.Name);
        Assert.Equal("TB", profile.Initials);
        Assert.EndsWith("/v24", fake.Requests.Last().Request.RequestUri!.AbsolutePath);
    }

    [Fact]
    public async Task Connect_with_expired_session_throws()
    {
        var (client, _) = Create((_, _) => FakeHandler.Json(Fixture.Read("sessionExpired.json")));
        await Assert.ThrowsAsync<SessionExpiredException>(() => client.ConnectAsync(default));
    }

    [Fact]
    public async Task Connect_with_no_working_version_throws()
    {
        var (client, _) = Create((_, _) => new HttpResponseMessage(HttpStatusCode.Gone));
        var ex = await Assert.ThrowsAsync<AulaException>(() => client.ConnectAsync(default));
        Assert.IsNotType<SessionExpiredException>(ex);
    }

    [Fact]
    public async Task Employee_events_are_posted_with_profile_id_and_local_dates()
    {
        var (client, fake) = Create(Connected);
        await client.ConnectAsync(default);
        var events = await client.GetEventsAsync(new ScheduleRef(ScheduleKind.Employee, "1001", "Anna"), new DateOnly(2026, 4, 13), new DateOnly(2026, 5, 24), default);
        Assert.Equal(3, events.Count);
        var (req, body) = fake.Requests.Last();
        Assert.Equal(HttpMethod.Post, req.Method);
        Assert.Equal("calendar.getEventsByProfileIdsAndResourceIds", Method(req));
        Assert.Contains("\"instProfileIds\":[1001]", body);
        Assert.Contains("\"resourceIds\":[]", body);
        Assert.Contains("\"start\":\"2026-04-13 00:00:00.0000+02:00\"", body);
        Assert.Contains("\"end\":\"2026-05-24 23:59:59.9990+02:00\"", body);
    }

    [Fact]
    public async Task Resource_events_are_posted_with_resource_id()
    {
        var (client, fake) = Create(Connected);
        await client.ConnectAsync(default);
        await client.GetEventsAsync(new ScheduleRef(ScheduleKind.Resource, "412", "53"), new DateOnly(2026, 4, 13), new DateOnly(2026, 4, 20), default);
        var body = fake.Requests.Last().Body;
        Assert.Contains("\"instProfileIds\":[]", body);
        Assert.Contains("\"resourceIds\":[412]", body);
    }

    [Fact]
    public async Task Group_events_use_get_with_escaped_plus()
    {
        var (client, fake) = Create(Connected);
        await client.ConnectAsync(default);
        await client.GetEventsAsync(new ScheduleRef(ScheduleKind.Group, "88231", "7A"), new DateOnly(2026, 4, 13), new DateOnly(2026, 4, 20), default);
        var uri = fake.Requests.Last().Request.RequestUri!.OriginalString;
        Assert.Contains("method=calendar.geteventsbygroupid", uri);
        Assert.Contains("groupId=88231", uri);
        Assert.Contains("start=2026-04-13T00:00:00.0000%2B02:00", uri);
        Assert.Contains("end=2026-04-20T23:59:59.9990%2B02:00", uri);
    }

    static Task<IReadOnlyList<AulaEvent>> GetAnna(AulaClient client) =>
        client.GetEventsAsync(new ScheduleRef(ScheduleKind.Employee, "1001", "H"), new DateOnly(2026, 4, 13), new DateOnly(2026, 4, 20), default);

    static Func<HttpRequestMessage, string?, HttpResponseMessage> EventsAnswer(Func<HttpResponseMessage> answer) =>
        (r, b) => Method(r) == "calendar.getEventsByProfileIdsAndResourceIds" ? answer() : Connected(r, b);

    [Fact]
    public async Task Http_403_means_no_access_not_expired()
    {
        var (client, _) = Create(EventsAnswer(() => new HttpResponseMessage(HttpStatusCode.Forbidden)));
        await client.ConnectAsync(default);
        var ex = await Assert.ThrowsAsync<ForbiddenException>(() => GetAnna(client));
        Assert.IsNotType<SessionExpiredException>(ex);
    }

    // Uden Aula-kode i svaret: HTTP 401 og 448 betyder udløbet login.
    [Theory]
    [InlineData(401)]
    [InlineData(448)]
    public async Task Bare_http_401_or_448_means_session_expired(int status)
    {
        var (client, _) = Create(EventsAnswer(() => new HttpResponseMessage((HttpStatusCode)status)));
        await client.ConnectAsync(default);
        await Assert.ThrowsAsync<SessionExpiredException>(() => GetAnna(client));
    }

    // Aulas kode 401 i svaret betyder "ingen tilladelse" og vinder over HTTP 401.
    [Fact]
    public async Task Http_401_with_code_401_in_body_means_no_access()
    {
        var (client, _) = Create(EventsAnswer(() => FakeHandler.Json("""{"status":{"code":401,"subCode":1,"message":"x"},"data":null}""", HttpStatusCode.Unauthorized)));
        await client.ConnectAsync(default);
        await Assert.ThrowsAsync<ForbiddenException>(() => GetAnna(client));
    }

    [Fact]
    public async Task Http_403_with_html_body_means_no_access()
    {
        var (client, _) = Create(EventsAnswer(() => new HttpResponseMessage(HttpStatusCode.Forbidden) { Content = new StringContent("<html>Forbidden</html>") }));
        await client.ConnectAsync(default);
        await Assert.ThrowsAsync<ForbiddenException>(() => GetAnna(client));
    }

    static Func<HttpRequestMessage, string?, HttpResponseMessage> LoginAnswer(Func<HttpResponseMessage> answer) =>
        (r, b) => Method(r) == "profiles.getProfilesByLogin" ? answer() : Connected(r, b);

    // Ved login og keep-alive er 401/403 på brugerens egen profil et ugyldigt login.
    [Theory]
    [InlineData(401)]
    [InlineData(403)]
    [InlineData(448)]
    public async Task Connect_with_rejected_own_profile_means_session_expired(int status)
    {
        var (client, _) = Create(LoginAnswer(() => new HttpResponseMessage((HttpStatusCode)status)));
        await Assert.ThrowsAsync<SessionExpiredException>(() => client.ConnectAsync(default));
    }

    [Theory]
    [InlineData("""{"status":{"code":403,"message":"Forbidden"},"data":null}""", HttpStatusCode.OK)]
    [InlineData("""{"status":{"code":20,"message":"x"},"data":null}""", HttpStatusCode.Gone)]
    [InlineData("""{"status":{"code":0,"subCode":23,"message":"x"},"data":null}""", HttpStatusCode.Gone)]
    public async Task Connect_with_session_or_access_code_in_body_means_session_expired(string body, HttpStatusCode status)
    {
        var (client, _) = Create(LoginAnswer(() => FakeHandler.Json(body, status)));
        await Assert.ThrowsAsync<SessionExpiredException>(() => client.ConnectAsync(default));
    }

    [Theory]
    [InlineData(451, "Du har ikke fået adgang til Aula endnu")]
    [InlineData(452, "Din Aula-bruger er deaktiveret")]
    public async Task Connect_with_account_access_code_stops_with_clear_text(int code, string text)
    {
        var (client, fake) = Create(LoginAnswer(() => FakeHandler.Json($$"""{"status":{"code":{{code}},"message":"x"},"data":null}""", (HttpStatusCode)code)));
        var ex = await Assert.ThrowsAsync<AccessDeniedException>(() => client.ConnectAsync(default));
        Assert.Equal(text, ex.Message);
        Assert.Single(fake.Requests); // ingen nye versionsforsøg
    }

    // Kun 404/410 betyder "prøv næste API-version"; en serverfejl stopper med det samme.
    [Fact]
    public async Task Connect_stops_on_server_error_instead_of_trying_next_version()
    {
        var (client, fake) = Create((_, _) => new HttpResponseMessage(HttpStatusCode.ServiceUnavailable));
        var ex = await Assert.ThrowsAsync<AulaException>(() => client.ConnectAsync(default));
        Assert.Contains("HTTP 503", ex.Message);
        Assert.Single(fake.Requests);
    }

    [Fact]
    public async Task Connect_skips_missing_versions()
    {
        var (client, fake) = Create((r, b) => r.RequestUri!.AbsolutePath.EndsWith("/v23") ? new HttpResponseMessage(HttpStatusCode.NotFound) : Connected(r, b));
        await client.ConnectAsync(default);
        Assert.EndsWith("/v24", fake.Requests.Last().Request.RequestUri!.AbsolutePath);
    }

    // 403 på getProfileContext må ikke stille give en forkert profil.
    [Fact]
    public async Task Connect_with_rejected_profile_context_does_not_fall_back()
    {
        var (client, _) = Create((r, b) => Method(r) == "profiles.getProfileContext" ? new HttpResponseMessage(HttpStatusCode.Forbidden) : Connected(r, b));
        var ex = await Assert.ThrowsAsync<AulaException>(() => client.ConnectAsync(default));
        Assert.Equal("Aula gav ikke adgang til din medarbejderprofil", ex.Message);
        Assert.Null(client.Profile);
    }

    [Theory]
    [InlineData(HttpStatusCode.Forbidden, "")]
    [InlineData(HttpStatusCode.OK, """{"status":{"code":403,"message":"Forbidden"},"data":null}""")]
    public async Task Ping_with_rejected_own_profile_means_session_expired(HttpStatusCode status, string body)
    {
        var rejected = false;
        var (client, _) = Create((r, b) => rejected ? new HttpResponseMessage(status) { Content = new StringContent(body) } : Connected(r, b));
        await client.ConnectAsync(default);
        rejected = true;
        await Assert.ThrowsAsync<SessionExpiredException>(() => client.PingAsync(default));
    }

    // Aulas egne koder i svaret vinder over HTTP-statussen.
    [Fact]
    public async Task Http_error_with_session_code_in_body_means_expired()
    {
        var (client, _) = Create(EventsAnswer(() => FakeHandler.Json("""{"status":{"code":0,"subCode":23,"message":"x"},"data":null}""", HttpStatusCode.Forbidden)));
        await client.ConnectAsync(default);
        await Assert.ThrowsAsync<SessionExpiredException>(() => GetAnna(client));
    }

    [Fact]
    public async Task Ok_http_with_403_in_body_means_no_access()
    {
        var (client, _) = Create(EventsAnswer(() => FakeHandler.Json(Fixture.Read("forbidden.json"))));
        await client.ConnectAsync(default);
        await Assert.ThrowsAsync<ForbiddenException>(() => GetAnna(client));
    }

    [Fact]
    public async Task Ping_ok_and_expired()
    {
        var expired = false;
        var (client, _) = Create((r, b) => expired ? FakeHandler.Json(Fixture.Read("sessionExpired.json")) : Connected(r, b));
        await client.ConnectAsync(default);
        await client.PingAsync(default);
        expired = true;
        await Assert.ThrowsAsync<SessionExpiredException>(() => client.PingAsync(default));
    }

    [Fact]
    public async Task Employees_are_resolved_filtered_and_sorted()
    {
        var (client, _) = Create(Connected);
        await client.ConnectAsync(default);
        var list = await client.GetEmployeesAsync(default);
        // Masterdata erstatter søgeresultaterne i batchen: Anna (fuldt navn) og Lene (leder); Per (technical) filtreres fra.
        Assert.Equal(["Anna Bech Eksempel", "Lene Leder"], list.Select(e => e.Name));
    }

    [Fact]
    public async Task Employees_fall_back_to_search_data_when_master_data_fails()
    {
        var (client, _) = Create((r, b) => Method(r) == "profiles.getProfileMasterData" ? FakeHandler.Json(Fixture.Read("errorStatus.json")) : Connected(r, b));
        await client.ConnectAsync(default);
        var list = await client.GetEmployeesAsync(default);
        Assert.Equal(["Anna Eksempel", "Kim Prøve"], list.Select(e => e.Name));
    }

    // Aula ignorerer portalRole (ental) og fylder så loftet med elever og forældre; portalRoles[] filtrerer på serveren.
    [Fact]
    public async Task Employee_search_filters_on_server_with_portalRoles()
    {
        var (client, fake) = Create(Connected);
        await client.ConnectAsync(default);
        await client.GetEmployeesAsync(default);
        var searches = fake.Requests.Select(x => Uri.UnescapeDataString(x.Request.RequestUri!.Query)).Where(q => q.Contains("search.findProfilesAndGroups")).ToList();
        Assert.NotEmpty(searches);
        Assert.All(searches, q => { Assert.Contains("portalRoles[]=employee", q); Assert.DoesNotContain("portalRole=", q); });
    }

    static string Text(HttpRequestMessage r) =>
        System.Web.HttpUtility.ParseQueryString(r.RequestUri!.Query)["text"] ?? "";

    // Loftet er hævet fra 100 (spiken: limit=1000 overholdes). Rammer et bogstav alligevel loftet, kan medarbejdere mangle.
    [Fact]
    public async Task Employee_search_warns_when_a_letter_hits_the_limit()
    {
        var full = "{\"status\":{\"code\":0,\"message\":\"OK\"},\"data\":{\"results\":[" +
            string.Join(",", Enumerable.Range(1, AulaClient.SearchLimit).Select(i => $"{{\"id\":\"{3000 + i}\",\"portalRole\":\"employee\",\"name\":\"Ansat {i}\",\"institutionRole\":\"teacher\"}}")) +
            "]}}";
        var warnings = new List<string>();
        var (client, fake) = Create((r, b) => Method(r) == "search.findProfilesAndGroups" && Text(r) == "a" ? FakeHandler.Json(full) : Connected(r, b), warnings);
        await client.ConnectAsync(default);
        await client.GetEmployeesAsync(default);
        var searches = fake.Requests.Select(x => Uri.UnescapeDataString(x.Request.RequestUri!.Query)).Where(q => q.Contains("search.findProfilesAndGroups")).ToList();
        Assert.All(searches, q => Assert.Contains($"limit={AulaClient.SearchLimit}&", q));
        Assert.Equal([$"Medarbejdersøgningen på 'a' ramte loftet på {AulaClient.SearchLimit} resultater — nogle medarbejdere kan mangle"], warnings);
    }

    static int Calls(FakeHandler fake, string method) => fake.Requests.Count(x => Method(x.Request) == method);

    // Klasser: findGroups med * giver alle grupper i ét kald.
    [Fact]
    public async Task Groups_come_from_one_wildcard_call()
    {
        var (client, fake) = Create((r, b) => Method(r) == "search.findGroups" ? FakeHandler.Json(Fixture.Read("findGroups.json")) : Connected(r, b));
        await client.ConnectAsync(default);
        Assert.Equal([new NamedItem("88233", "0B"), new NamedItem("88231", "7A")], await client.GetGroupsAsync(default));
        Assert.Equal(1, Calls(fake, "search.findGroups"));
        Assert.Equal(0, Calls(fake, "search.findProfilesAndGroups"));
        var query = Uri.UnescapeDataString(fake.Requests.Single(x => Method(x.Request) == "search.findGroups").Request.RequestUri!.Query);
        Assert.Contains("text=*", query);
        Assert.Contains("institutionCodes[]=101001", query);
    }

    [Fact]
    public async Task Groups_fall_back_to_letter_search_when_wildcard_fails()
    {
        var (client, fake) = Create((r, b) => Method(r) == "search.findGroups" ? FakeHandler.Json(Fixture.Read("errorStatus.json")) : Connected(r, b));
        await client.ConnectAsync(default);
        Assert.Equal([new NamedItem("88231", "7A")], await client.GetGroupsAsync(default));
        Assert.True(Calls(fake, "search.findProfilesAndGroups") > 1);
    }

    // Lokaler: query=* giver alle i ét kald; kun aktive lokaler (location/extra-location) beholdes.
    [Fact]
    public async Task Resources_come_from_one_wildcard_call()
    {
        var (client, fake) = Create((r, b) => Method(r) == "resources.listResources" && r.RequestUri!.Query.Contains("query=*")
            ? FakeHandler.Json(Fixture.Read("resourcesAll.json")) : Connected(r, b));
        await client.ConnectAsync(default);
        Assert.Equal([new NamedItem("412", "53"), new NamedItem("413", "Gymnastiksal"), new NamedItem("416", "Hal 2"), new NamedItem("417", "Musiklokale")],
            await client.GetResourcesAsync(default));
        Assert.Equal(1, Calls(fake, "resources.listResources"));
    }

    [Fact]
    public async Task Resources_fall_back_to_letter_search_when_wildcard_is_empty()
    {
        var (client, fake) = Create((r, b) => Method(r) == "resources.listResources" && r.RequestUri!.Query.Contains("query=*")
            ? FakeHandler.Json(Fixture.Read("eventsEmpty.json")) : Connected(r, b));
        await client.ConnectAsync(default);
        Assert.Equal(2, (await client.GetResourcesAsync(default)).Count);
        Assert.True(Calls(fake, "resources.listResources") > 1);
    }

    [Fact]
    public async Task Expired_session_during_wildcard_is_not_swallowed()
    {
        var (client, _) = Create((r, b) => Method(r) == "search.findGroups" ? FakeHandler.Json(Fixture.Read("sessionExpired.json")) : Connected(r, b));
        await client.ConnectAsync(default);
        await Assert.ThrowsAsync<SessionExpiredException>(() => client.GetGroupsAsync(default));
    }

    // En konto uden adgang (451/452) sendes videre; der søges ikke bogstav for bogstav.
    [Theory]
    [InlineData(451)]
    [InlineData(452)]
    public async Task Access_denied_during_wildcard_is_not_swallowed(int code)
    {
        var denied = $$"""{"status":{"code":{{code}},"subCode":0,"message":"x"},"data":null}""";
        var (client, fake) = Create((r, b) =>
            Method(r) == "search.findGroups" || Method(r) == "resources.listResources" && r.RequestUri!.Query.Contains("query=*")
                ? FakeHandler.Json(denied) : Connected(r, b));
        await client.ConnectAsync(default);
        await Assert.ThrowsAsync<AccessDeniedException>(() => client.GetGroupsAsync(default));
        await Assert.ThrowsAsync<AccessDeniedException>(() => client.GetResourcesAsync(default));
        Assert.Equal(0, Calls(fake, "search.findProfilesAndGroups"));
        Assert.Equal(1, Calls(fake, "resources.listResources"));
    }

    // HTTP 403 på *-kaldet: ingen adgang til netop det kald, så der søges bogstav for bogstav med én advarsel.
    [Fact]
    public async Task Forbidden_wildcard_falls_back_with_one_warning()
    {
        var warnings = new List<string>();
        var (client, _) = Create((r, b) =>
            Method(r) == "search.findGroups" || Method(r) == "resources.listResources" && r.RequestUri!.Query.Contains("query=*")
                ? new HttpResponseMessage(HttpStatusCode.Forbidden) : Connected(r, b), warnings);
        await client.ConnectAsync(default);
        Assert.Equal([new NamedItem("88231", "7A")], await client.GetGroupsAsync(default));
        Assert.Equal(2, (await client.GetResourcesAsync(default)).Count);
        Assert.Collection(warnings,
            w => Assert.StartsWith("Gruppelisten med * fejlede", w),
            w => Assert.StartsWith("Lokalelisten med * fejlede", w));
    }

    // Kender Aula ikke *-kaldene (HTTP 404), kommer grupper og lokaler fra bogstavsøgningen.
    [Fact]
    public async Task Groups_and_resources_fall_back_when_wildcard_is_unknown()
    {
        var (client, fake) = Create((r, b) => Method(r) == "resources.listResources" && r.RequestUri!.Query.Contains("query=*")
            ? new HttpResponseMessage(HttpStatusCode.NotFound) : Connected(r, b));
        await client.ConnectAsync(default);
        Assert.Equal([new NamedItem("88231", "7A")], await client.GetGroupsAsync(default));
        Assert.Equal(2, (await client.GetResourcesAsync(default)).Count);
        Assert.True(Calls(fake, "search.findProfilesAndGroups") > 1);
        Assert.True(Calls(fake, "resources.listResources") > 1);
    }
}
