using System.Net;
using AulaSync.Core;

namespace AulaSync.Tests;

public class AulaSessionTests
{
    static readonly Cookie[] Valid =
    [
        new("PHPSESSID", "abc", "/", ".aula.dk"),
        new("Csrfp-Token", "tok123", "/", "www.aula.dk"),
    ];

    [Fact]
    public void Requires_both_cookies()
    {
        Assert.True(AulaSession.HasRequiredCookies(Valid));
        Assert.False(AulaSession.HasRequiredCookies(Valid.Take(1)));
        Assert.False(AulaSession.HasRequiredCookies([new("PHPSESSID", "", "/", ".aula.dk"), Valid[1]]));
        Assert.Throws<ArgumentException>(() => new AulaSession(Valid.Take(1)));
    }

    [Fact]
    public async Task Http_client_sends_csrf_header_and_accept()
    {
        var fake = new FakeHandler((_, _) => FakeHandler.Json("{}"));
        using var http = new AulaSession(Valid).CreateHttpClient(inner: fake);
        await http.GetAsync("https://www.aula.dk/api/v23?method=x");
        var req = fake.Requests.Single().Request;
        Assert.Equal("tok123", req.Headers.GetValues("csrfp-token").Single());
        Assert.Contains("application/json", req.Headers.Accept.ToString());
    }
}
