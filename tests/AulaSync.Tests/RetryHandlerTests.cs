using System.Net;
using AulaSync.Core;

namespace AulaSync.Tests;

public class RetryHandlerTests
{
    static readonly Func<TimeSpan, CancellationToken, Task> NoDelay = (_, _) => Task.CompletedTask;

    [Fact]
    public async Task Retries_server_errors_then_succeeds()
    {
        int calls = 0;
        var fake = new FakeHandler((_, _) => ++calls < 3 ? new HttpResponseMessage(HttpStatusCode.ServiceUnavailable) : FakeHandler.Json("{}"));
        using var http = new HttpClient(new RetryHandler(fake, NoDelay));
        var resp = await http.GetAsync("https://x/");
        Assert.Equal(HttpStatusCode.OK, resp.StatusCode);
        Assert.Equal(3, calls);
    }

    [Fact]
    public async Task Gives_up_after_three_attempts()
    {
        int calls = 0;
        var fake = new FakeHandler((_, _) => { calls++; return new HttpResponseMessage(HttpStatusCode.BadGateway); });
        using var http = new HttpClient(new RetryHandler(fake, NoDelay));
        Assert.Equal(HttpStatusCode.BadGateway, (await http.GetAsync("https://x/")).StatusCode);
        Assert.Equal(3, calls);
    }

    [Fact]
    public async Task Does_not_retry_client_errors()
    {
        int calls = 0;
        var fake = new FakeHandler((_, _) => { calls++; return new HttpResponseMessage(HttpStatusCode.Forbidden); });
        using var http = new HttpClient(new RetryHandler(fake, NoDelay));
        await http.GetAsync("https://x/");
        Assert.Equal(1, calls);
    }

    [Fact]
    public async Task Retries_network_errors_and_resends_post_body()
    {
        int calls = 0;
        var fake = new FakeHandler((_, _) => ++calls == 1 ? throw new HttpRequestException("netværk") : FakeHandler.Json("{}"));
        using var http = new HttpClient(new RetryHandler(fake, NoDelay));
        await http.PostAsync("https://x/", new StringContent("{\"a\":1}"));
        Assert.Equal(2, calls);
        Assert.All(fake.Requests, r => Assert.Equal("{\"a\":1}", r.Body));
    }
}
