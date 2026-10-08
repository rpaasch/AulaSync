using System.Net;
using AulaSync.Core;

namespace AulaSync.Tests;

sealed class TempDir : IDisposable
{
    public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "aulasync-test-" + Guid.NewGuid().ToString("N"));
    public TempDir() => Directory.CreateDirectory(Path);
    public string File(string name) => System.IO.Path.Combine(Path, name);
    public void Dispose() { try { Directory.Delete(Path, true); } catch (IOException) { } }
}

static class Fixture
{
    public static string Read(string name) => File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));
}

sealed class FakeHandler(Func<HttpRequestMessage, string?, HttpResponseMessage> respond) : HttpMessageHandler
{
    public List<(HttpRequestMessage Request, string? Body)> Requests { get; } = [];

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
    {
        var body = request.Content is null ? null : await request.Content.ReadAsStringAsync(ct);
        Requests.Add((request, body));
        return respond(request, body);
    }

    public static HttpResponseMessage Json(string json, HttpStatusCode code = HttpStatusCode.OK) =>
        new(code) { Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json") };
}

sealed class FakeAulaClient : IAulaClient
{
    public Func<ScheduleRef, DateOnly, DateOnly, IReadOnlyList<AulaEvent>> Events { get; set; } = (_, _, _) => [];
    public Action OnPing { get; set; } = () => { };
    public int EventCalls;
    public int PingCalls;

    public Task<IReadOnlyList<AulaEvent>> GetEventsAsync(ScheduleRef schedule, DateOnly from, DateOnly to, CancellationToken ct)
    {
        Interlocked.Increment(ref EventCalls);
        return Task.FromResult(Events(schedule, from, to));
    }

    public Task PingAsync(CancellationToken ct)
    {
        Interlocked.Increment(ref PingCalls);
        OnPing();
        return Task.CompletedTask;
    }
}

static class TestData
{
    public static readonly TimeZoneInfo Copenhagen = TimeZoneInfo.FindSystemTimeZoneById("Europe/Copenhagen");

    public static AulaEvent Ev(string title = "DAN", string location = "53", string[]? groups = null,
        (string Name, string Initials)[]? teachers = null, bool substitute = false, (string, string)? substituteFor = null,
        string id = "1", string start = "2026-04-13T08:00:00+02:00", string end = "2026-04-13T08:45:00+02:00") =>
        new(id, title, DateTimeOffset.Parse(start), DateTimeOffset.Parse(end), location,
            groups ?? [], (teachers ?? []).Select(t => new Participant(t.Name, t.Initials)).ToList(),
            substitute, substituteFor is { } s ? new Participant(s.Item1, s.Item2) : null);

    public static async Task WaitUntil(Func<bool> condition)
    {
        for (int i = 0; i < 300 && !condition(); i++) await Task.Delay(10);
        Assert.True(condition(), "Betingelsen blev ikke opfyldt inden for 3 sekunder");
    }
}
