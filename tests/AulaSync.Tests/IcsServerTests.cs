using System.Net;
using System.Net.Sockets;
using AulaSync.Core;

namespace AulaSync.Tests;

public class IcsServerTests : IDisposable
{
    readonly TempDir _dir = new();
    readonly string _file;

    public IcsServerTests()
    {
        _file = _dir.File("lokale-412.ics");
        File.WriteAllText(_file, "BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n");
        File.WriteAllText(_dir.File("abonnementer.json"), "[]");
    }

    public void Dispose() => _dir.Dispose();

    [Fact]
    public void Serves_existing_calendar()
    {
        var d = IcsServer.Resolve(_dir.Path, "/lokale-412.ics", "GET", null);
        Assert.Equal(200, d.StatusCode);
        Assert.Equal(Path.Combine(_dir.Path, "lokale-412.ics"), d.FilePath);
        Assert.NotNull(d.LastModified);
    }

    // Filen kan hedde noget foran nøglen (CalendarFiles); adressen er stadig nøglen.
    [Fact]
    public void Serves_readable_file_name_at_the_key_address()
    {
        var readable = _dir.File("123456-7A-klasse-88231.ics");
        File.WriteAllText(readable, "BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n");
        var d = IcsServer.Resolve(_dir.Path, "/klasse-88231.ics", "GET", null);
        Assert.Equal(200, d.StatusCode);
        Assert.Equal(readable, d.FilePath);
        Assert.Equal(404, IcsServer.Resolve(_dir.Path, "/123456-7A-klasse-88231.ics", "GET", null).StatusCode);
    }

    [Theory]
    [InlineData("/lokale-999.ics")]
    [InlineData("/abonnementer.json")]
    [InlineData("/../lokale-412.ics")]
    [InlineData("/%2e%2e%2flokale-412.ics")]
    [InlineData("/kalendere/lokale-412.ics")]
    [InlineData("/LOKALE-412.ICS")]
    [InlineData("/")]
    public void Unknown_or_unsafe_paths_are_404(string path) =>
        Assert.Equal(404, IcsServer.Resolve(_dir.Path, path, "GET", null).StatusCode);

    [Fact]
    public void Only_get_and_head() =>
        Assert.Equal(405, IcsServer.Resolve(_dir.Path, "/lokale-412.ics", "POST", null).StatusCode);

    [Fact]
    public void Not_modified_since()
    {
        var lm = IcsServer.Resolve(_dir.Path, "/lokale-412.ics", "GET", null).LastModified!.Value;
        Assert.Equal(304, IcsServer.Resolve(_dir.Path, "/lokale-412.ics", "GET", lm).StatusCode);
        Assert.Equal(200, IcsServer.Resolve(_dir.Path, "/lokale-412.ics", "GET", lm.AddSeconds(-1)).StatusCode);
    }

    [Fact]
    public void Urls_use_the_users_port()
    {
        var s = new ScheduleRef(ScheduleKind.Resource, "412", "53");
        Assert.Equal("http://localhost:9876/lokale-412.ics", IcsServer.UrlFor(s, 9876));
        Assert.Equal("webcal://localhost:9877/lokale-412.ics", IcsServer.WebcalFor(s, 9877));
    }

    static int FreePort()
    {
        var l = new TcpListener(IPAddress.Loopback, 0);
        l.Start();
        var port = ((IPEndPoint)l.LocalEndpoint).Port;
        l.Stop();
        return port;
    }

    [Fact]
    public async Task Real_server_serves_get_and_head()
    {
        var port = FreePort();
        using var server = new IcsServer(_dir.Path, new FileLog(_dir.File("log.txt")), port);
        Assert.True(server.TryStart());
        using var http = new HttpClient();

        var get = await http.GetAsync($"http://localhost:{port}/lokale-412.ics");
        Assert.Equal(HttpStatusCode.OK, get.StatusCode);
        Assert.Equal("text/calendar", get.Content.Headers.ContentType!.MediaType);
        Assert.Equal("utf-8", get.Content.Headers.ContentType.CharSet);
        Assert.StartsWith("BEGIN:VCALENDAR", await get.Content.ReadAsStringAsync());
        Assert.NotNull(get.Content.Headers.LastModified);

        var head = await http.SendAsync(new HttpRequestMessage(HttpMethod.Head, $"http://localhost:{port}/lokale-412.ics"));
        Assert.Equal(HttpStatusCode.OK, head.StatusCode);

        var missing = await http.GetAsync($"http://localhost:{port}/abonnementer.json");
        Assert.Equal(HttpStatusCode.NotFound, missing.StatusCode);
    }

    // Apple Kalender kan slå localhost op som 127.0.0.1 eller ::1; serveren skal svare på begge.
    [Fact]
    public async Task Real_server_answers_on_both_loopback_addresses()
    {
        var port = FreePort();
        var log = _dir.File("log.txt");
        using var server = new IcsServer(_dir.Path, new FileLog(log), port);
        Assert.True(server.TryStart());
        using var http = new HttpClient();

        var v4 = await http.GetAsync($"http://127.0.0.1:{port}/lokale-412.ics");
        Assert.Equal(HttpStatusCode.OK, v4.StatusCode);
        Assert.StartsWith("BEGIN:VCALENDAR", await v4.Content.ReadAsStringAsync());
        Assert.Contains($"127.0.0.1:{port}", File.ReadAllText(log));

        if (!Ipv6LoopbackAvailable()) return; // fx en container uden IPv6
        var v6 = await http.GetAsync($"http://[::1]:{port}/lokale-412.ics");
        Assert.Equal(HttpStatusCode.OK, v6.StatusCode);
        Assert.Contains($"[::1]:{port}", File.ReadAllText(log));
    }

    static bool Ipv6LoopbackAvailable()
    {
        try
        {
            using var s = new Socket(AddressFamily.InterNetworkV6, SocketType.Stream, ProtocolType.Tcp);
            s.Bind(new IPEndPoint(IPAddress.IPv6Loopback, 0));
            return true;
        }
        catch (SocketException) { return false; }
    }

    // DNS-rebinding: en hjemmeside, hvis navn peger på 127.0.0.1, må ikke kunne læse skemaerne.
    [Theory]
    [InlineData("rebind.example:{0}")]
    [InlineData("localhost:1")]
    [InlineData("localhost.evil.example:{0}")]
    public async Task Foreign_host_is_rejected(string host)
    {
        var port = FreePort();
        using var server = new IcsServer(_dir.Path, new FileLog(_dir.File("log.txt")), port);
        Assert.True(server.TryStart());
        using var http = new HttpClient();
        var request = new HttpRequestMessage(HttpMethod.Get, $"http://127.0.0.1:{port}/lokale-412.ics");
        request.Headers.Host = string.Format(host, port);
        var response = await http.SendAsync(request);
        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.DoesNotContain("VCALENDAR", await response.Content.ReadAsStringAsync());
    }

    [Theory]
    [InlineData("localhost:{0}")]
    [InlineData("LOCALHOST:{0}")]
    [InlineData("127.0.0.1:{0}")]
    [InlineData("[::1]:{0}")]
    public async Task Loopback_host_names_are_accepted(string host)
    {
        var port = FreePort();
        using var server = new IcsServer(_dir.Path, new FileLog(_dir.File("log.txt")), port);
        Assert.True(server.TryStart());
        using var http = new HttpClient();
        var request = new HttpRequestMessage(HttpMethod.Get, $"http://127.0.0.1:{port}/lokale-412.ics?t=1");
        request.Headers.Host = string.Format(host, port);
        Assert.Equal(HttpStatusCode.OK, (await http.SendAsync(request)).StatusCode);
    }

    // Så man i loggen kan se, om kalenderprogrammet overhovedet nåede AulaSync, og hvad det fik.
    [Fact]
    public async Task Fetches_and_refusals_are_logged()
    {
        var port = FreePort();
        var log = _dir.File("log.txt");
        using var server = new IcsServer(_dir.Path, new FileLog(log), port);
        Assert.True(server.TryStart());
        using var http = new HttpClient();

        await http.GetAsync($"http://localhost:{port}/lokale-412.ics");
        await http.GetAsync($"http://localhost:{port}/lokale-412.ics");
        await http.GetAsync($"http://localhost:{port}/klasse-999.ics");
        var text = File.ReadAllText(log);
        Assert.Single(text.Split('\n'), l => l.Contains("GET /lokale-412.ics → 200")); // kun første hentning
        Assert.Contains("GET /klasse-999.ics → 404", text);
    }

    [Fact]
    public async Task Log_names_the_program_that_fetched()
    {
        var port = FreePort();
        var log = _dir.File("log.txt");
        using var server = new IcsServer(_dir.Path, new FileLog(log), port);
        Assert.True(server.TryStart());
        using var http = new HttpClient();
        http.DefaultRequestHeaders.UserAgent.ParseAdd("CalendarAgent/1000");
        await http.GetAsync($"http://localhost:{port}/lokale-412.ics");
        Assert.Contains("GET /lokale-412.ics → 200 (CalendarAgent/1000)", File.ReadAllText(log));
    }

    // Prøver et program https (krypteret) mod den almindelige http-port, skal loggen sige det, og programmet skal få en
    // almindelig TLS-afvisning (handshake_failure) med det samme. Så ser Kalender "serveren kan ikke https" og prøver http,
    // i stedet for at forbindelsen bare forsvinder.
    [Fact]
    public async Task Https_attempt_is_logged_and_refused_with_a_tls_alert()
    {
        var port = FreePort();
        var log = _dir.File("log.txt");
        using var server = new IcsServer(_dir.Path, new FileLog(log), port);
        Assert.True(server.TryStart());
        using var client = new TcpClient();
        await client.ConnectAsync(IPAddress.Loopback, port);
        var stream = client.GetStream();
        await stream.WriteAsync(new byte[] { 0x16, 0x03, 0x01, 0x00, 0x05, 0x01, 0x00, 0x00, 0x01, 0x03 });
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5)); // med det samme, ikke efter timeout
        var reply = new MemoryStream();
        await stream.CopyToAsync(reply, cts.Token);
        Assert.Equal([0x15, 0x03, 0x01, 0x00, 0x02, 0x02, 0x28], reply.ToArray()); // alert, fatal, handshake_failure
        Assert.Contains("https", File.ReadAllText(log));
    }

    [Theory]
    [InlineData("HALLO\r\n\r\n", "HALLO")]
    [InlineData("\r\n\r\n", "")]
    public async Task Garbage_request_is_logged_and_refused(string request, string logged)
    {
        var port = FreePort();
        var log = _dir.File("log.txt");
        using var server = new IcsServer(_dir.Path, new FileLog(log), port);
        Assert.True(server.TryStart());
        using var client = new TcpClient();
        await client.ConnectAsync(IPAddress.Loopback, port);
        var stream = client.GetStream();
        await stream.WriteAsync(System.Text.Encoding.ASCII.GetBytes(request));
        var reply = await new StreamReader(stream).ReadToEndAsync();
        Assert.StartsWith("HTTP/1.1 400", reply);
        Assert.Contains($"ugyldig forespørgsel \"{logged}\"", File.ReadAllText(log));
    }

    // Hver vellykket hentning (200 og 304) meldes, så AulaSync kan vise, at kalenderprogrammet har fået skemaet.
    [Fact]
    public async Task Successful_fetches_are_reported()
    {
        var port = FreePort();
        using var server = new IcsServer(_dir.Path, new FileLog(_dir.File("log.txt")), port);
        var fetched = new List<string>();
        server.Fetched += name => { lock (fetched) fetched.Add(name); };
        Assert.True(server.TryStart());
        using var http = new HttpClient();

        File.WriteAllText(_dir.File("123456-7A-klasse-88231.ics"), "BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n");
        await http.GetAsync($"http://localhost:{port}/klasse-88231.ics"); // adressens navn, ikke filens
        var first = await http.GetAsync($"http://localhost:{port}/lokale-412.ics");
        var again = new HttpRequestMessage(HttpMethod.Get, $"http://localhost:{port}/lokale-412.ics");
        again.Headers.IfModifiedSince = first.Content.Headers.LastModified;
        Assert.Equal(HttpStatusCode.NotModified, (await http.SendAsync(again)).StatusCode);
        await http.GetAsync($"http://localhost:{port}/klasse-999.ics");
        var foreign = new HttpRequestMessage(HttpMethod.Get, $"http://127.0.0.1:{port}/lokale-412.ics");
        foreign.Headers.Host = $"rebind.example:{port}";
        await http.SendAsync(foreign);

        lock (fetched) Assert.Equal(["klasse-88231.ics", "lokale-412.ics", "lokale-412.ics"], fetched);
    }

    [Fact]
    public async Task Not_modified_over_the_wire()
    {
        var port = FreePort();
        using var server = new IcsServer(_dir.Path, new FileLog(_dir.File("log.txt")), port);
        Assert.True(server.TryStart());
        using var http = new HttpClient();
        File.WriteAllText(_dir.File("123456-7A-klasse-88231.ics"), "BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n");
        await http.GetAsync($"http://localhost:{port}/klasse-88231.ics"); // adressens navn, ikke filens
        var first = await http.GetAsync($"http://localhost:{port}/lokale-412.ics");
        var again = new HttpRequestMessage(HttpMethod.Get, $"http://localhost:{port}/lokale-412.ics");
        again.Headers.IfModifiedSince = first.Content.Headers.LastModified;
        var second = await http.SendAsync(again);
        Assert.Equal(HttpStatusCode.NotModified, second.StatusCode);
        Assert.Empty(await second.Content.ReadAsByteArrayAsync());
    }

    [Fact]
    public async Task Oversized_request_is_refused()
    {
        var port = FreePort();
        using var server = new IcsServer(_dir.Path, new FileLog(_dir.File("log.txt")), port);
        Assert.True(server.TryStart());
        using var http = new HttpClient();
        var request = new HttpRequestMessage(HttpMethod.Get, $"http://localhost:{port}/lokale-412.ics");
        request.Headers.Add("X-Fyld", new string('x', 20_000));
        Assert.Equal(HttpStatusCode.RequestHeaderFieldsTooLarge, (await http.SendAsync(request)).StatusCode);
    }

    [Fact]
    public async Task Stops_listening_when_disposed()
    {
        var port = FreePort();
        var server = new IcsServer(_dir.Path, new FileLog(_dir.File("log.txt")), port);
        Assert.True(server.TryStart());
        server.Dispose();
        using var client = new TcpClient();
        await Assert.ThrowsAnyAsync<SocketException>(() => client.ConnectAsync(IPAddress.Loopback, port));
    }

    // ::1 er kun valgfri, når computeren ikke har IPv6. Optager et andet program [::1]:9876, ville Kalender (som prøver ::1
    // først) ende hos det, mens AulaSync sagde, at serveren kørte; så skal serveren i stedet ikke starte ("port optaget").
    [Theory]
    [InlineData(SocketError.AddressFamilyNotSupported, true)]
    [InlineData(SocketError.AddressNotAvailable, true)]
    [InlineData(SocketError.ProtocolNotSupported, true)]
    [InlineData(SocketError.AddressAlreadyInUse, false)]
    [InlineData(SocketError.AccessDenied, false)]
    public void Ipv6_may_be_missing_but_not_busy(SocketError error, bool missing) => Assert.Equal(missing, IcsServer.Ipv6Missing(error));

    [Fact]
    public void Busy_ipv6_loopback_stops_the_server()
    {
        if (!Ipv6LoopbackAvailable()) return; // fx en container uden IPv6; testen kører på Mac og Windows
        var port = FreePort();
        var other = new TcpListener(IPAddress.IPv6Loopback, port);
        other.Start();
        try
        {
            using var server = new IcsServer(_dir.Path, new FileLog(_dir.File("log.txt")), port);
            Assert.False(server.TryStart());
            var probe = new TcpListener(IPAddress.Loopback, port); // 127.0.0.1 er sluppet igen
            probe.Start();
            probe.Stop();
        }
        finally { other.Stop(); }
    }

    // Metoden kommer fra klienten; kun almindelige metoder (store bogstaver) godtages, og loggen får ingen kontroltegn.
    [Fact]
    public async Task Odd_method_is_refused_and_not_logged_raw()
    {
        var port = FreePort();
        var log = _dir.File("log.txt");
        using var server = new IcsServer(_dir.Path, new FileLog(log), port);
        Assert.True(server.TryStart());
        using var client = new TcpClient();
        await client.ConnectAsync(IPAddress.Loopback, port);
        var stream = client.GetStream();
        await stream.WriteAsync(System.Text.Encoding.ASCII.GetBytes($"GE\u001b[2KT /lokale-412.ics HTTP/1.1\r\nHost: localhost:{port}\r\n\r\n"));
        var reply = await new StreamReader(stream).ReadToEndAsync();
        Assert.StartsWith("HTTP/1.1 400", reply);
        var text = File.ReadAllText(log);
        Assert.Contains("ugyldig forespørgsel", text);
        Assert.DoesNotContain('\u001b', text);
    }

    // Svaret må ikke gå tabt, fordi klienten stadig sender, når serveren lukker (ulæste data giver ellers en TCP-reset).
    [Fact]
    public async Task Refusal_arrives_even_when_the_client_keeps_sending()
    {
        var port = FreePort();
        using var server = new IcsServer(_dir.Path, new FileLog(_dir.File("log.txt")), port);
        Assert.True(server.TryStart());
        using var client = new TcpClient();
        await client.ConnectAsync(IPAddress.Loopback, port);
        var stream = client.GetStream();
        var head = System.Text.Encoding.ASCII.GetBytes($"GET /lokale-412.ics HTTP/1.1\r\nHost: localhost:{port}\r\nX-Fyld: ");
        await stream.WriteAsync(head);
        await stream.WriteAsync(new byte[200_000].Select(_ => (byte)'x').ToArray()); // langt over grænsen på 8 KB
        await Task.Delay(300); // serveren har svaret og lukket, før svaret læses
        var buffer = new byte[64];
        var read = await stream.ReadAsync(buffer);
        Assert.StartsWith("HTTP/1.1 431", System.Text.Encoding.ASCII.GetString(buffer, 0, read));
    }

    [Fact]
    public void Second_server_on_same_port_fails_to_start()
    {
        var port = FreePort();
        var log = new FileLog(_dir.File("log.txt"));
        using var first = new IcsServer(_dir.Path, log, port);
        using var second = new IcsServer(_dir.Path, log, port);
        Assert.True(first.TryStart());
        Assert.False(second.TryStart());
    }
}

// Porten gælder for brugeren (config.json). Første gang vælges 9876 eller den første ledige op til 9899, så en anden bruger
// på samme computer (hurtigt brugerskift) eller en kørende AulaSync 2 ikke spærrer. Derefter bruges kun den gemte port,
// for kalenderprogrammerne abonnerer på adresser med den.
public class UserPortTests : IDisposable
{
    readonly TempDir _dir = new();
    readonly ConfigStore _config;
    readonly List<int> _tried = [];

    public UserPortTests() => _config = new ConfigStore(_dir.File("config.json"));

    public void Dispose() => _dir.Dispose();

    Func<int, string?> Free(params int[] free) => port =>
    {
        _tried.Add(port);
        return free.Contains(port) ? $"server {port}" : null;
    };

    [Fact]
    public void First_start_takes_9876_and_saves_it()
    {
        Assert.Equal("server 9876", UserPort.Start(_config, Free(9876, 9877)));
        Assert.Equal([9876], _tried);
        Assert.Equal(9876, _config.Load().Port);
    }

    [Fact]
    public void First_start_takes_the_next_free_port_when_9876_is_busy()
    {
        _config.Save(new AppConfig(CalendarApp.OutlookClassic, FirstRunDone: true));
        Assert.Equal("server 9878", UserPort.Start(_config, Free(9878)));
        Assert.Equal([9876, 9877, 9878], _tried);
        Assert.Equal(new AppConfig(CalendarApp.OutlookClassic, FirstRunDone: true, Port: 9878), _config.Load());
    }

    [Fact]
    public void A_saved_port_is_the_only_one_tried()
    {
        _config.Save(new AppConfig(Port: 9880));
        Assert.Null(UserPort.Start(_config, Free(9876)));
        Assert.Equal([9880], _tried);
        Assert.Equal(9880, _config.Load().Port);

        Assert.Equal("server 9880", UserPort.Start(_config, Free(9880)));
    }

    [Fact]
    public void Nothing_is_saved_when_every_port_is_busy()
    {
        Assert.Null(UserPort.Start(_config, Free()));
        Assert.Equal(Enumerable.Range(9876, 24), _tried);
        Assert.Null(_config.Load().Port);
        Assert.Equal(IcsServer.DefaultPort, _config.Load().ServerPort);
    }
}
