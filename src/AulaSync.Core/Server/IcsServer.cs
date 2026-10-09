using System.Globalization;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;

namespace AulaSync.Core;

public sealed record ServeDecision(int StatusCode, string? FilePath, DateTimeOffset? LastModified);

// En lille HTTP-server til .ics-filerne. Den lytter på 127.0.0.1 og ::1, fordi kalenderprogrammet kan slå localhost op
// som begge (HttpListener lytter kun på den første adresse). Den svarer kun, når værtsnavnet er localhost, 127.0.0.1
// eller [::1] med den rigtige port, så en hjemmeside, hvis navn peger på 127.0.0.1, ikke kan læse skemaerne.
public sealed partial class IcsServer : IDisposable
{
    public const int DefaultPort = 9876;
    const int MaxHeadBytes = 8192;
    const int MaxLoggedRequests = 200;
    const byte TlsHandshake = 0x16; // første byte, når et program taler https
    static readonly byte[] TlsHandshakeFailure = [0x15, 0x03, 0x01, 0x00, 0x02, 0x02, 0x28]; // TLS-alarm: fatal, handshake_failure
    static readonly TimeSpan ReadTimeout = TimeSpan.FromSeconds(10);

    [GeneratedRegex(@"^(medarbejder|klasse|lokale)-[A-Za-z0-9-]{1,40}\.ics$", RegexOptions.CultureInvariant)]
    private static partial Regex CalendarFileName();

    [GeneratedRegex(@"^[A-Z]{1,16}$", RegexOptions.CultureInvariant)]
    private static partial Regex Method();

    readonly string _calendarDir;
    readonly FileLog _log;
    readonly int _port;
    readonly List<TcpListener> _listeners = [];
    readonly CancellationTokenSource _stop = new();
    readonly HashSet<string> _logged = [];

    public IcsServer(string calendarDir, FileLog log, int port = DefaultPort)
    {
        _calendarDir = calendarDir;
        _log = log;
        _port = port;
    }

    // Et program har hentet en kalenderfil (svaret var 200 eller 304): navnet i adressen, fx "klasse-7.ics", og programmets
    // User-Agent ("" uden).
    public event Action<string, string>? Fetched;

    public static string UrlFor(ScheduleRef schedule, int port) => $"http://localhost:{port}/{schedule.FileName}";

    public static string WebcalFor(ScheduleRef schedule, int port) => $"webcal://localhost:{port}/{schedule.FileName}";

    // Adressen er skemaets nøgle (medarbejder-1001.ics); filen på disken kan have mere foran (CalendarFiles).
    public static ServeDecision Resolve(string calendarDir, string absolutePath, string method, DateTimeOffset? ifModifiedSince)
    {
        if (method is not ("GET" or "HEAD")) return new(405, null, null);
        var name = NameOf(absolutePath);
        if (!CalendarFileName().IsMatch(name)) return new(404, null, null);
        if (CalendarFiles.Find(calendarDir, name[..^".ics".Length]).FirstOrDefault() is not { } path) return new(404, null, null);
        var written = new DateTimeOffset(File.GetLastWriteTimeUtc(path), TimeSpan.Zero);
        var lastModified = written.AddTicks(-(written.Ticks % TimeSpan.TicksPerSecond)); // HTTP-datoer har hele sekunder
        if (ifModifiedSince is { } since && lastModified <= since) return new(304, path, lastModified);
        return new(200, path, lastModified);
    }

    static string NameOf(string absolutePath) => Uri.UnescapeDataString(absolutePath.TrimStart('/'));

    // 127.0.0.1 skal lykkes (ellers er porten optaget). ::1 skal også lykkes, medmindre computeren ikke har IPv6: optager
    // et andet program [::1]:port, ville kalenderprogrammer, der prøver ::1 først, ende dér, mens AulaSync sagde, at alt kørte.
    public bool TryStart()
    {
        if (!TryListen(IPAddress.Loopback, out var error))
        {
            _log.Error("Kalender-serveren kunne ikke starte", error);
            return false;
        }
        if (!TryListen(IPAddress.IPv6Loopback, out error))
        {
            if (!Ipv6Missing(error!.SocketErrorCode))
            {
                _log.Error($"Kalender-serveren kunne ikke starte: port {_port} er optaget på [::1]", error);
                foreach (var listener in _listeners) listener.Stop();
                _listeners.Clear();
                return false;
            }
            _log.Info($"Kalender-serveren lytter ikke på [::1]:{_port}: {error.Message}");
        }
        _log.Info($"Kalender-server lytter på {string.Join(", ", _listeners.Select(l => $"http://{l.LocalEndpoint}/"))}");
        foreach (var listener in _listeners) _ = Task.Run(() => AcceptLoopAsync(listener));
        return true;
    }

    // Fejlkoder, der betyder, at computeren ikke har IPv6 (ikke at porten er optaget).
    public static bool Ipv6Missing(SocketError error) =>
        error is SocketError.AddressFamilyNotSupported or SocketError.AddressNotAvailable or SocketError.ProtocolNotSupported;

    bool TryListen(IPAddress address, out SocketException? error)
    {
        TcpListener? listener = null;
        try
        {
            listener = new TcpListener(address, _port);
            if (OperatingSystem.IsWindows()) listener.ExclusiveAddressUse = true; // intet andet program kan dele porten
            listener.Start();
            _listeners.Add(listener);
            error = null;
            return true;
        }
        catch (SocketException ex)
        {
            listener?.Dispose();
            error = ex;
            return false;
        }
    }

    async Task AcceptLoopAsync(TcpListener listener)
    {
        while (!_stop.IsCancellationRequested)
        {
            TcpClient client;
            try { client = await listener.AcceptTcpClientAsync(_stop.Token); }
            catch (Exception ex) when (_stop.IsCancellationRequested || ex is ObjectDisposedException) { return; }
            catch (SocketException ex)
            {
                _log.Error("Kalender-serveren kunne ikke tage imod en forbindelse", ex);
                await Task.Delay(100);
                continue;
            }
            _ = Task.Run(() => HandleAsync(client));
        }
    }

    async Task HandleAsync(TcpClient client)
    {
        using (client)
        using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(_stop.Token))
        {
            try
            {
                timeout.CancelAfter(ReadTimeout);
                var stream = client.GetStream();
                var head = await ReadHeadAsync(stream, timeout.Token);
                if (head is null) { await ReplyAsync(client, stream, 431, null, null, timeout.Token); return; }
                if (head.Length > 0 && head[0] == TlsHandshake)
                {
                    // Højst én linje i minuttet, så loggen viser rækkefølgen ("https, https, GET 200") ved et nyt abonnement.
                    LogOnce($"https {DateTimeOffset.UtcNow:yyyyMMddHHmm}", "Kalender-server: et program forsøgte at hente via https; AulaSync svarer kun på http");
                    // En almindelig TLS-afvisning i stedet for en forbindelse, der bare forsvinder, så programmet prøver http.
                    await stream.WriteAsync(TlsHandshakeFailure, timeout.Token);
                    await CloseGentlyAsync(client, stream, timeout.Token);
                    return;
                }
                var request = Parse(head);
                if (request is null)
                {
                    var line = Printable(head.Split("\r\n")[0]);
                    LogOnce($"ugyldig {line}", $"Kalender-server: ugyldig forespørgsel \"{line}\"");
                    await ReplyAsync(client, stream, 400, null, null, timeout.Token);
                    return;
                }
                var (method, target, host, since, agent) = request.Value;
                var from = agent == "" ? "" : $" ({Printable(agent)})";
                if (!IsLoopbackHost(host))
                {
                    LogOnce($"host {host}", $"Kalender-server: afviste {method} {Printable(target)} med værten \"{Printable(host)}\"{from}");
                    await ReplyAsync(client, stream, 400, null, null, timeout.Token);
                    return;
                }
                var path = target.Split('?', '#')[0];
                var decision = path.StartsWith('/') ? Resolve(_calendarDir, path, method, since) : new(400, null, null);
                if (decision.StatusCode != 304)
                    LogOnce($"{method} {decision.StatusCode} {path}", $"Kalender-server: {method} {Printable(path)} → {decision.StatusCode}{from}");
                if (decision.StatusCode is 200 or 304) Fetched?.Invoke(NameOf(path), agent);
                await ReplyAsync(client, stream, decision.StatusCode, decision, method == "GET" ? decision.FilePath : null, timeout.Token);
            }
            catch (Exception ex) when (ex is IOException or SocketException or OperationCanceledException or ObjectDisposedException)
            {
                // Klienten lukkede, var for langsom, eller serveren stopper.
            }
            catch (Exception ex) { _log.Error("ICS-forespørgsel fejlede", ex); }
        }
    }

    // Læser til den tomme linje efter headerne; null, hvis den ikke kommer inden for MaxHeadBytes. Begynder forbindelsen
    // med et TLS-håndtryk (https), returneres det læste straks, for så kommer den tomme linje aldrig.
    static async Task<string?> ReadHeadAsync(NetworkStream stream, CancellationToken ct)
    {
        var buffer = new byte[MaxHeadBytes];
        var length = 0;
        while (length < buffer.Length)
        {
            var read = await stream.ReadAsync(buffer.AsMemory(length), ct);
            if (read == 0) throw new IOException("Forbindelsen blev lukket");
            var from = Math.Max(0, length - 3);
            length += read;
            if (buffer[0] == TlsHandshake) return Encoding.Latin1.GetString(buffer, 0, length);
            var end = buffer.AsSpan(from, length - from).IndexOf("\r\n\r\n"u8);
            if (end >= 0) return Encoding.Latin1.GetString(buffer, 0, from + end);
        }
        return null;
    }

    static (string Method, string Target, string Host, DateTimeOffset? Since, string Agent)? Parse(string head)
    {
        var lines = head.Split("\r\n");
        var parts = lines[0].Split(' ');
        if (parts.Length != 3 || !Method().IsMatch(parts[0]) || !parts[2].StartsWith("HTTP/1.", StringComparison.Ordinal)) return null;
        string? host = null;
        DateTimeOffset? since = null;
        var agent = "";
        foreach (var line in lines.Skip(1))
        {
            var colon = line.IndexOf(':');
            if (colon <= 0) return null;
            var name = line[..colon];
            var value = line[(colon + 1)..].Trim();
            if (name.Equals("Host", StringComparison.OrdinalIgnoreCase))
            {
                if (host is not null) return null; // to Host-headere er en ugyldig forespørgsel
                host = value;
            }
            else if (name.Equals("If-Modified-Since", StringComparison.OrdinalIgnoreCase)
                && DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var parsed))
                since = parsed;
            else if (name.Equals("User-Agent", StringComparison.OrdinalIgnoreCase))
                agent = value;
        }
        return host is null ? null : (parts[0], parts[1], host, since, agent);
    }

    bool IsLoopbackHost(string host) =>
        host.Equals($"localhost:{_port}", StringComparison.OrdinalIgnoreCase)
        || host == $"127.0.0.1:{_port}"
        || host == $"[::1]:{_port}";

    static async Task ReplyAsync(TcpClient client, NetworkStream stream, int status, ServeDecision? decision, string? bodyPath, CancellationToken ct)
    {
        await WriteAsync(stream, status, decision, bodyPath, ct);
        await CloseGentlyAsync(client, stream, ct);
    }

    // Luk pænt: send FIN, og læs det, klienten stadig sender (højst 1 s). Ellers giver ulæste data en TCP-reset, som på
    // Windows kan slette svaret hos klienten, før den har læst det.
    static async Task CloseGentlyAsync(TcpClient client, NetworkStream stream, CancellationToken ct)
    {
        client.Client.Shutdown(SocketShutdown.Send);
        using var linger = CancellationTokenSource.CreateLinkedTokenSource(ct);
        linger.CancelAfter(TimeSpan.FromSeconds(1));
        var buffer = new byte[8192];
        while (await stream.ReadAsync(buffer, linger.Token) > 0) { }
    }

    static async Task WriteAsync(NetworkStream stream, int status, ServeDecision? decision, string? bodyPath, CancellationToken ct)
    {
        byte[] body = [];
        if (bodyPath is not null && status == 200)
        {
            // Del-adgang, så sync kan erstatte filen, mens den sendes.
            await using var file = new FileStream(bodyPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            body = new byte[file.Length];
            await file.ReadExactlyAsync(body, ct);
        }
        var length = status == 200 && decision?.FilePath is { } path && bodyPath is null ? new FileInfo(path).Length : body.Length;
        var head = new StringBuilder();
        head.Append(CultureInfo.InvariantCulture, $"HTTP/1.1 {status} {Reason(status)}\r\n");
        head.Append(CultureInfo.InvariantCulture, $"Date: {DateTimeOffset.UtcNow:R}\r\n");
        if (status == 200) head.Append("Content-Type: text/calendar; charset=utf-8\r\n");
        if (decision?.LastModified is { } lm) head.Append(CultureInfo.InvariantCulture, $"Last-Modified: {lm:R}\r\n");
        if (status is 200 or 304) head.Append("Cache-Control: no-cache\r\n");
        if (status == 405) head.Append("Allow: GET, HEAD\r\n");
        if (status != 304) head.Append(CultureInfo.InvariantCulture, $"Content-Length: {length}\r\n");
        head.Append("Connection: close\r\n\r\n");
        await stream.WriteAsync(Encoding.ASCII.GetBytes(head.ToString()), ct);
        if (body.Length > 0) await stream.WriteAsync(body, ct);
    }

    static string Reason(int status) => status switch
    {
        200 => "OK",
        304 => "Not Modified",
        400 => "Bad Request",
        404 => "Not Found",
        405 => "Method Not Allowed",
        431 => "Request Header Fields Too Large",
        _ => "",
    };

    // Hver slags svar logges én gang pr. kørsel (kalenderprogrammer henter tit), og højst MaxLoggedRequests i alt.
    void LogOnce(string key, string message)
    {
        lock (_logged)
            if (_logged.Count >= MaxLoggedRequests || !_logged.Add(key)) return;
        _log.Info(message);
    }

    static string Printable(string s)
    {
        var clean = new string(s.Select(c => char.IsControl(c) ? '?' : c).ToArray());
        return clean.Length > 120 ? clean[..120] + "…" : clean;
    }

    public void Dispose()
    {
        _stop.Cancel();
        foreach (var listener in _listeners) listener.Stop();
        _listeners.Clear();
    }
}
