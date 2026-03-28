using System.Text.Json;

namespace AulaSync;

/// <summary>
/// Aula REST API client.
/// </summary>
class AulaApi
{
    private HttpClient _client;
    private string _apiUrl = "https://www.aula.dk/api/v23";

    public string UserName { get; private set; } = "";
    public string Role { get; private set; } = "";
    public string Institution { get; private set; } = "";

    private static readonly Dictionary<string, string> RoleMap = new()
    {
        ["employee"] = "Medarbejder",
        ["parent"] = "Forælder",
        ["guardian"] = "Forælder",
        ["child"] = "Elev",
        ["teacher"] = "Lærer",
    };

    public string RoleDanish => RoleMap.GetValueOrDefault(Role, Role);

    private static readonly string LogFile = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".aulasync", "aulasync.log");

    private static void Log(string msg)
    {
        try { File.AppendAllText(LogFile, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [API] {msg}\n"); } catch { }
    }

    public AulaApi(HttpClient client)
    {
        _client = client;
    }

    public async Task<bool> ConnectAsync()
    {
        // Find API-version
        for (int v = 23; v < 33; v++)
        {
            _apiUrl = $"https://www.aula.dk/api/v{v}";
            var resp = await _client.GetAsync($"{_apiUrl}?method=profiles.getProfilesByLogin");
            if (resp.StatusCode == System.Net.HttpStatusCode.Gone) continue;
            if (!resp.IsSuccessStatusCode) continue;

            var json = await resp.Content.ReadAsStringAsync();
            var doc = JsonDocument.Parse(json);
            var status = doc.RootElement.GetProperty("status").GetProperty("message").GetString();
            if (status != "OK") continue;

            // Første profil matcher login-valget (API sorterer efter aktiv rolle)
            var profiles = doc.RootElement.GetProperty("data").GetProperty("profiles");
            foreach (var profile in profiles.EnumerateArray())
            {
                UserName = profile.GetProperty("displayName").GetString() ?? "Ukendt";
                if (profile.TryGetProperty("institutionProfiles", out var ips))
                {
                    foreach (var ip in ips.EnumerateArray())
                    {
                        Role = ip.TryGetProperty("role", out var r) ? r.GetString() ?? "" : "";
                        Institution = ip.TryGetProperty("institutionName", out var i) ? i.GetString() ?? "" : "";
                        break;
                    }
                }
                break;
            }
            Log($"Profil: {UserName} | {RoleDanish} | {Institution}");
            return UserName != "";
        }
        return false;
    }

    public async Task<List<JsonElement>> GetRecentThreadsAsync(int pages = 2)
    {
        var all = new List<JsonElement>();
        for (int page = 0; page < pages; page++)
        {
            var resp = await _client.GetAsync(
                $"{_apiUrl}?method=messaging.getThreads&sortOn=date&orderDirection=desc&page={page}");
            resp.EnsureSuccessStatusCode();
            var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
            var threads = doc.RootElement.GetProperty("data").GetProperty("threads");
            if (threads.GetArrayLength() == 0) break;
            foreach (var t in threads.EnumerateArray())
                all.Add(t.Clone());
        }
        return all;
    }

    public async Task<AulaMessage?> GetFullMessageAsync(JsonElement thread)
    {
        var threadId = thread.GetProperty("id").GetInt64();

        JsonDocument doc;
        try
        {
            var resp = await _client.GetAsync(
                $"{_apiUrl}?method=messaging.getMessagesForThread&threadId={threadId}&page=0");
            if (resp.StatusCode == System.Net.HttpStatusCode.Forbidden)
            {
                return new AulaMessage
                {
                    ThreadId = threadId,
                    Subject = thread.TryGetProperty("subject", out var s) ? s.GetString() ?? "" : "",
                    Sender = "Ukendt",
                    Text = "Denne besked kræver MitID-login.",
                    IsSensitive = true,
                };
            }
            resp.EnsureSuccessStatusCode();
            doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
        }
        catch { return null; }

        var data = doc.RootElement.GetProperty("data");
        var messages = data.GetProperty("messages");
        if (messages.GetArrayLength() == 0) return null;

        // Find seneste besked
        JsonElement latest = messages[0];
        foreach (var msg in messages.EnumerateArray())
        {
            if (msg.TryGetProperty("messageType", out var mt) && mt.GetString() == "Message")
            {
                latest = msg;
                break;
            }
        }

        var text = "";
        if (latest.TryGetProperty("text", out var textEl))
        {
            if (textEl.ValueKind == JsonValueKind.Object && textEl.TryGetProperty("html", out var html))
                text = html.GetString() ?? "";
            else
                text = textEl.ToString();
        }

        var sender = "Ukendt";
        var senderMeta = "";
        if (latest.TryGetProperty("sender", out var senderEl))
        {
            sender = senderEl.TryGetProperty("fullName", out var fn) ? fn.GetString() ?? "Ukendt" : "Ukendt";
            senderMeta = senderEl.TryGetProperty("metadata", out var md) ? md.GetString() ?? "" : "";
        }

        var timestamp = latest.TryGetProperty("sendDateTime", out var ts) ? ts.GetString() ?? "" : "";
        var subject = data.TryGetProperty("subject", out var subj) ? subj.GetString() ?? "" : "";
        var isRead = thread.TryGetProperty("read", out var rd) && rd.GetBoolean();

        // Modtagere
        var recipients = data.TryGetProperty("recipients", out var rcp) ? rcp : default;
        int totalRecipients = recipients.ValueKind == JsonValueKind.Array ? recipients.GetArrayLength() : 0;
        var totalMessages = data.TryGetProperty("totalMessageCount", out var tmc) ? tmc.GetInt32() : messages.GetArrayLength();

        // Tidligere beskeder
        var allMessages = new List<AulaSubMessage>();
        foreach (var m in messages.EnumerateArray())
        {
            if (m.TryGetProperty("messageType", out var mmt) && mmt.GetString() == "Message")
            {
                var mText = "";
                if (m.TryGetProperty("text", out var mTextEl))
                    mText = mTextEl.ValueKind == JsonValueKind.Object && mTextEl.TryGetProperty("html", out var mh)
                        ? mh.GetString() ?? "" : mTextEl.ToString();

                allMessages.Add(new AulaSubMessage
                {
                    Sender = m.TryGetProperty("sender", out var ms) && ms.TryGetProperty("fullName", out var mfn)
                        ? mfn.GetString() ?? "" : "",
                    SenderMeta = m.TryGetProperty("sender", out var ms2) && ms2.TryGetProperty("metadata", out var mmd)
                        ? mmd.GetString() ?? "" : "",
                    Text = mText,
                    Timestamp = m.TryGetProperty("sendDateTime", out var mts) ? mts.GetString() ?? "" : "",
                });
            }
        }

        return new AulaMessage
        {
            ThreadId = threadId,
            Subject = subject,
            Sender = sender,
            SenderMeta = senderMeta,
            Text = text,
            Timestamp = timestamp,
            IsRead = isRead,
            TotalRecipients = totalRecipients,
            TotalMessages = totalMessages,
            AllMessages = allMessages,
        };
    }

    public async Task TestSessionAsync()
    {
        var resp = await _client.GetAsync($"{_apiUrl}?method=profiles.getProfilesByLogin");
        var json = await resp.Content.ReadAsStringAsync();
        var doc = JsonDocument.Parse(json);
        var code = doc.RootElement.GetProperty("status");
        if (code.TryGetProperty("code", out var c) && c.GetInt32() == 403)
            throw new UnauthorizedAccessException("Session udløbet");
    }
}

class AulaMessage
{
    public long ThreadId { get; set; }
    public string Subject { get; set; } = "";
    public string Sender { get; set; } = "";
    public string SenderMeta { get; set; } = "";
    public string Text { get; set; } = "";
    public string Timestamp { get; set; } = "";
    public bool IsRead { get; set; }
    public bool IsSensitive { get; set; }
    public int TotalRecipients { get; set; }
    public int TotalMessages { get; set; }
    public List<AulaSubMessage> AllMessages { get; set; } = new();
}

class AulaSubMessage
{
    public string Sender { get; set; } = "";
    public string SenderMeta { get; set; } = "";
    public string Text { get; set; } = "";
    public string Timestamp { get; set; } = "";
}
