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
    public string InstitutionCode { get; private set; } = "";
    public string MyProfileId { get; private set; } = "";

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
                        InstitutionCode = ip.TryGetProperty("institutionCode", out var ic) ? ic.GetString() ?? "" : "";
                        MyProfileId = ip.TryGetProperty("id", out var pid) ? pid.ToString() : "";
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

    // === Medarbejdere og kalender ===

    public class Employee
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public string Initials { get; set; } = "";
        public string Role { get; set; } = "";
        public string RoleDanish => Role switch
        {
            "teacher" => "Lærer",
            "preschool-teacher" => "Pædagog",
            "leader" => "Leder",
            "other" => "Anden",
            _ => Role,
        };
    }

    public async Task<List<Employee>> GetEmployeesAsync()
    {
        var employees = new Dictionary<string, Employee>();
        var alphabet = "abcdefghijklmnopqrstuvwxyzæøå";

        foreach (var letter in alphabet)
        {
            var resp = await _client.GetAsync(
                $"{_apiUrl}?method=search.findProfilesAndGroups" +
                $"&text={letter}&instCodes[]={InstitutionCode}&typeahead=true&limit=100" +
                $"&portalRole=employee");
            var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
            foreach (var p in doc.RootElement.GetProperty("data").GetProperty("results").EnumerateArray())
            {
                var id = p.GetProperty("id").GetString() ?? "";
                var portal = p.TryGetProperty("portalRole", out var pr) ? pr.GetString() : "";
                if (portal != "employee" || employees.ContainsKey(id)) continue;

                var name = p.TryGetProperty("name", out var n) ? n.GetString() ?? "" : "";
                var initials = p.TryGetProperty("shortName", out var sn) ? sn.GetString() ?? "" : "";
                var role = p.TryGetProperty("institutionRole", out var ir) ? ir.GetString() ?? "" : "";
                employees[id] = new Employee
                {
                    Id = id,
                    Name = name,
                    Initials = initials,
                    Role = role,
                };
            }
        }

        // Resolve rigtige instProfileIds via getProfileMasterData
        // (findProfilesAndGroups returnerer ikke altid instProfileId)
        var resolved = new Dictionary<string, Employee>();
        var idsToResolve = employees.Keys.ToList();

        // Batch i grupper af 20
        for (int i = 0; i < idsToResolve.Count; i += 20)
        {
            var batch = idsToResolve.Skip(i).Take(20);
            var query = string.Join("&", batch.Select(id => $"instProfileIds[]={id}"));
            try
            {
                var resp = await _client.GetAsync($"{_apiUrl}?method=profiles.getProfileMasterData&{query}&fromAdministration=false");
                var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
                if (doc.RootElement.TryGetProperty("data", out var data) &&
                    data.TryGetProperty("institutionProfiles", out var profiles))
                {
                    foreach (var ip in profiles.EnumerateArray())
                    {
                        var instId = ip.TryGetProperty("id", out var iid) ? iid.ToString() : "";
                        var fullName = ip.TryGetProperty("fullName", out var fn) ? fn.GetString() ?? "" : "";
                        var shortName = ip.TryGetProperty("shortName", out var sn2) ? sn2.GetString() ?? "" : "";
                        var instRole = ip.TryGetProperty("institutionRole", out var ir2) ? ir2.GetString() ?? "" : "";

                        if (instId != "" && !resolved.ContainsKey(instId))
                        {
                            resolved[instId] = new Employee
                            {
                                Id = instId,
                                Name = fullName,
                                Initials = shortName,
                                Role = instRole,
                            };
                        }
                    }
                }
            }
            catch { }
        }

        // Brug resolved hvis tilgængelig, ellers original
        var final = resolved.Count > 0 ? resolved : employees;
        Log($"Fandt {final.Count} medarbejdere (resolved: {resolved.Count})");

        return final.Values
            .Where(e => e.Role is "teacher" or "preschool-teacher" or "leader" or "other")
            .OrderBy(e => e.Name)
            .ToList();
    }

    public Task<List<CalendarEvent>> GetCalendarAsync(string profileId, DateTime start, DateTime end)
        => GetCalendarAsync(new[] { int.Parse(profileId) }, start, end);

    public async Task<List<CalendarEvent>> GetCalendarAsync(IEnumerable<int> profileIds, DateTime start, DateTime end)
    {
        var startStr = new DateTimeOffset(start, TimeZoneInfo.Local.GetUtcOffset(start)).ToString("yyyy-MM-dd 00:00:00.0000zzz");
        var endStr = new DateTimeOffset(end, TimeZoneInfo.Local.GetUtcOffset(end)).ToString("yyyy-MM-dd 23:59:59.9990zzz");
        var body = new { instProfileIds = profileIds.ToArray(), resourceIds = Array.Empty<int>(), start = startStr, end = endStr };
        var json = JsonSerializer.Serialize(body, new JsonSerializerOptions
        {
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        });
        Log($"Calendar POST: {profileIds.Count()} profiler, {startStr} - {endStr}");

        var resp = await _client.PostAsync(
            $"{_apiUrl}?method=calendar.getEventsByProfileIdsAndResourceIds",
            new StringContent(json, System.Text.Encoding.UTF8, "application/json"));
        return ParseCalendarEvents(await resp.Content.ReadAsStringAsync());
    }

    public static string EventsToIcs(List<CalendarEvent> events, string name, string initials)
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("BEGIN:VCALENDAR");
        sb.AppendLine("VERSION:2.0");
        sb.AppendLine("PRODID:-//AulaSync//Calendar//DA");
        sb.AppendLine($"X-WR-CALNAME:{name} ({initials})");
        sb.AppendLine("CALSCALE:GREGORIAN");
        sb.AppendLine("METHOD:PUBLISH");

        foreach (var ev in events)
        {
            if (!DateTime.TryParse(ev.Start, out var dtStart)) continue;
            if (!DateTime.TryParse(ev.End, out var dtEnd)) continue;

            sb.AppendLine("BEGIN:VEVENT");
            sb.AppendLine($"UID:aula-{ev.Id}");
            sb.AppendLine($"DTSTART:{dtStart.ToUniversalTime():yyyyMMddTHHmmssZ}");
            sb.AppendLine($"DTEND:{dtEnd.ToUniversalTime():yyyyMMddTHHmmssZ}");
            // Titel: "FAG (Klasse)" eller bare "FAG"
            var summary = !string.IsNullOrEmpty(ev.Groups) ? $"{ev.Title} ({ev.Groups})" : ev.Title;
            sb.AppendLine($"SUMMARY:{summary}");
            if (!string.IsNullOrEmpty(ev.Location))
                sb.AppendLine($"LOCATION:{ev.Location}");
            var desc = new List<string>();
            if (!string.IsNullOrEmpty(ev.Teacher)) desc.Add(ev.Teacher);
            if (!string.IsNullOrEmpty(ev.Groups)) desc.Add($"Klasse: {ev.Groups}");
            if (!string.IsNullOrEmpty(ev.Location)) desc.Add($"Lokale: {ev.Location}");
            if (desc.Count > 0) sb.AppendLine($"DESCRIPTION:{string.Join(" | ", desc)}");
            sb.AppendLine("END:VEVENT");
        }

        sb.AppendLine("END:VCALENDAR");
        return sb.ToString();
    }

    public class CalendarEvent
    {
        public string Id { get; set; } = "";
        public string Title { get; set; } = "";
        public string Start { get; set; } = "";
        public string End { get; set; } = "";
        public string Teacher { get; set; } = "";
        public string Groups { get; set; } = "";
        public string Location { get; set; } = "";
        public List<string> BelongsToProfiles { get; set; } = new();
    }

    public class GroupInfo
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class ResourceInfo
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
    }

    public async Task<List<GroupInfo>> GetGroupsAsync()
    {
        var groups = new Dictionary<string, GroupInfo>();
        var chars = "abcdefghijklmnopqrstuvwxyz0123456789æøå";
        foreach (var c in chars)
        {
            try
            {
                var resp = await _client.GetAsync(
                    $"{_apiUrl}?method=search.findProfilesAndGroups&text={c}&instCodes[]={InstitutionCode}&typeahead=true&limit=100");
                var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
                var data = doc.RootElement.GetProperty("data");
                if (data.ValueKind == JsonValueKind.Null) continue;
                foreach (var p in data.GetProperty("results").EnumerateArray())
                {
                    var gt = p.TryGetProperty("groupType", out var gtp) ? gtp.GetString() ?? "" : "";
                    if (gt == "") continue;
                    var id = p.GetProperty("id").GetString() ?? "";
                    if (groups.ContainsKey(id)) continue;
                    groups[id] = new GroupInfo
                    {
                        Id = id,
                        Name = p.TryGetProperty("name", out var n) ? n.GetString() ?? "" : "",
                        Type = gt,
                    };
                }
            }
            catch { }
        }
        return groups.Values.OrderBy(g => g.Name).ToList();
    }

    public async Task<List<ResourceInfo>> GetResourcesAsync()
    {
        var resources = new Dictionary<string, ResourceInfo>();
        var chars = "abcdefghijklmnopqrstuvwxyz";
        foreach (var c in chars)
        {
            try
            {
                var resp = await _client.GetAsync(
                    $"{_apiUrl}?method=resources.listResources&query={c}&institutionCodes[]={InstitutionCode}");
                var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
                foreach (var r in doc.RootElement.GetProperty("data").EnumerateArray())
                {
                    var id = r.GetProperty("id").ToString();
                    if (resources.ContainsKey(id)) continue;
                    resources[id] = new ResourceInfo
                    {
                        Id = id,
                        Name = r.TryGetProperty("name", out var n) ? n.GetString() ?? "" : "",
                    };
                }
            }
            catch { }
        }
        return resources.Values.OrderBy(r => r.Name).ToList();
    }

    public async Task<List<CalendarEvent>> GetGroupCalendarAsync(string groupId, DateTime start, DateTime end)
    {
        var startStr = new DateTimeOffset(start, TimeZoneInfo.Local.GetUtcOffset(start))
            .ToString("yyyy-MM-ddT00:00:00.0000zzz");
        var endStr = new DateTimeOffset(end, TimeZoneInfo.Local.GetUtcOffset(end))
            .ToString("yyyy-MM-ddT23:59:59.9990zzz");

        // URL-encode + som %2B (GET request)
        startStr = startStr.Replace("+", "%2B");
        endStr = endStr.Replace("+", "%2B");

        var resp = await _client.GetAsync(
            $"{_apiUrl}?method=calendar.geteventsbygroupid&groupId={groupId}&start={startStr}&end={endStr}&includeOnlyInvitedEvents=false");
        return ParseCalendarEvents(await resp.Content.ReadAsStringAsync());
    }

    public async Task<List<CalendarEvent>> GetResourceCalendarAsync(string resourceId, DateTime start, DateTime end)
    {
        var startStr = new DateTimeOffset(start, TimeZoneInfo.Local.GetUtcOffset(start))
            .ToString("yyyy-MM-dd 00:00:00.0000zzz");
        var endStr = new DateTimeOffset(end, TimeZoneInfo.Local.GetUtcOffset(end))
            .ToString("yyyy-MM-dd 23:59:59.9990zzz");

        var json = JsonSerializer.Serialize(new
        {
            instProfileIds = Array.Empty<int>(),
            resourceIds = new[] { int.Parse(resourceId) },
            start = startStr,
            end = endStr,
        }, new JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping });

        var resp = await _client.PostAsync(
            $"{_apiUrl}?method=calendar.getEventsByProfileIdsAndResourceIds",
            new StringContent(json, System.Text.Encoding.UTF8, "application/json"));
        return ParseCalendarEvents(await resp.Content.ReadAsStringAsync());
    }

    private List<CalendarEvent> ParseCalendarEvents(string responseBody)
    {
        var doc = JsonDocument.Parse(responseBody);
        if (!doc.RootElement.TryGetProperty("data", out var data) || data.ValueKind != JsonValueKind.Array)
            return new();

        var events = new List<CalendarEvent>();
        foreach (var ev in data.EnumerateArray())
        {
            try
            {
                var teacher = "";
                if (ev.TryGetProperty("lesson", out var les) && les.TryGetProperty("participants", out var parts))
                {
                    var names = new List<string>();
                    foreach (var p in parts.EnumerateArray())
                    {
                        var tn = p.TryGetProperty("teacherName", out var tn2) ? tn2.GetString() ?? "" : "";
                        var ti = p.TryGetProperty("teacherInitials", out var ti2) ? ti2.GetString() ?? "" : "";
                        if (tn != "") names.Add($"{tn} ({ti})");
                    }
                    teacher = string.Join(", ", names);
                }

                var groups = "";
                if (ev.TryGetProperty("invitedGroups", out var grps))
                {
                    var gnames = new List<string>();
                    foreach (var g in grps.EnumerateArray())
                        if (g.TryGetProperty("name", out var gn)) gnames.Add(gn.GetString() ?? "");
                    groups = string.Join(", ", gnames);
                }

                var belongsTo = new List<string>();
                if (ev.TryGetProperty("belongsToProfiles", out var btp) && btp.ValueKind == JsonValueKind.Array)
                    foreach (var p in btp.EnumerateArray())
                        belongsTo.Add(p.ToString());

                var location = ev.TryGetProperty("primaryResourceText", out var prt) ? prt.GetString() ?? "" : "";

                events.Add(new CalendarEvent
                {
                    Id = ev.TryGetProperty("id", out var id) ? id.ToString() : "",
                    Title = ev.TryGetProperty("title", out var t) ? t.GetString() ?? "" : "",
                    Start = ev.TryGetProperty("startDateTime", out var s) ? s.GetString() ?? "" : "",
                    End = ev.TryGetProperty("endDateTime", out var e2) ? e2.GetString() ?? "" : "",
                    Teacher = teacher,
                    Groups = groups,
                    Location = location,
                    BelongsToProfiles = belongsTo,
                });
            }
            catch { }
        }
        return events;
    }

    // === Beskeder ===

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
