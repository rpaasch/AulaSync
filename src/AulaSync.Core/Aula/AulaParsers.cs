using System.Globalization;
using System.Text.Json;

namespace AulaSync.Core;

public sealed record EventParseResult(IReadOnlyList<AulaEvent> Events, int Skipped);

public static class AulaParsers
{
    public static void EnsureOk(string json) => Data(json);

    public static Profile ParseProfileContext(string json)
    {
        var data = Data(json);
        if (!TryObject(data, "institutionProfile", out var ip))
            throw new AulaException("Profil mangler i svaret fra Aula");
        TryObject(ip, "institution", out var inst);
        var name = $"{Str(ip, "firstName")} {Str(ip, "lastName")}".Trim();
        // Rollen er institutionsrollen fra egen institution (data.institutions[]); institutionProfile.role er blot "employee".
        var id = Str(ip, "id");
        var role = Items(data, "institutions").Where(i => Str(i, "institutionProfileId") == id).Select(i => Str(i, "institutionRole")).FirstOrDefault() ?? "";
        return new Profile(id, name == "" ? "Ukendt" : name, Clean(Str(ip, "shortName")), role,
            Str(inst, "institutionCode"), Clean(Str(inst, "institutionName")));
    }

    public static Profile? ParseFirstProfile(string json)
    {
        var data = Data(json);
        foreach (var p in Items(data, "profiles"))
            foreach (var ip in Items(p, "institutionProfiles"))
            {
                var name = Clean(Str(p, "displayName"));
                return new Profile(Str(ip, "id"), name == "" ? "Ukendt" : name, "", Str(ip, "institutionRole"),
                    Str(ip, "institutionCode"), Clean(Str(ip, "institutionName")));
            }
        return null;
    }

    // Antal resultater i et søgesvar, også grupper og andre roller (de tæller med i Aulas loft).
    public static int CountResults(string json) => Items(Data(json), "results").Count();

    public static IReadOnlyList<Employee> ParseEmployeeSearch(string json) =>
        Items(Data(json), "results")
            .Where(r => Str(r, "portalRole") == "employee")
            .Select(r => new Employee(Str(r, "id"), Clean(Str(r, "name")), Clean(Str(r, "shortName")), Str(r, "institutionRole")))
            .Where(e => ScheduleRef.IsValidId(e.Id))
            .ToList();

    public static IReadOnlyList<Employee> ParseProfileMasterData(string json) =>
        Items(Data(json), "institutionProfiles")
            .Select(r => new Employee(Str(r, "id"), Clean(Str(r, "fullName")), Clean(Str(r, "shortName")), Str(r, "institutionRole")))
            .Where(e => ScheduleRef.IsValidId(e.Id))
            .ToList();

    // Hovedgrupper (klasser). Feltet hedder groupType i findProfilesAndGroups og type i findGroups.
    public static IReadOnlyList<NamedItem> ParseGroupSearch(string json) =>
        Items(Data(json), "results")
            .Where(r => string.Equals(Str(r, "groupType") is { Length: > 0 } g ? g : Str(r, "type"), "Hovedgruppe", StringComparison.OrdinalIgnoreCase))
            .Select(r => new NamedItem(Str(r, "id"), Clean(Str(r, "name"))))
            .Where(g => ScheduleRef.IsValidId(g.Id) && g.Name != "")
            .ToList();

    public static IReadOnlyList<NamedItem> ParseResources(string json)
    {
        var data = Data(json);
        if (data.ValueKind != JsonValueKind.Array) return [];
        // Kun lokaler: resourceType "location" eller "extra-location" som i webappens lokalevælger (ikke udstyr o.l.), og ikke
        // inaktive. Uden kategori beholdes ressourcen.
        return data.EnumerateArray()
            .Where(r => !TryObject(r, "resourceCategory", out var category)
                || Str(category, "resourceType").ToLowerInvariant() is "" or "location" or "extra-location")
            .Where(r => Str(r, "inactiveReason") == "")
            .Select(r => new NamedItem(Str(r, "id"), Clean(Str(r, "name"))))
            .Where(r => ScheduleRef.IsValidId(r.Id) && r.Name != "")
            .ToList();
    }

    public static EventParseResult ParseEvents(string json)
    {
        var data = Data(json);
        if (data.ValueKind != JsonValueKind.Array)
            throw new AulaException("Kalendersvaret fra Aula mangler data");

        var events = new List<AulaEvent>();
        int skipped = 0;
        foreach (var ev in data.EnumerateArray())
        {
            var id = Str(ev, "id");
            if (id == "" || !TryTime(Str(ev, "startDateTime"), out var start) || !TryTime(Str(ev, "endDateTime"), out var end))
            {
                skipped++;
                continue;
            }

            var teachers = new List<Participant>();
            Participant? substituteFor = null;
            var isSubstitute = false;
            if (TryObject(ev, "lesson", out var lesson))
            {
                isSubstitute = Str(lesson, "lessonStatus") == "substitute";
                foreach (var p in Items(lesson, "participants"))
                {
                    var name = Clean(Str(p, "teacherName"));
                    if (name == "") continue;
                    var person = new Participant(name, Clean(Str(p, "teacherInitials")));
                    if (isSubstitute && Str(p, "participantRole") == "primaryTeacher") substituteFor = person;
                    else teachers.Add(person);
                }
            }

            var groups = Items(ev, "invitedGroups").Select(g => Clean(Str(g, "name"))).Where(n => n != "").ToList();
            var location = Clean(Str(ev, "primaryResourceText"));
            if (location == "" && TryObject(ev, "primaryResource", out var resource)) location = Clean(Str(resource, "name"));

            events.Add(new AulaEvent(id, Clean(Str(ev, "title")), start, end, location, groups, teachers, isSubstitute, substituteFor));
        }
        return new EventParseResult(events, skipped);
    }

    // Aulas egne fejlkoder i status (fra webappens kode, se docs/superpowers/research/2026-10-06-aula-api-kortlaegning.md §3.1):
    // code 20 eller subCode 23 = sessionen er udløbet, code 448 = login påkrævet, 451/452 = kontoen har ikke (længere) adgang,
    // subCode 33 = profilen er deaktiveret, code 401/403 = ingen adgang/tilladelse.
    // Returnerer null, hvis svaret ikke har en af de koder (eller ikke er JSON).
    public static AulaException? KnownError(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            return TryObject(doc.RootElement, "status", out var status) ? KnownError(status) : null;
        }
        catch (JsonException) { return null; }
    }

    static AulaException? KnownError(JsonElement status)
    {
        var code = Int(status, "code");
        var subCode = Int(status, "subCode");
        if (code is 20 or 448 || subCode == 23) return new SessionExpiredException();
        if (code == 451) return new AccessDeniedException("Du har ikke fået adgang til Aula endnu");
        if (code == 452) return new AccessDeniedException("Din Aula-bruger er deaktiveret");
        if (subCode == 33) return new ForbiddenException("Profilen er deaktiveret i Aula");
        if (code is 401 or 403) return new ForbiddenException();
        return null;
    }

    static int? Int(JsonElement el, string name) =>
        el.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out var n) ? n : null;

    // Returnerer data-elementet. Kaster SessionExpiredException/ForbiddenException ved Aulas egne fejlkoder,
    // ellers AulaException ved alt andet end OK.
    static JsonElement Data(string json)
    {
        JsonDocument doc;
        try { doc = JsonDocument.Parse(json); }
        catch (JsonException ex) { throw new AulaException("Aula sendte et ugyldigt svar", ex); }
        using (doc)
        {
            var root = doc.RootElement;
            if (!TryObject(root, "status", out var status))
                throw new AulaException("Svaret fra Aula mangler status");
            if (KnownError(status) is { } error) throw error;
            var message = Str(status, "message");
            if (message != "OK")
                throw new AulaException($"Aula svarede: {(message == "" ? "ukendt fejl" : message)}");
            return root.TryGetProperty("data", out var data) ? data.Clone() : default;
        }
    }

    static bool TryObject(JsonElement el, string name, out JsonElement value)
    {
        value = default;
        return el.ValueKind == JsonValueKind.Object && el.TryGetProperty(name, out value) && value.ValueKind == JsonValueKind.Object;
    }

    static IEnumerable<JsonElement> Items(JsonElement el, string name) =>
        el.ValueKind == JsonValueKind.Object && el.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Array
            ? v.EnumerateArray().ToList()
            : Enumerable.Empty<JsonElement>();

    static string Str(JsonElement el, string name)
    {
        if (el.ValueKind != JsonValueKind.Object || !el.TryGetProperty(name, out var v)) return "";
        return v.ValueKind switch
        {
            JsonValueKind.String => v.GetString() ?? "",
            JsonValueKind.Number => v.GetRawText(),
            _ => "",
        };
    }

    static bool TryTime(string s, out DateTimeOffset t) =>
        DateTimeOffset.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out t);

    static string Clean(string s) => HtmlText.Strip(s).Trim();
}
