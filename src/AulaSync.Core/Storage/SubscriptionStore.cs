using System.Text.Json;

namespace AulaSync.Core;

// AddedAt: brugeren har klikket hovedknappen (abonnér/kopiér). ImportedAt/ImportHash: brugeren har importeret filen.
// FetchedAt: et kalenderprogram hentede filen fra AulaSync (første hentning efter AddedAt).
// ImportUntil: slutningen af det tidsrum, importen dækker (90 dage efter ImportedAt); ImportHash dækker lektionerne fra
// ImportedAt til ImportUntil (null: hele filen, fra før tidsrummet blev gemt).
public sealed record Subscription(ScheduleRef Schedule, DateTimeOffset? AddedAt = null, DateTimeOffset? ImportedAt = null,
    string? ImportHash = null, DateTimeOffset? FetchedAt = null, DateTimeOffset? ImportUntil = null)
{
    public string Key => Schedule.Key;

    // Kalenderprogrammet har hentet filen efter brugerens seneste klik (eller helt uden klik, fx en adresse sat ind før).
    public bool Fetched => FetchedAt is { } fetched && (AddedAt is not { } added || fetched >= added);
}

public sealed class SubscriptionStore(string path)
{
    sealed record Entry(string? Kind, string? Id, string? Name, string? Initials,
        DateTimeOffset? AddedAt, DateTimeOffset? ImportedAt, string? ImportHash, string? Role = null, DateTimeOffset? FetchedAt = null,
        DateTimeOffset? ImportUntil = null);

    static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    public IReadOnlyList<Subscription> Load()
    {
        if (!File.Exists(path)) return [];
        var json = File.ReadAllText(path);
        List<Entry?>? entries;
        try { entries = JsonSerializer.Deserialize<List<Entry?>>(json); }
        catch (JsonException)
        {
            File.Copy(path, path + ".beskadiget", overwrite: true);
            return [];
        }

        var result = new List<Subscription>();
        foreach (var e in entries ?? [])
        {
            if (e?.Kind is null || e.Id is null || !ScheduleRef.IsValidId(e.Id)) continue;
            if (!Enum.TryParse<ScheduleKind>(e.Kind, out var kind) || !Enum.IsDefined(kind) || int.TryParse(e.Kind, out _)) continue;
            result.Add(new Subscription(new ScheduleRef(kind, e.Id, e.Name ?? "", e.Initials ?? "", e.Role ?? ""), e.AddedAt, e.ImportedAt, e.ImportHash, e.FetchedAt, e.ImportUntil));
        }
        return result;
    }

    public void Save(IEnumerable<Subscription> items) =>
        AtomicFile.WriteAllText(path, JsonSerializer.Serialize(
            items.Select(s => new Entry(s.Schedule.Kind.ToString(), s.Schedule.Id, s.Schedule.Name, s.Schedule.Initials,
                s.AddedAt, s.ImportedAt, s.ImportHash, s.Schedule.Role, s.FetchedAt, s.ImportUntil)).ToList(), Options));
}
