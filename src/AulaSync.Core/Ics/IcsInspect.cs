using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace AulaSync.Core;

public sealed record IcsSummary(int Events, string ContentHash);

public static class IcsInspect
{
    // Antal begivenheder og en hash af indholdet. DTSTAMP ændres ved hver skrivning og tæller derfor ikke med, og det gør
    // statusbegivenheden heller ikke (IcsWriter.StatusUidPrefix). from/until: kun begivenheder, der starter i tidsrummet
    // (begge grænser med), tæller.
    public static IcsSummary Read(string ics, DateTimeOffset? from = null, DateTimeOffset? until = null)
    {
        var unfolded = ics.Replace("\r\n ", "").Replace("\r\n\t", "");
        var events = new List<(string Uid, string Body, DateTimeOffset? Start)>();
        List<string>? current = null;
        var uid = "";
        DateTimeOffset? start = null;
        foreach (var line in unfolded.Split("\r\n"))
        {
            if (line == "BEGIN:VEVENT") { current = []; uid = ""; start = null; continue; }
            if (line == "END:VEVENT" && current is not null)
            {
                if (!uid.StartsWith(IcsWriter.StatusUidPrefix, StringComparison.Ordinal))
                    events.Add((uid, string.Join("\n", current), start));
                current = null;
                continue;
            }
            if (current is null || line.StartsWith("DTSTAMP:", StringComparison.Ordinal)) continue;
            if (line.StartsWith("UID:", StringComparison.Ordinal)) uid = line[4..];
            if (line.StartsWith("DTSTART:", StringComparison.Ordinal)) start = Utc(line[8..]);
            current.Add(line);
        }
        if (from is not null || until is not null)
            events = events.Where(e => e.Start is { } s && (from is not { } f || s >= f) && (until is not { } u || s <= u)).ToList();
        var canonical = string.Join("\n\n", events.OrderBy(e => e.Uid, StringComparer.Ordinal).Select(e => e.Body));
        return new IcsSummary(events.Count, Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(canonical))));
    }

    // AulaSync skriver DTSTART i UTC (IcsWriter), fx 20260413T060000Z.
    static DateTimeOffset? Utc(string value) =>
        DateTimeOffset.TryParseExact(value, "yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var t) ? t : null;

    public static IcsSummary? ReadFile(string path, DateTimeOffset? from = null, DateTimeOffset? until = null)
    {
        try { return Read(File.ReadAllText(path), from, until); }
        catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException or IOException or UnauthorizedAccessException) { return null; }
    }
}
