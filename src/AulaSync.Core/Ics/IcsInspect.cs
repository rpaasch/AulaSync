using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace AulaSync.Core;

// Fetched: hvornår skemaet blev hentet fra Aula (X-AULASYNC-FETCHED; i filer fra før 3.4 lektionernes DTSTAMP), null i en
// fil uden begge.
public sealed record IcsSummary(int Events, string ContentHash, DateTimeOffset? Fetched = null);

public static partial class IcsInspect
{
    // Tidspunktet sidst i en lektions beskrivelse (IcsWriter med StatusStyle.InDescription), fx "…\n\nOpd. 091026@07:30",
    // også når AulaSync er logget ud ("…\n\n⚠ Log ind i AulaSync · Opd. 091026@07:30").
    [GeneratedRegex(@"^DESCRIPTION:(?<rest>.*?)(?:\\n\\n)?(?:" + EventFormatter.LoggedOutPrefix + @")?Opd\. \d{6}@\d{2}:\d{2}$",
        RegexOptions.CultureInvariant)]
    internal static partial Regex StatusInDescription();

    // Linjerne med fortsættelseslinjerne sat sammen (RFC 5545 §3.1).
    internal static string[] Unfold(string ics) => ics.Replace("\r\n ", "").Replace("\r\n\t", "").Split("\r\n");

    // Antal begivenheder og en hash af indholdet. DTSTAMP ændres ved hver skrivning og tæller derfor ikke med, og det gør
    // statusbegivenhederne og tidspunktet sidst i beskrivelsen heller ikke (IcsWriter.StatusUidPrefix, StatusInDescription). from/until: kun begivenheder, der starter i tidsrummet
    // (begge grænser med), tæller.
    public static IcsSummary Read(string ics, DateTimeOffset? from = null, DateTimeOffset? until = null)
    {
        var events = new List<(string Uid, string Body, DateTimeOffset? Start)>();
        List<string>? current = null;
        var uid = "";
        DateTimeOffset? start = null;
        DateTimeOffset? fetched = null, stamped = null;
        foreach (var line in Unfold(ics))
        {
            if (line.StartsWith(IcsWriter.FetchedProperty + ":", StringComparison.Ordinal)) fetched ??= Utc(line[(IcsWriter.FetchedProperty.Length + 1)..]);
            if (line.StartsWith("DTSTAMP:", StringComparison.Ordinal) && current is not null) stamped ??= Utc(line[8..]);
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
            if (StatusInDescription().Match(line) is { Success: true } m)
            {
                if (m.Groups["rest"].Value != "") current.Add($"DESCRIPTION:{m.Groups["rest"].Value}");
                continue;
            }
            current.Add(line);
        }
        if (from is not null || until is not null)
            events = events.Where(e => e.Start is { } s && (from is not { } f || s >= f) && (until is not { } u || s <= u)).ToList();
        var canonical = string.Join("\n\n", events.OrderBy(e => e.Uid, StringComparer.Ordinal).Select(e => e.Body));
        return new IcsSummary(events.Count, Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(canonical))),
            fetched ?? stamped);
    }

    // AulaSync skriver DTSTART i UTC (IcsWriter), fx 20260413T060000Z.
    internal static DateTimeOffset? Utc(string value) =>
        DateTimeOffset.TryParseExact(value, "yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var t) ? t : null;

    public static IcsSummary? ReadFile(string path, DateTimeOffset? from = null, DateTimeOffset? until = null)
    {
        try { return Read(File.ReadAllText(path), from, until); }
        catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException or IOException or UnauthorizedAccessException) { return null; }
    }
}
