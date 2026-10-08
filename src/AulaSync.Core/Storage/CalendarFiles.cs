using System.Text;

namespace AulaSync.Core;

// Kalenderfilernes navne på disken, fx 123456-AE-medarbejder-1001.ics: institutionens nummer, initialer (medarbejder)
// eller navn (klasse, lokale) og til sidst skemaets nøgle, så filerne er lette at kende fra hinanden i Stifinder og
// Finder. Adressen, kalenderprogrammet henter, er stadig nøglen (ScheduleRef.FileName, fx medarbejder-1001.ics), og
// filen findes ud fra nøglen sidst i navnet. Filer fra 3.0.0 hedder bare nøglen; de får det nye navn ved næste opdatering.
public static class CalendarFiles
{
    const int MaxPart = 40;

    public static string Name(string institution, ScheduleRef schedule)
    {
        var label = schedule.Kind == ScheduleKind.Employee && schedule.Initials != "" ? schedule.Initials : schedule.Name;
        return string.Join("-", new[] { Clean(institution), Clean(label), schedule.Key }.Where(p => p != "")) + ".ics";
    }

    // Skemaets filer, nyeste først. Normalt én; lige når en fil får nyt navn, kan der kort være to.
    public static IReadOnlyList<string> Find(string dir, string key)
    {
        var exact = key + ".ics";
        var suffix = "-" + exact;
        try
        {
            if (!Directory.Exists(dir)) return [];
            return Directory.EnumerateFiles(dir, "*.ics")
                .Where(path => Path.GetFileName(path) is var name && (name == exact || name.EndsWith(suffix, StringComparison.Ordinal)))
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .ToList();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return []; }
    }

    // Samme fil? På Windows og Mac er "ae-…" og "AE-…" samme fil, og et filsystem kan give æ, ø og å tilbage skilt ad (NFD).
    public static bool SameName(string path, string? other) =>
        other is not null && string.Equals(Path.GetFileName(path).Normalize(NormalizationForm.FormC),
            Path.GetFileName(other).Normalize(NormalizationForm.FormC), StringComparison.OrdinalIgnoreCase);

    // Bogstaver (også æ, ø og å) og cifre; alt andet bliver til én bindestreg. Højst 40 tegn.
    static string Clean(string text)
    {
        var sb = new StringBuilder();
        foreach (var c in text.Normalize(NormalizationForm.FormC))
        {
            if (char.IsLetterOrDigit(c)) sb.Append(c);
            else if (sb.Length > 0 && sb[^1] != '-') sb.Append('-');
        }
        var clean = sb.ToString();
        return (clean.Length > MaxPart ? clean[..MaxPart] : clean).TrimEnd('-');
    }
}
