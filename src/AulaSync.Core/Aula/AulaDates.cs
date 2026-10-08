using System.Globalization;

namespace AulaSync.Core;

public static class AulaDates
{
    // Aulas format: "2026-04-13 00:00:00.0000+02:00" (POST) eller med 'T' (GET).
    public static string Format(DateOnly date, bool endOfDay, char separator, TimeZoneInfo tz)
    {
        var time = endOfDay ? new TimeOnly(23, 59, 59, 999) : TimeOnly.MinValue;
        var local = date.ToDateTime(time, DateTimeKind.Unspecified);
        var value = new DateTimeOffset(local, tz.GetUtcOffset(local));
        var pattern = separator == 'T' ? "yyyy-MM-dd'T'HH:mm:ss.ffffzzz" : "yyyy-MM-dd HH:mm:ss.ffffzzz";
        return value.ToString(pattern, CultureInfo.InvariantCulture);
    }
}
