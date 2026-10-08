namespace AulaSync.Core;

// Valgene under "Hent skemaer fra Aula" i Indstillinger. Hver opdatering henter hvert valgt skema fra Aula
// (ScheduleFetcher), så sjældnere opdatering giver færre forespørgsler.
public static class UpdateIntervals
{
    public const int DefaultMinutes = 240;

    public static IReadOnlyList<int> Minutes { get; } = [30, 60, 120, DefaultMinutes, 480];

    public static TimeSpan Of(int? minutes) =>
        TimeSpan.FromMinutes(minutes is { } m && Minutes.Contains(m) ? m : DefaultMinutes);

    public static string Label(int minutes) => minutes switch
    {
        30 => "Hver halve time",
        60 => "Hver time",
        DefaultMinutes => "Hver 4. time (standard)",
        _ => $"Hver {minutes / 60}. time",
    };

    // Hvor tit kalenderprogrammet bør hente filen fra AulaSync (REFRESH-INTERVAL): højst en time, så en ny opdatering
    // ikke venter længe på at blive vist. Den hentning sker på computeren og spørger ikke Aula.
    public static TimeSpan Refresh(TimeSpan interval) => interval < TimeSpan.FromHours(1) ? interval : TimeSpan.FromHours(1);
}
