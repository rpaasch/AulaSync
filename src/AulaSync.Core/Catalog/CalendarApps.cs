namespace AulaSync.Core;

public enum MainActionKind { Subscribe, Import, CopyAddress }

public sealed record MainAction(string Label, MainActionKind Kind);

public static class CalendarApps
{
    public static readonly CalendarApp[] All = [CalendarApp.AppleCalendar, CalendarApp.OutlookClassic, CalendarApp.OutlookImport, CalendarApp.Other];

    public static MainAction ActionFor(CalendarApp app) => app switch
    {
        CalendarApp.AppleCalendar => new("Tilføj til Kalender", MainActionKind.Subscribe),
        CalendarApp.OutlookClassic => new("Tilføj til Outlook", MainActionKind.Subscribe),
        CalendarApp.OutlookImport => new("Importér…", MainActionKind.Import),
        _ => new("Kopiér adresse", MainActionKind.CopyAddress),
    };

    public static string Title(CalendarApp app) => app switch
    {
        CalendarApp.AppleCalendar => "Apple Kalender",
        CalendarApp.OutlookClassic => "Outlook (klassisk)",
        CalendarApp.OutlookImport => "Ny Outlook, Outlook til Mac eller web",
        _ => "Andet program",
    };

    // Kun import giver et øjebliksbillede; de andre abonnerer på filen og opdateres af sig selv.
    public static bool UpdatesAutomatically(CalendarApp app) => ActionFor(app).Kind != MainActionKind.Import;

    public static string ModeLabel(CalendarApp app) => UpdatesAutomatically(app) ? "Opdateres automatisk" : "Øjebliksbillede";

    public static string Description(CalendarApp app) => app switch
    {
        CalendarApp.AppleCalendar => "Kalender-appen på denne Mac. Skemaerne opdateres af sig selv, så længe AulaSync kører.",
        CalendarApp.OutlookClassic => "Den klassiske Outlook til Windows. Skemaerne tilføjes som internetkalendere og opdateres af sig selv, så længe AulaSync kører.",
        CalendarApp.OutlookImport => "Henter kalendere gennem Microsofts servere, som ikke kan nå denne computer. Du importerer en fil, og AulaSync siger til, når skemaet er ændret, så du kan importere igen.",
        // Adressen er http://localhost:…, så den virker kun i et program på samme computer (ikke fx Google Kalender).
        _ => "Fx Thunderbird. Kopiér en .ics-adresse, og indsæt den i et kalenderprogram på denne computer. Opdateres, så længe AulaSync kører.",
    };

    public static CalendarApp? DefaultFor(bool isMac) => isMac ? CalendarApp.AppleCalendar : null;
}
