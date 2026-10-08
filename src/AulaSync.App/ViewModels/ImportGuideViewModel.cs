using AulaSync.Core;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace AulaSync.App;

// Import-guiden til ny Outlook, Outlook til Mac og Outlook på nettet: tre trin, hver med sin egen knap. filePath slås op,
// når den bruges: filen kan få nyt navn, mens guiden er åben (CalendarFiles).
public sealed partial class ImportGuideViewModel(ScheduleRef schedule, Func<string> filePath, IPlatform platform, TimeProvider time) : ObservableObject
{
    public const string OutlookWebUrl = "https://outlook.office.com/calendar";

    public string Title => $"Importér {schedule.CalendarName}";
    public string CalendarName => schedule.CalendarName;
    public string FileName => Path.GetFileName(filePath());
    public string RevealLabel => $"Vis fil i {platform.FileManager}";
    public string Hint => "Importen er et øjebliksbillede. Når AulaSync viser Ændret siden import, skal du slette kalenderen og importere igen.";

    [ObservableProperty] public partial string CopyNameLabel { get; set; } = "Kopiér navn";

    [RelayCommand]
    async Task CopyName()
    {
        CopyNameLabel = CopyResult.Label(await platform.CopyTextAsync(schedule.CalendarName));
        await Task.Delay(CopyResult.Shown, time);
        CopyNameLabel = "Kopiér navn";
    }

    [RelayCommand] void OpenOutlook() => platform.Open(OutlookWebUrl);

    [RelayCommand] void Reveal() => platform.RevealFile(filePath());
}

// "Åbnede Outlook ikke kalenderen?" efter Tilføj til Outlook (klassisk Outlook). port: kalender-serverens port.
public sealed partial class OutlookFallbackViewModel(ScheduleRef schedule, int port, IPlatform platform, TimeProvider time) : ObservableObject
{
    public string Address => IcsServer.UrlFor(schedule, port);

    [ObservableProperty] public partial string CopyLabel { get; set; } = "Kopiér adresse";

    [RelayCommand]
    async Task CopyAddress()
    {
        CopyLabel = CopyResult.Label(await platform.CopyTextAsync(Address));
        await Task.Delay(CopyResult.Shown, time);
        CopyLabel = "Kopiér adresse";
    }
}

// Knappen viser kort, om kopieringen lykkedes.
static class CopyResult
{
    public static readonly TimeSpan Shown = TimeSpan.FromSeconds(1.2);

    public static string Label(bool copied) => copied ? "✓ Kopieret" : "Kunne ikke kopiere";
}
