using AulaSync.Core;

namespace AulaSync.App;

// Et importeret skema er ændret (spec §3.2): en besked, hvor et klik åbner hovedvinduet med "Importér igen…".
// Kun når kalenderprogrammet importerer; med et abonnement henter programmet selv den nye fil.
public sealed class ImportReminder(SyncService sync, ConfigStore config, IPlatform platform, INotifier notifier,
    Action showMain, Action<Action> ui)
{
    public void Start() => sync.ImportsChanged += changed => ui(() => Show(changed));

    // Importerede skemaer, der er ændret, når kalenderprogrammet importerer (til prikken og menuen); ellers ingen.
    public IReadOnlyList<ScheduleRef> Pending() => Importing ? sync.ChangedImports() : [];

    bool Importing => CalendarChoice.Current(config, platform.IsMac) == CalendarApp.OutlookImport;

    void Show(IReadOnlyList<ScheduleRef> changed)
    {
        if (Importing) notifier.Show($"{StatusText.ImportChanged(changed)}. Klik her for at importere igen.", showMain);
    }
}
