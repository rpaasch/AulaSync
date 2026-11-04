using System.Text.Json;
using System.Text.Json.Serialization;

namespace AulaSync.Core;

public enum CalendarApp { AppleCalendar, OutlookClassic, OutlookImport, Other }

// Hvordan opdateringstiden vises (Indstillinger): en aftale hver mandag på et klokkeslæt, en heldagsbegivenhed hver mandag,
// eller nederst i hver lektions beskrivelse. Timed først, så en config.json fra før valget (kun StatusEvent) giver aftalen.
public enum StatusStyle { Timed, AllDay, InDescription }

// Port: kalender-serverens port for denne bruger; den vælges ved første start (UserPort).
// UpdateMinutes: hvor tit skemaerne hentes fra Aula (Indstillinger); null eller en ukendt værdi er hver 4. time.
// StatusEvent: kalenderne viser, hvornår skemaet i kalenderen blev hentet (IcsWriter), på den måde, StatusStyle siger.
// StatusMinutes: aftalens klokkeslæt i minutter efter midnat (Indstillinger); null eller en ugyldig værdi er kl. 5.45.
public sealed record AppConfig(CalendarApp? CalendarApp = null, bool FirstRunDone = false, int? Port = null,
    int? UpdateMinutes = null, bool StatusEvent = false, int? StatusMinutes = null, StatusStyle StatusStyle = StatusStyle.Timed)
{
    public static readonly TimeOnly DefaultStatusTime = new(5, 45);

    [JsonIgnore] public int ServerPort => Port ?? IcsServer.DefaultPort;
    [JsonIgnore] public TimeSpan UpdateInterval => UpdateIntervals.Of(UpdateMinutes);
    [JsonIgnore] public TimeOnly StatusTime =>
        StatusMinutes is >= 0 and < 24 * 60 ? TimeOnly.FromTimeSpan(TimeSpan.FromMinutes(StatusMinutes.Value)) : DefaultStatusTime;
}

public sealed class ConfigStore(string path)
{
    static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
    };

    public AppConfig Load()
    {
        if (!File.Exists(path)) return new AppConfig();
        try { return JsonSerializer.Deserialize<AppConfig>(File.ReadAllText(path), Options) ?? new AppConfig(); }
        catch (JsonException) { return new AppConfig(); }
    }

    public void Save(AppConfig config) => AtomicFile.WriteAllText(path, JsonSerializer.Serialize(config, Options));
}
