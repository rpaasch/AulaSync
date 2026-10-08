using System.Text.Json;
using System.Text.Json.Serialization;

namespace AulaSync.Core;

public enum CalendarApp { AppleCalendar, OutlookClassic, OutlookImport, Other }

// Port: kalender-serverens port for denne bruger; den vælges ved første start (UserPort).
// UpdateMinutes: hvor tit skemaerne hentes fra Aula (Indstillinger); null eller en ukendt værdi er hver 4. time.
// StatusEvent: en privat aftale mandag kl. 5.45 i hver kalender viser, hvornår skemaet sidst blev hentet.
public sealed record AppConfig(CalendarApp? CalendarApp = null, bool FirstRunDone = false, int? Port = null,
    int? UpdateMinutes = null, bool StatusEvent = false)
{
    [JsonIgnore] public int ServerPort => Port ?? IcsServer.DefaultPort;
    [JsonIgnore] public TimeSpan UpdateInterval => UpdateIntervals.Of(UpdateMinutes);
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
