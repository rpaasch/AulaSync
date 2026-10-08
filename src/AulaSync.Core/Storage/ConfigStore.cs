using System.Text.Json;
using System.Text.Json.Serialization;

namespace AulaSync.Core;

public enum CalendarApp { AppleCalendar, OutlookClassic, OutlookImport, Other }

// Port: kalender-serverens port for denne bruger; den vælges ved første start (UserPort).
public sealed record AppConfig(CalendarApp? CalendarApp = null, bool FirstRunDone = false, int? Port = null)
{
    [JsonIgnore] public int ServerPort => Port ?? IcsServer.DefaultPort;
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
