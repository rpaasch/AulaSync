using Avalonia.Controls;
using Avalonia.Platform;

namespace AulaSync.App;

public enum TrayState { Normal, Updating, Attention }

// Ikonet i systembakken (Windows) og menulinjen (Mac), tegnet af packaging/icon/make-icons.py. Windows: AulaSyncs eget
// ikon i alle størrelser fra 16 til 64 px, så Windows vælger den rigtige ved hver skalering (et enkelt billede blev
// skaleret ned og sløret). Mac: samme kalenderblad i sort som skabelon-ikon, som macOS farver efter menulinjen.
// "Opdaterer" = prik midt i bladet, "kræver handling" = prik i hjørnet.
public static class TrayIconImage
{
    static readonly Dictionary<(TrayState, bool), WindowIcon> Cache = new();

    public static Uri Uri(TrayState state, bool template)
    {
        var suffix = state switch { TrayState.Updating => "-updating", TrayState.Attention => "-attention", _ => "" };
        return new($"avares://AulaSync/Assets/Tray/{(template ? "menubar" : "tray")}{suffix}.{(template ? "png" : "ico")}");
    }

    // Ikonet skiftes ved hver statusændring; hvert af de seks læses kun én gang.
    public static WindowIcon Create(TrayState state, bool template)
    {
        lock (Cache)
        {
            if (!Cache.TryGetValue((state, template), out var icon))
            {
                using var stream = AssetLoader.Open(Uri(state, template));
                Cache[(state, template)] = icon = new WindowIcon(stream);
            }
            return icon;
        }
    }
}
