namespace AulaSync.Core;

// Kalender-serverens port gælder for brugeren og står i config.json. Første gang vælges 9876 eller den første ledige op til
// 9899, så en anden bruger på samme computer (hurtigt brugerskift) eller en kørende AulaSync 2 ikke spærrer porten. Derefter
// bruges kun den gemte port, for kalenderprogrammerne abonnerer på adresser med netop den.
public static class UserPort
{
    public const int First = IcsServer.DefaultPort;
    public const int Last = 9899;

    // tryStart: start serveren på porten, eller null, hvis porten er optaget.
    public static T? Start<T>(ConfigStore config, Func<int, T?> tryStart) where T : class
    {
        if (config.Load().Port is { } saved) return tryStart(saved);
        for (var port = First; port <= Last; port++)
        {
            if (tryStart(port) is not { } server) continue;
            config.Save(config.Load() with { Port = port });
            return server;
        }
        return null;
    }
}
