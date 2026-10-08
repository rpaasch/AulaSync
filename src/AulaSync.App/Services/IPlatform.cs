namespace AulaSync.App;

// Det, brugerfladen skal bruge fra styresystemet. Testene bruger en falsk udgave.
public interface IPlatform
{
    bool IsMac { get; }

    string FileManager => IsMac ? "Finder" : "Stifinder";

    string TrayPlace => IsMac ? "menulinjen" : "systembakken";

    // Åbn en adresse (webcal://, https://), en fil eller en mappe i det tilknyttede program.
    void Open(string target);

    // Åbn en http-adresse i Apple Kalender, som så tilbyder at abonnere på den. Kalender laver webcal:// om til https og
    // prøver ikke http bagefter; med http prøver den https, får AulaSyncs TLS-afvisning og henter så over http.
    void OpenInCalendar(string url);

    // Vis filen markeret i Finder/Stifinder.
    void RevealFile(string path);

    // false, når teksten ikke kom i udklipsholderen (fx fordi et andet program holder den).
    Task<bool> CopyTextAsync(string text);
}
