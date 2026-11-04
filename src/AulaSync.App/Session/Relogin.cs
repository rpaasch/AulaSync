using AulaSync.Core;

namespace AulaSync.App;

// Session-levetid (spec §5): er sessionen udløbet, prøves først et stille genlogin med den persistente profil (højst 30 s).
// Lykkes det ikke, viser ikonet "kræver handling" (via status), og notifikationsboksen vises, som Indstillinger › Besked, når
// du er logget ud af Aula siger (config, AppConfig.LoggedOutNotice): hver dag, til brugeren er logget ind igen, kun én gang,
// eller aldrig. Hvornår den sidst blev vist, gemmes (LoggedOutNotifiedAt), så det også gælder efter en genstart. Mens
// AulaSync er logget ud, prøves et stille genlogin igen en gang om dagen (RemindAsync), uanset indstillingen. Uden config:
// én gang pr. kørsel. Klik på boksen åbner login (AppHost).
public sealed class Relogin(Func<CancellationToken, Task<bool>> trySilent, INotifier notifier, FileLog log, TimeProvider time,
    bool silentSupported, ConfigStore? config = null)
{
    public static readonly TimeSpan Timeout = TimeSpan.FromSeconds(30);
    // En ny dag begynder her: et nyt forsøg og (med Hver dag) en ny besked tidligst dette klokkeslæt, ikke midt om natten.
    public static readonly TimeOnly DailyFrom = new(7, 0);
    public const string ExpiredText = "Du er logget ud af Aula. Klik her for at logge ind igen.";
    public const string StartText = "Klik her for at logge ind på Aula, så AulaSync kan holde dine skemaer opdateret.";

    DateTimeOffset? _triedAt;     // seneste forsøg i denne kørsel, der ikke lykkedes
    DateTimeOffset? _notifiedAt;  // uden config
    bool _running;

    // Et forsøg er i gang (et andet kald afgør, om AulaSync er logget ud).
    public bool Running => _running;

    LoggedOutNotice Notice => config?.Load().LoggedOutNotice ?? LoggedOutNotice.Once;

    DateTimeOffset? NotifiedAt
    {
        get => config is null ? _notifiedAt : config.Load().LoggedOutNotifiedAt;
        set
        {
            if (config is null) { _notifiedAt = value; return; }
            try
            {
                var saved = config.Load();
                if (saved.LoggedOutNotifiedAt != value) config.Save(saved with { LoggedOutNotifiedAt = value });
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { log.Error("Kunne ikke gemme tidspunktet for beskeden", ex); }
        }
    }

    bool Due() => Notice switch
    {
        LoggedOutNotice.Never => false,
        LoggedOutNotice.Daily => NotifiedAt is not { } at || NewDay(at),
        _ => NotifiedAt is null,
    };

    // En ny dag efter since, og klokken er mindst DailyFrom. Et tidspunkt før DailyFrom hører til dagen før, så en besked om
    // natten ikke udskyder morgenens.
    bool NewDay(DateTimeOffset since)
    {
        var now = TimeZoneInfo.ConvertTime(time.GetUtcNow(), time.LocalTimeZone);
        var then = TimeZoneInfo.ConvertTime(since, time.LocalTimeZone);
        var day = TimeOnly.FromDateTime(then.DateTime) < DailyFrom ? then.Date.AddDays(-1) : then.Date;
        return now.Date > day && TimeOnly.FromDateTime(now.DateTime) >= DailyFrom;
    }

    // AppHost hver time, mens AulaSync er logget ud: en ny dag prøves et stille genlogin igen, og lykkes det ikke, kommer en
    // besked, hvis indstillingen siger det (med Hver dag en ny en ny dag).
    public Task<bool> RemindAsync(string notification) =>
        _triedAt is { } at && !NewDay(at) ? Task.FromResult(false) : TryAsync(notification);

    // notification = null: ingen notifikation ved fejl (fx når login-vinduet alligevel vises).
    public async Task<bool> TryAsync(string? notification)
    {
        if (_running) return false;
        _running = true;
        try
        {
            _triedAt = time.GetUtcNow();
            if (silentSupported)
            {
                using var cts = new CancellationTokenSource(Timeout, time);
                if (await trySilent(cts.Token))
                {
                    log.Info("Stille genlogin lykkedes");
                    SignedIn();
                    return true;
                }
                log.Info("Stille genlogin lykkedes ikke");
            }
            if (notification is not null && Due())
            {
                notifier.Show(notification);
                NotifiedAt = time.GetUtcNow();
            }
            return false;
        }
        finally { _running = false; }
    }

    // Brugeren loggede ind: næste udløb må give en ny notifikation.
    public void SignedIn()
    {
        _triedAt = null;
        NotifiedAt = null;
    }
}
