using AulaSync.Core;

namespace AulaSync.App;

// Session-levetid (spec §5): er sessionen udløbet, prøves først et stille genlogin med den persistente profil (højst 30 s).
// Lykkes det ikke, viser ikonet "kræver handling" (via status), og notifikationsboksen vises højst én gang, til brugeren er
// logget ind igen. Klik på boksen åbner login (AppHost).
public sealed class Relogin(Func<CancellationToken, Task<bool>> trySilent, INotifier notifier, FileLog log, TimeProvider time, bool silentSupported)
{
    public static readonly TimeSpan Timeout = TimeSpan.FromSeconds(30);
    public const string ExpiredText = "Du er logget ud af Aula. Klik her for at logge ind igen.";
    public const string StartText = "Klik her for at logge ind på Aula, så AulaSync kan holde dine skemaer opdateret.";

    bool _notified;
    bool _running;

    // notification = null: ingen notifikation ved fejl (fx når login-vinduet alligevel vises).
    public async Task<bool> TryAsync(string? notification)
    {
        if (_running) return false;
        _running = true;
        try
        {
            if (silentSupported)
            {
                using var cts = new CancellationTokenSource(Timeout, time);
                if (await trySilent(cts.Token))
                {
                    log.Info("Stille genlogin lykkedes");
                    _notified = false;
                    return true;
                }
                log.Info("Stille genlogin lykkedes ikke");
            }
            if (notification is not null && !_notified)
            {
                notifier.Show(notification);
                _notified = true;
            }
            return false;
        }
        finally { _running = false; }
    }

    // Brugeren loggede selv ind: næste udløb må give en ny notifikation.
    public void SignedIn() => _notified = false;
}
